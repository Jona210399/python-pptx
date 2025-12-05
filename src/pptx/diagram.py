"""SmartArt-related objects.

A SmartArt graphic is a visual representation of information that can be created in PowerPoint.
"""

# pyright: ignore[reportUnknownMemberType, reportAttributeAccessIssue, reportUnknownArgumentType, reportUnknownVariableType]

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast
from uuid import uuid4

from lxml import etree

from pptx.oxml.ns import qn
from pptx.shared import ParentedElementProxy
from pptx.text.text import set_text_preserve_formatting

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DataModel, CT_Pt, CT_PtList
    from pptx.oxml.text import CT_TextParagraph
    from pptx.parts.diagram import DiagramDataPart
    from pptx.types import ProvidesPart


class SmartArtNodeFactory:
    """Factory for creating and removing SmartArt diagram nodes."""

    def __init__(self, data_model: CT_DataModel) -> None:
        self._data_model = data_model

    def create_node_with_transitions(
        self, ptLst: CT_PtList, node_id: str, cxn_id: str, par_trans_id: str, sib_trans_id: str
    ) -> CT_Pt:
        """Create a data node and its associated transition nodes."""
        pt = self._create_data_node(ptLst, node_id)
        self._create_transition_nodes(ptLst, par_trans_id, sib_trans_id, cxn_id)
        return pt

    def create_presentation_nodes(
        self, ptLst: CT_PtList, data_node_id: str, template_data_node_id: str | None = None
    ) -> None:
        """Create presentation nodes for a data node.

        Copies the presentation node structure from an existing template node.
        If no template is provided, no presentation nodes are created (allows
        SmartArt diagrams without presentation nodes to work correctly).

        Args:
            ptLst: The point list to add nodes to
            data_node_id: The model ID of the data node to create pres nodes for
            template_data_node_id: Optional model ID of an existing data node to
                                  copy presentation node structure from
        """
        import uuid

        # Find template presentation nodes if provided
        template_pres_nodes = []
        if template_data_node_id:
            for pt in ptLst.findall(qn("dgm:pt")):
                if pt.get("type") == "pres":
                    prSet = pt.find(qn("dgm:prSet"))
                    if prSet is not None and prSet.get("presAssocID") == template_data_node_id:
                        template_pres_nodes.append(pt)

        # If no template nodes found, don't create presentation nodes
        # This allows SmartArt without presentation nodes to work
        if not template_pres_nodes:
            return

        # Sort by presName to ensure consistent order
        template_pres_nodes.sort(key=lambda pt: pt.find(qn("dgm:prSet")).get("presName", ""))

        # Count existing editable nodes (excluding the one we're creating for)
        # This determines presStyleIdx for the new nodes
        editable_node_count = 0
        for pt in ptLst.findall(qn("dgm:pt")):
            pt_type = pt.get("type")
            if (
                pt_type not in ("pres", "parTrans", "sibTrans", "doc")
                and pt.get("modelId") != data_node_id
            ):
                editable_node_count += 1

        # presStyleCnt is the total count including the new node
        pres_style_cnt = editable_node_count + 1
        # presStyleIdx is 0-based, so it's the count of existing nodes
        pres_style_idx = editable_node_count

        # Create presentation nodes by copying from template
        created_pres_ids = []
        textRect_pres_id = None  # Track textRect for presOf connection

        for template_pt in template_pres_nodes:
            # Generate unique ID for this pres node
            pres_id = "{" + str(uuid.uuid4()).upper() + "}"
            created_pres_ids.append(pres_id)

            # Create the presentation node
            pres_pt = ptLst.add_pt(pres_id, "pres")
            pres_prSet = etree.SubElement(pres_pt, qn("dgm:prSet"))  # pyright: ignore[reportUnknownMemberType]

            # Set presAssocID to link to new data node
            pres_prSet.set("presAssocID", data_node_id)

            # Copy all attributes from template except presAssocID, presStyleIdx, presStyleCnt
            template_prSet = template_pt.find(qn("dgm:prSet"))
            pres_name = None
            if template_prSet is not None:
                for attr_name, attr_value in template_prSet.attrib.items():
                    clean_name = attr_name.split("}")[-1] if "}" in attr_name else attr_name
                    # Skip presAssocID, presStyleIdx, presStyleCnt as we'll set them explicitly
                    if clean_name not in ("presAssocID", "presStyleIdx", "presStyleCnt"):
                        pres_prSet.set(clean_name, attr_value)
                    if clean_name == "presName":
                        pres_name = attr_value

                # Set presStyleIdx and presStyleCnt with calculated values
                # Special handling: compNode keeps presStyleCnt from template (usually "0")
                # Other nodes (pictRect, textRect) use the incremented count
                template_pres_style_cnt = template_prSet.get("presStyleCnt")
                if template_pres_style_cnt is not None:
                    if pres_name == "compNode":
                        # compNode keeps its original presStyleCnt (usually "0")
                        pres_prSet.set("presStyleCnt", template_pres_style_cnt)
                    else:
                        # Other nodes use the calculated presStyleCnt
                        pres_prSet.set("presStyleCnt", str(pres_style_cnt))

                # presStyleIdx is always calculated (for nodes that have it)
                if template_prSet.get("presStyleIdx") is not None:
                    pres_prSet.set("presStyleIdx", str(pres_style_idx))

                # Track textRect for presOf connection
                if pres_name == "textRect":
                    textRect_pres_id = pres_id

                # Copy child elements deeply
                from copy import deepcopy

                for child in template_prSet:
                    pres_prSet.append(deepcopy(child))

            # Copy spPr element structure from template
            template_spPr = template_pt.find(qn("dgm:spPr"))
            if template_spPr is not None:
                from copy import deepcopy

                pres_pt.append(deepcopy(template_spPr))
            else:
                # Add empty spPr if template doesn't have one
                etree.SubElement(pres_pt, qn("dgm:spPr"))  # pyright: ignore[reportUnknownMemberType]

        # Create presentation connections by copying from template
        self._create_presentation_connections_from_template(
            created_pres_ids, data_node_id, template_pres_nodes
        )

        # Create presOf connection from data node to textRect
        if textRect_pres_id:
            self._create_presof_connection(data_node_id, textRect_pres_id)

    def _create_presentation_connections_from_template(
        self, pres_ids: list[str], data_node_id: str, template_pres_nodes: list
    ) -> None:
        """Create connections for presentation nodes by copying template pattern.

        Finds all connections involving the template nodes and creates
        equivalent connections for the new nodes. Does NOT create presOf connections
        as those are handled separately.
        """
        if not template_pres_nodes or len(pres_ids) != len(template_pres_nodes):
            return

        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))
        if cxnLst_elem is None:
            return

        # Get template node IDs
        template_node_ids = [pt.get("modelId") for pt in template_pres_nodes]

        # Find all connections involving template nodes (EXCEPT presOf)
        template_connections = []
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):
            src_id = cxn.get("srcId")
            dest_id = cxn.get("destId")
            cxn_type = cxn.get("type")

            # Skip presOf connections - they are handled separately
            if cxn_type == "presOf":
                continue

            # Check if this connection involves any template nodes
            if src_id in template_node_ids or dest_id in template_node_ids:
                template_connections.append((cxn_type, src_id, dest_id, cxn))

        # Create new connections by copying the pattern
        import uuid

        for cxn_type, src_id, dest_id, template_cxn in template_connections:
            new_cxn = etree.SubElement(cxnLst_elem, qn("dgm:cxn"))  # pyright: ignore[reportUnknownMemberType]
            new_cxn.set("modelId", "{" + str(uuid.uuid4()).upper() + "}")

            # Map template node IDs to new node IDs
            new_src_id = src_id
            new_dest_id = dest_id

            if src_id in template_node_ids:
                idx = template_node_ids.index(src_id)
                new_src_id = pres_ids[idx]
            if dest_id in template_node_ids:
                idx = template_node_ids.index(dest_id)
                new_dest_id = pres_ids[idx]

            # Copy all attributes from template except modelId, srcId, destId
            for attr_name, attr_value in template_cxn.attrib.items():
                clean_name = attr_name.split("}")[-1] if "}" in attr_name else attr_name
                if clean_name not in ("modelId",):
                    new_cxn.set(clean_name, attr_value)

            # Set the new source and destination
            new_cxn.set("srcId", new_src_id)
            new_cxn.set("destId", new_dest_id)

    def _create_presof_connection(self, data_node_id: str, textRect_pres_id: str) -> None:
        """Create a presOf connection from data node to textRect presentation node.

        This connection is needed for proper text rendering in SmartArt.
        PowerPoint uses srcOrd="0" destOrd="0" for presOf connections.
        """
        import uuid

        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))
        if cxnLst_elem is None:
            return

        # Find the presId from existing presOf connections
        pres_id = None
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):
            if cxn.get("type") == "presOf":
                pres_id = cxn.get("presId")
                break

        # If no presId found, don't create the connection
        if not pres_id:
            return

        # Create presOf connection with srcOrd="0" destOrd="0" (PowerPoint standard)
        new_cxn = etree.SubElement(cxnLst_elem, qn("dgm:cxn"))  # pyright: ignore[reportUnknownMemberType]
        new_cxn.set("modelId", "{" + str(uuid.uuid4()).upper() + "}")
        new_cxn.set("type", "presOf")
        new_cxn.set("srcId", data_node_id)
        new_cxn.set("destId", textRect_pres_id)
        new_cxn.set("srcOrd", "0")
        new_cxn.set("destOrd", "0")
        new_cxn.set("presId", pres_id)

    def _find_parent_presentation_node_and_pres_id(
        self, cxnLst_elem: object
    ) -> tuple[str | None, str | None]:
        """Find the parent presentation node and presId used by existing connections.

        Looks for a pres node that is the source of presParOf connections to
        existing compNodes, and extracts the presId attribute from those connections.

        Returns:
            Tuple of (parent_pres_id, presId) where both can be None if not found
        """
        # Find existing compNode IDs from presentation nodes
        comp_node_ids = set()
        for pt in self._data_model.pt_lst:
            if pt.get("type") == "pres":
                prSet = pt.find(qn("dgm:prSet"))
                if prSet is not None and prSet.get("presName") == "compNode":
                    comp_node_ids.add(pt.modelId)

        # Find which pres node is the parent of these compNodes and get presId
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):  # pyright: ignore[reportAttributeAccessIssue]
            if cxn.get("type") == "presParOf":
                dest_id = cxn.get("destId")
                if dest_id in comp_node_ids:
                    # This is a parent of a compNode
                    return cxn.get("srcId"), cxn.get("presId")

        return None, None

    def _extract_template_pres_id(self, cxnLst_elem: object) -> str | None:
        """Extract presId from existing template presentation connections.

        Looks for presParOf connections to pictRect elements, which indicate
        the layout structure used by template nodes. We should match this presId
        for new nodes to ensure consistency.

        Returns the presId (e.g., "urn:microsoft.com/office/officeart/2005/8/layout/pList1")
        """
        # Find presParOf connections that go to pictRect elements
        # These have the presId we should use for new nodes
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):  # pyright: ignore[reportAttributeAccessIssue]
            if cxn.get("type") == "presParOf" and cxn.get("destId", "").startswith("pict"):
                pres_id = cxn.get("presId")
                if pres_id:
                    return pres_id

        return None

    def remove_node_structure(self, ptLst: CT_PtList, node_id: str) -> None:
        """Remove a node and its associated presentation and transition nodes."""
        self._remove_pres_nodes(node_id, ptLst)
        self._remove_transition_nodes(ptLst)

    def _create_data_node(self, ptLst: CT_PtList, node_id: str) -> CT_Pt:
        """Create a data node with property and shape elements."""
        pt = ptLst.add_pt(node_id, "node")
        # Remove the type attribute - PowerPoint's editable nodes don't have it
        # delattr(pt, 'type') doesn't work with lxml, so delete from attrib dict
        if "type" in pt.attrib:
            del pt.attrib["type"]
        etree.SubElement(pt, qn("dgm:prSet"))  # pyright: ignore[reportUnknownMemberType]
        # Don't set phldrT or phldr attributes - PowerPoint leaves prSet empty
        etree.SubElement(pt, qn("dgm:spPr"))  # pyright: ignore[reportUnknownMemberType]
        pt._ensure_text_structure()  # pyright: ignore[reportPrivateUsage]
        return pt

    def _create_transition_nodes(
        self, ptLst: CT_PtList, par_id: str, sib_id: str, cxn_id: str
    ) -> None:
        """Create parent and sibling transition nodes."""
        for pt_id, pt_type in [(par_id, "parTrans"), (sib_id, "sibTrans")]:
            trans_pt = ptLst.add_pt(pt_id, pt_type)
            trans_pt.set("cxnId", cxn_id)  # pyright: ignore[reportUnknownMemberType]
            etree.SubElement(trans_pt, qn("dgm:prSet"))  # pyright: ignore[reportUnknownMemberType]
            etree.SubElement(trans_pt, qn("dgm:spPr"))  # pyright: ignore[reportUnknownMemberType]
            trans_pt._ensure_text_structure()  # pyright: ignore[reportPrivateUsage]

    def _remove_pres_nodes(self, node_id: str, ptLst: CT_PtList) -> None:
        """Remove presentation nodes associated with the given node."""
        pres_nodes: list[object] = []
        for pt in ptLst.findall(qn("dgm:pt")):
            if pt.get("type") == "pres":
                prSet = pt.find(qn("dgm:prSet"))
                if prSet is not None and prSet.get("presAssocID") == node_id:
                    pres_nodes.append(pt)

        for pres_pt in pres_nodes:
            ptLst.remove_pt(cast("CT_Pt", pres_pt))

    def _remove_transition_nodes(self, ptLst: CT_PtList) -> None:
        """Remove dangling transition nodes that have no valid connections."""
        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))
        if cxnLst_elem is None:
            return

        trans_ids_in_use: set[str] = set()
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):
            par_trans_id = cxn.get("parTransId")
            sib_trans_id = cxn.get("sibTransId")
            if par_trans_id:
                trans_ids_in_use.add(par_trans_id)
            if sib_trans_id:
                trans_ids_in_use.add(sib_trans_id)

        for pt in list(ptLst.findall(qn("dgm:pt"))):
            pt_type = pt.get("type")
            if pt_type in ("parTrans", "sibTrans"):
                pt_id = cast("CT_Pt", pt).modelId
                if pt_id not in trans_ids_in_use:
                    ptLst.remove_pt(cast("CT_Pt", pt))


class SmartArtConnectionManager:
    """Manager for SmartArt diagram connections."""

    def __init__(self, data_model: CT_DataModel) -> None:
        self._data_model = data_model

    def create_connection(
        self,
        cxnLst_elem: object,
        parent_node: CT_Pt,
        node_id: str,
        cxn_id: str,
        par_trans_id: str,
        sib_trans_id: str,
    ) -> None:
        """Create a connection from parent to new node."""
        # Collect existing PARENT-CHILD connections from this parent (type='' or None)
        existing_cxns = []
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):  # pyright: ignore[reportAttributeAccessIssue]
            if cxn.get("srcId") == parent_node.modelId and cxn.get("type") in (None, ""):
                existing_cxns.append(cxn)

        # Create new connection with srcOrd = number of existing children
        cxn_elem = etree.SubElement(cxnLst_elem, qn("dgm:cxn"))
        cxn_elem.set("modelId", cxn_id)
        # Parent-child connections don't have a type attribute (not even empty string)
        cxn_elem.set("srcId", parent_node.modelId)
        cxn_elem.set("destId", node_id)
        cxn_elem.set("srcOrd", str(len(existing_cxns)))  # New node gets next sequential srcOrd
        cxn_elem.set("destOrd", "0")
        cxn_elem.set("parTransId", par_trans_id)
        cxn_elem.set("sibTransId", sib_trans_id)

    def remove_node_connections(self, node_id: str) -> None:
        """Remove connections related to the specified node."""
        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))
        if cxnLst_elem is None:
            return

        for cxn in list(cxnLst_elem.findall(qn("dgm:cxn"))):
            if cxn.get("srcId") == node_id or cxn.get("destId") == node_id:
                cxnLst_elem.remove(cxn)

    def find_parent_from_connections(
        self, cxnLst_elem: object, editable_node_ids: set[str]
    ) -> CT_Pt | None:
        """Find parent node from existing connections."""
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):  # pyright: ignore[reportAttributeAccessIssue]
            dest_id = cxn.get("destId")
            cxn_type = cxn.get("type")
            if dest_id in editable_node_ids and (cxn_type is None or cxn_type == ""):
                src_id = cxn.get("srcId")
                for pt in self._data_model.pt_lst:
                    if pt.modelId == src_id:
                        return pt
        return None

    def cleanup_orphaned_connections(self) -> None:
        """Remove connections that reference non-existent nodes.

        Prevents PowerPoint validation warnings after node removal.
        """
        ptLst_elem = self._data_model.find(qn("dgm:ptLst"))
        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))

        if ptLst_elem is None or cxnLst_elem is None:
            return

        valid_node_ids = {
            pt.get("modelId")
            for pt in ptLst_elem.findall(qn("dgm:pt"))
            if pt.get("modelId") is not None
        }

        for cxn in list(cxnLst_elem.findall(qn("dgm:cxn"))):
            src_id, dst_id = cxn.get("srcId"), cxn.get("destId")
            if src_id not in valid_node_ids or dst_id not in valid_node_ids:
                cxnLst_elem.remove(cxn)

        # Synchronize presOf srcOrd with parent-child srcOrd for data nodes
        # This ensures visual layout order matches data node order
        doc_node_id = None
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            if pt.get("type") == "doc":
                doc_node_id = pt.get("modelId")
                break

        if doc_node_id:
            # Build a map of data node ID -> parent-child srcOrd
            data_node_to_srcord = {}
            for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
                if cxn_elem.get("type") in (None, "") and cxn_elem.get("srcId") == doc_node_id:
                    dest_id = cxn_elem.get("destId")
                    srcOrd = cxn_elem.get("srcOrd")
                    if dest_id and srcOrd:
                        data_node_to_srcord[dest_id] = srcOrd

            # Update presOf srcOrd to match parent-child srcOrd
            for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
                if cxn_elem.get("type") == "presOf":
                    src_id = cxn_elem.get("srcId")
                    if src_id in data_node_to_srcord:
                        cxn_elem.set("srcOrd", data_node_to_srcord[src_id])


class SmartArt(ParentedElementProxy):
    """A SmartArt graphic in a presentation.

    Provides access to the nodes (data points) in the SmartArt diagram.
    """

    part: DiagramDataPart  # pyright: ignore

    def __init__(self, data_model: CT_DataModel, parent: ProvidesPart):
        super().__init__(data_model, parent)
        self._data_model = data_model
        self._node_factory = SmartArtNodeFactory(data_model)
        self._connection_manager = SmartArtConnectionManager(data_model)

    @property
    def nodes(self) -> SmartArtNodes:
        """Collection of nodes in this SmartArt diagram.

        Returns a |SmartArtNodes| collection providing access to individual nodes.
        """
        return SmartArtNodes(self._data_model, self)

    @property
    def editable_nodes(self) -> tuple[SmartArtNode, ...]:
        """Editable nodes in this SmartArt as a sequence.
        Returns a concrete sequence (tuple) of nodes that are editable (not
        document or assistant nodes).
        """
        return tuple(node for node in self.nodes if node.is_editable)

    def add_node(
        self,
        text: str = "",
        parent: SmartArtNode | str | None = None,
    ) -> SmartArtNode:
        """Add a new editable node to this SmartArt diagram.

        Args:
            text: Optional text content for the new node.
            parent: Parent for the new node. Can be:
                   - None (default): adds as sibling to last editable node
                   - SmartArtNode: adds as child of specified node
                   - "root": explicitly adds as child of document root

        Returns:
            The newly created SmartArtNode.
        """
        # IMPORTANT: Capture parent_node BEFORE creating the new node,
        # because creating the node modifies the node list and affects
        # the resolution logic for parent=None
        cxnLst_elem = self._get_or_create_connection_list()
        parent_node = self._resolve_parent_node(parent, cxnLst_elem)

        ptLst = self._get_or_create_pt_list()
        node_id, cxn_id, par_trans_id, sib_trans_id = self._generate_node_ids()

        pt = self._node_factory.create_node_with_transitions(
            ptLst, node_id, cxn_id, par_trans_id, sib_trans_id
        )

        if parent_node is not None:
            self._connection_manager.create_connection(
                cxnLst_elem, parent_node, node_id, cxn_id, par_trans_id, sib_trans_id
            )

        # Auto-create presentation nodes if this is an image layout
        self._auto_create_presentation_nodes(ptLst, node_id)

        # Synchronize presOf ordering to ensure consistent state
        self.synchronize_presof_ordering()

        node = SmartArtNode(pt, self)
        if text:
            node.text = text

        return node

    def _auto_create_presentation_nodes(self, ptLst: CT_PtList, new_node_id: str) -> None:
        """Automatically create presentation nodes for the new data node.

        Detects if existing nodes have presentation nodes and if so, creates
        matching presentation nodes for the new node by copying the structure.
        """
        # Check if any existing editable node has presentation nodes
        template_node_id = None
        for node in self.editable_nodes:
            if node._element.modelId != new_node_id:
                # Check if this node has any presentation nodes
                for pt in ptLst.findall(qn("dgm:pt")):
                    if pt.get("type") == "pres":
                        prSet = pt.find(qn("dgm:prSet"))
                        if prSet is not None and prSet.get("presAssocID") == node._element.modelId:
                            template_node_id = node._element.modelId
                            break
                if template_node_id:
                    break

        # If we found a template with presentation nodes, create matching ones
        if template_node_id:
            self._node_factory.create_presentation_nodes(ptLst, new_node_id, template_node_id)

    def remove_node(self, node: SmartArtNode | int) -> None:
        """Remove a node from this SmartArt diagram.

        Args:
            node: Either a SmartArtNode instance or an integer index.

        Raises:
            IndexError: If index is out of range.
            ValueError: If the node is not found in this SmartArt.
        """
        ptLst = self._get_pt_list()
        if ptLst is None:
            raise ValueError("No nodes to remove")

        if isinstance(node, int):
            node = self.editable_nodes[node]

        if node._element not in self._data_model.pt_lst:  # pyright: ignore[reportPrivateUsage]
            raise ValueError("Node does not belong to this SmartArt")

        node_id = node._element.modelId  # pyright: ignore[reportPrivateUsage]

        self._connection_manager.remove_node_connections(node_id)
        self._node_factory.remove_node_structure(ptLst, node_id)
        ptLst.remove_pt(node._element)  # pyright: ignore[reportPrivateUsage]

        self._connection_manager.cleanup_orphaned_connections()

    def embed_image_for_node(self, node: SmartArtNode, image_file: str) -> str:
        """Embed an image file for a SmartArt node with an image placeholder.

        Args:
            node: The SmartArtNode with an image placeholder.
            image_file: Path to the image file or file-like object.

        Returns:
            The relationship ID (rId) of the embedded image.

        Raises:
            ValueError: If the node doesn't have an image placeholder.
        """
        if not node.has_image_placeholder:
            raise ValueError(f"Node {node.node_id} doesn't have an image placeholder")

        # Get the presentation package to add the image
        image_part = self.part.package.get_or_add_image_part(image_file)

        # Create a NEW relationship from the diagram part to the image part
        # We must create a unique relationship for each node, even if they use the same image
        # Using _rels._add_relationship directly to bypass get_or_add which would reuse existing
        from pptx.opc.constants import RELATIONSHIP_TYPE as RT

        rId = self.part._rels._add_relationship(RT.IMAGE, image_part)

        # Find the presentation node with presName="pictRect" or similar
        # This is the node that should get the blipFill for the image
        node_id = node.node_id

        # Iterate through all nodes to find pres nodes associated with this data node
        for pt in self._data_model.pt_lst:
            if pt.get("type") == "pres":
                prSet = pt.find(qn("dgm:prSet"))
                if prSet is not None:
                    pres_assoc_id = prSet.get("presAssocID")
                    pres_name = prSet.get("presName", "")

                    # Check if this pres node is associated with our data node
                    # and has a picture-related presName
                    if pres_assoc_id == node_id and "pict" in pres_name.lower():
                        # Add blipFill to this presentation node's spPr
                        spPr = pt.find(qn("dgm:spPr"))
                        if spPr is None:
                            spPr = etree.SubElement(pt, qn("dgm:spPr"))  # pyright: ignore[reportUnknownMemberType]

                        # Clear any existing fill
                        for child in list(spPr):
                            spPr.remove(child)

                        # Create blipFill structure
                        blipFill = etree.SubElement(spPr, qn("a:blipFill"))  # pyright: ignore[reportUnknownMemberType]
                        blip = etree.SubElement(blipFill, qn("a:blip"))  # pyright: ignore[reportUnknownMemberType]
                        blip.set(qn("r:embed"), rId)
                        stretch = etree.SubElement(blipFill, qn("a:stretch"))  # pyright: ignore[reportUnknownMemberType]
                        etree.SubElement(stretch, qn("a:fillRect"))  # pyright: ignore[reportUnknownMemberType]

        return rId

    def synchronize_presof_ordering(self) -> None:
        """Synchronize presOf and presParOf connections srcOrd with parent-child srcOrd.

        This ensures that:
        1. presOf srcOrd (data node → text rect) matches parent-child srcOrd
        2. presParOf srcOrd (parent pres → child compNode) is sequential 0,1,2,...
        3. Orphaned presParOf connections are removed
        4. Unreferenced pres nodes are removed
        5. Orphaned transition nodes are removed

        This controls the visual layout order in PowerPoint.
        Call this after adjusting SmartArt nodes if visual order appears incorrect.
        """
        ptLst_elem = self._data_model.find(qn("dgm:ptLst"))
        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))

        if ptLst_elem is None or cxnLst_elem is None:
            return

        # First pass: update presStyleCnt for all presentation nodes
        self._update_pres_style_cnt(ptLst_elem)

        # Second pass: remove unreferenced nodes that are causing repair issues
        self._remove_unreferenced_pres_nodes(ptLst_elem, cxnLst_elem)

        # Clean up orphaned transition nodes
        self._node_factory._remove_transition_nodes(cast("CT_PtList", ptLst_elem))
        doc_node_id = None
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            if pt.get("type") == "doc":
                doc_node_id = pt.get("modelId")
                break

        if not doc_node_id:
            return

        # Build a map of data node ID -> parent-child srcOrd
        data_node_to_srcord = {}
        for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
            if cxn_elem.get("type") in (None, "") and cxn_elem.get("srcId") == doc_node_id:
                dest_id = cxn_elem.get("destId")
                srcOrd = cxn_elem.get("srcOrd")
                if dest_id and srcOrd:
                    data_node_to_srcord[dest_id] = srcOrd

        # Note: presOf connections should NOT be updated - they always use srcOrd="0" destOrd="0"
        # PowerPoint maintains these as fixed values, not synchronized with parent-child order

        # Fix presParOf srcOrd to be sequential for correct visual sibling order
        # Build map of data node -> compNode pres node ID
        data_to_compnode = {}
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            if pt.get("type") == "pres":
                prSet = pt.find(qn("dgm:prSet"))
                if prSet is not None and prSet.get("presName") == "compNode":
                    pres_id = pt.get("modelId")
                    data_id = prSet.get("presAssocID")
                    if pres_id and data_id:
                        data_to_compnode[data_id] = pres_id

        # Find parent pres node (Name0)
        parent_pres_id = None
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            if pt.get("type") == "pres":
                prSet = pt.find(qn("dgm:prSet"))
                if prSet is not None and prSet.get("presName") == "Name0":
                    parent_pres_id = pt.get("modelId")
                    break

        if parent_pres_id:
            # Collect presParOf connections from parent to children
            parent_children_to_remove = []
            parent_children = []

            for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
                if cxn_elem.get("type") == "presParOf" and cxn_elem.get("srcId") == parent_pres_id:
                    dest_id = cxn_elem.get("destId")

                    # Check if dest is a compNode associated with active data
                    for pt in ptLst_elem.findall(qn("dgm:pt")):
                        if pt.get("modelId") == dest_id:
                            prSet = pt.find(qn("dgm:prSet"))
                            if prSet is not None and prSet.get("presName") == "compNode":
                                data_id = prSet.get("presAssocID")
                                # Only keep if it's associated with active data
                                if data_id in data_node_to_srcord:
                                    data_order = int(data_node_to_srcord[data_id])
                                    parent_children.append((data_order, cxn_elem, data_id))
                                else:
                                    # Orphaned compNode - remove it
                                    parent_children_to_remove.append(cxn_elem)
                            else:
                                # Not a compNode - remove presParOf to it
                                parent_children_to_remove.append(cxn_elem)
                            break

            # Remove all orphaned/invalid presParOf connections
            for cxn_elem in parent_children_to_remove:
                cxnLst_elem.remove(cxn_elem)

            # Renumber presParOf srcOrd sequentially based on data node order
            parent_children.sort(key=lambda x: x[0])
            for new_ord, (_, cxn_elem, _) in enumerate(parent_children):
                cxn_elem.set("srcOrd", str(new_ord))

        # Ensure all connections have unique ID attributes
        self._ensure_connection_ids(cxnLst_elem)

    def _ensure_connection_ids(self, cxnLst_elem) -> None:
        """Ensure all connections have unique id attributes required by PowerPoint."""
        # Get the max existing ID number
        max_id = 0
        for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
            cid = cxn_elem.get("id")
            if cid:
                try:
                    cid_num = int(cid)
                    max_id = max(max_id, cid_num)
                except (ValueError, TypeError):
                    pass

        # Assign IDs to connections that don't have them
        next_id = max_id + 1
        for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
            if not cxn_elem.get("id"):
                cxn_elem.set("id", str(next_id))
                next_id += 1

    def _remove_unreferenced_pres_nodes(self, ptLst_elem, cxnLst_elem) -> None:
        """Remove unreferenced pres nodes that are orphaned and cause PowerPoint repair issues."""
        # Build set of all node IDs referenced in connections (srcId/destId)
        referenced_ids = set()
        for cxn_elem in cxnLst_elem.findall(qn("dgm:cxn")):
            src = cxn_elem.get("srcId")
            dest = cxn_elem.get("destId")
            if src:
                referenced_ids.add(src)
            if dest:
                referenced_ids.add(dest)

        # Find and remove unreferenced pres nodes
        nodes_to_remove = []
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            mid = pt.get("modelId")
            ptype = pt.get("type")
            if mid and mid not in referenced_ids and ptype == "pres":
                nodes_to_remove.append(pt)

        # Remove the orphaned nodes
        for pt in nodes_to_remove:
            ptLst_elem.remove(pt)

    def _update_pres_style_cnt(self, ptLst_elem) -> None:
        """Update presStyleCnt for all presentation nodes based on total data node count.

        After adding or removing nodes, all presentation nodes (except special types) should
        have their presStyleCnt updated to reflect the current total number of data nodes.
        """
        # Count editable data nodes
        data_node_count = 0
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            pt_type = pt.get("type")
            if pt_type not in ("pres", "parTrans", "sibTrans", "doc"):
                data_node_count += 1

        # Update all presentation nodes
        for pt in ptLst_elem.findall(qn("dgm:pt")):
            if pt.get("type") == "pres":
                prSet = pt.find(qn("dgm:prSet"))
                if prSet is not None:
                    pres_name = prSet.get("presName")
                    # These special presentation nodes keep presStyleCnt="0":
                    # - compNode: composite node
                    # - sibTrans, parTrans: transition layout nodes
                    # - Name0, Name1, etc.: doc/layout nodes
                    # Only update pictRect and textRect nodes
                    if pres_name in ("pictRect", "textRect"):
                        prSet.set("presStyleCnt", str(data_node_count))

    def _get_or_create_pt_list(self) -> CT_PtList:
        """Get or create the point list element."""
        ptLst_elem = self._data_model.find(qn("dgm:ptLst"))
        if ptLst_elem is None:
            ptLst_elem = etree.SubElement(self._data_model, qn("dgm:ptLst"))  # pyright: ignore[reportUnknownMemberType]
        return cast("CT_PtList", ptLst_elem)

    def _get_pt_list(self) -> CT_PtList | None:
        """Get the point list element if it exists."""
        ptLst_elem = self._data_model.find(qn("dgm:ptLst"))
        return cast("CT_PtList", ptLst_elem) if ptLst_elem is not None else None

    def _generate_node_ids(self) -> tuple[str, str, str, str]:
        """Generate four unique node IDs in UUID format."""
        return tuple(f"{{{uuid4()!s}}}".upper() for _ in range(4))  # type: ignore[return-value]

    def _get_or_create_connection_list(self) -> object:
        """Get or create the connection list element."""
        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))
        if cxnLst_elem is None:
            cxnLst_elem = etree.SubElement(self._data_model, qn("dgm:cxnLst"))  # pyright: ignore[reportUnknownMemberType]
        return cxnLst_elem

    def _resolve_parent_node(
        self, parent: SmartArtNode | str | None, cxnLst_elem: object
    ) -> CT_Pt | None:
        """Resolve the parent node based on parent parameter.

        Returns the parent node element that the new node should be connected to.
        When parent is None, finds the parent of the last editable node so the
        new node becomes a sibling of the last node.
        """
        if parent == "root":
            return self._data_model.get_doc_node()
        if isinstance(parent, SmartArtNode):
            return parent._element  # pyright: ignore[reportPrivateUsage]
        if parent is None and self.editable_nodes:
            # Get the last editable node
            last_node = self.editable_nodes[-1]
            last_node_id = last_node._element.modelId  # pyright: ignore[reportPrivateUsage]

            # Find the parent of the last node from connections
            # This makes the new node a sibling of the last node
            for cxn in cxnLst_elem.findall(qn("dgm:cxn")):  # pyright: ignore[reportAttributeAccessIssue]
                dest_id = cxn.get("destId")
                cxn_type = cxn.get("type")
                # Look for parent-child connection where last_node is the child
                if dest_id == last_node_id and (cxn_type is None or cxn_type == ""):
                    src_id = cxn.get("srcId")
                    # Find the parent node element
                    for pt in self._data_model.pt_lst:
                        if pt.modelId == src_id:
                            return pt
        return self._data_model.get_doc_node()


class SmartArtNode:
    """A single node (data point) in a SmartArt diagram.

    Each node can contain text and has properties like an ID and type.
    """

    def __init__(self, pt_element: CT_Pt, parent: SmartArt):
        self._element = pt_element
        self._parent: SmartArt = parent

    @property
    def text(self) -> str:
        """Text content of this node.

        Returns the text string associated with this diagram node, or an empty
        string if the node contains no text.
        """
        return self._element.text_value

    @text.setter
    def text(self, value: str):
        """Set text content preserving formatting.

        Preserves existing run properties (color, font, etc.) from the first run
        or from endParaRPr if no runs exist. Also preserves paragraph properties
        like indentation level and bullet points.
        """
        self._remove_placeholder_flag()

        # Get paragraph element (dgm:t/a:p)
        t_elem = self._element.find(qn("dgm:t"))
        if t_elem is None:
            return

        p_elem = t_elem.find(qn("a:p"))
        if p_elem is None:
            return

        set_text_preserve_formatting(cast("CT_TextParagraph", p_elem), value)

    def _remove_placeholder_flag(self) -> None:
        prSet_elem = self._element.find(qn("dgm:prSet"))
        if prSet_elem is not None and "phldr" in prSet_elem.attrib:
            del prSet_elem.attrib["phldr"]

    @property
    def placeholder_type(self) -> str | None:
        """The placeholder type of this node, e.g., '[Text]' or '[Image]'.

        Returns None if the node has no placeholder attribute (phldrT).
        """
        prSet = self._element.find(qn("dgm:prSet"))
        if prSet is None:
            return None
        return prSet.get("phldrT")

    @property
    def has_image_placeholder(self) -> bool:
        pt_list = self._element.getparent()
        if pt_list is None:
            print("2. has_image_placeholder: no parent ptLst found")
            return False

        data_model = pt_list.getparent()
        if data_model is None:
            print("3. has_image_placeholder: no parent dataModel found")
            return False

        node_id = self._element.get("modelId")

        all_pts = data_model.pt_lst
        for pres_node in all_pts:
            if pres_node.get("type") == "pres":
                prSet = pres_node.find(qn("dgm:prSet"))
                if prSet is not None:
                    assoc_id = prSet.get("presAssocID")
                    if assoc_id == node_id:
                        pres_name = prSet.get("presName", "") or ""
                        pres_name_l = pres_name.lower()
                        image_indicators = ("pict", "image", "img", "pic", "picture", "fgimgplace")
                        for indicator in image_indicators:
                            if indicator in pres_name_l:
                                print("4. has_image_placeholder: found image placeholder")
                                return True

        print("5. has_image_placeholder: no image placeholder found")
        return False

    @property
    def image_path(self) -> str | None:
        """Get the image file path for this node's image placeholder.

        Returns None if no image has been set yet.

        Note: This is a simplified implementation that stores the path reference.
        Full image embedding requires access to the presentation's image parts
        and the slide's shape tree to create/update the actual picture shape.
        """
        prSet = self._element.find(qn("dgm:prSet"))
        if prSet is None:
            return None
        return prSet.get("{http://schemas.python-pptx.org/}imagePath")

    @image_path.setter
    def image_path(self, value: str | None):
        """Set image file path for this node's image placeholder.

        Args:
            value: Path to an image file, or None to clear.

        Raises:
            ValueError: If this node doesn't have an image placeholder.

        Note: This stores the path and attempts to embed the actual image.
        The image is added to the presentation's image parts and linked
        via the SmartArt diagram part relationships.
        """
        if not self.has_image_placeholder:
            raise ValueError(
                f"This node has placeholder type '{self.placeholder_type}', "
                "not '[Image]'. Only nodes with image placeholders can have images set."
            )

        # Don't remove placeholder flag - keep it to preserve text rendering
        # self._remove_placeholder_flag()
        prSet = self._element.find(qn("dgm:prSet"))
        if prSet is None:
            prSet = etree.SubElement(self._element, qn("dgm:prSet"))  # pyright: ignore

        if value is None:
            # Clear the image reference
            if "{http://schemas.python-pptx.org/}imagePath" in prSet.attrib:
                del prSet.attrib["{http://schemas.python-pptx.org/}imagePath"]
            # Also clear the blip relationship if it exists
            if "{http://schemas.python-pptx.org/}blipRId" in prSet.attrib:
                del prSet.attrib["{http://schemas.python-pptx.org/}blipRId"]
        else:
            # Store the path
            prSet.set("{http://schemas.python-pptx.org/}imagePath", value)

            # Attempt to embed the actual image using parent SmartArt
            try:
                self._parent.embed_image_for_node(self, value)
            except Exception as e:
                # If embedding fails, at least the path is stored
                import warnings

                warnings.warn(
                    f"Failed to embed image '{value}': {e}. "
                    "Path has been stored but image won't display in PowerPoint.",
                    UserWarning,
                )

    @property
    def node_id(self) -> str:
        """Unique identifier for this node within the diagram."""
        return self._element.modelId

    @property
    def node_type(self) -> str:
        """Type of this node (e.g., 'node', 'doc', 'asst')."""
        return self._element.type

    @property
    def is_editable(self) -> bool:
        """Whether this node is editable (not a document or assistant node)."""
        return self.node_type == "node"

    def __repr__(self) -> str:
        """Provide helpful representation for debugging."""
        text_preview = self.text[:30] + "..." if len(self.text) > 30 else self.text
        return f"<SmartArtNode id='{self.node_id}' text='{text_preview}'>"


class SmartArtNodes:
    """Collection of SmartArt nodes.

    Provides iteration and indexed access to the nodes in a SmartArt diagram.
    """

    def __init__(self, data_model: CT_DataModel, parent: SmartArt):
        self._data_model = data_model
        self._parent = parent

    def __getitem__(self, index: int) -> SmartArtNode:
        """Enable indexed access to nodes, e.g., `nodes[0]`."""
        pt_elements = self._data_model.pt_lst
        return SmartArtNode(pt_elements[index], self._parent)

    def __iter__(self) -> Iterator[SmartArtNode]:
        """Enable iteration over nodes, e.g., `for node in nodes:`."""
        for pt_element in self._data_model.pt_lst:
            yield SmartArtNode(pt_element, self._parent)

    def __len__(self) -> int:
        """Return the number of nodes in this collection."""
        return len(self._data_model.pt_lst)

    def __repr__(self) -> str:
        """Provide helpful representation for debugging."""
        return f"<SmartArtNodes count={len(self)}>"
