"""SmartArt-related objects.

A SmartArt graphic is a visual representation of information that can be created in PowerPoint.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator, cast
from uuid import uuid4

from lxml import etree

from pptx.oxml.ns import qn
from pptx.shared import ParentedElementProxy
from pptx.text.text import set_text_preserve_formatting

if TYPE_CHECKING:
    from lxml.etree import _Element

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

    def remove_node_structure(self, ptLst: CT_PtList, node_id: str) -> None:
        """Remove a node and its associated presentation and transition nodes."""
        self._remove_pres_nodes(node_id, ptLst)
        self._remove_transition_nodes(ptLst)

    def _create_data_node(self, ptLst: CT_PtList, node_id: str) -> CT_Pt:
        """Create a data node with property and shape elements."""
        pt = ptLst.add_pt(node_id, "node")
        prSet = etree.SubElement(pt, qn("dgm:prSet"))  # pyright: ignore[reportUnknownMemberType]
        prSet.set("phldrT", "[Text]")
        prSet.set("phldr", "1")
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
        pres_nodes: list[_Element] = []
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
                if cast("CT_Pt", pt).modelId not in trans_ids_in_use:
                    ptLst.remove_pt(cast("CT_Pt", pt))


class SmartArtConnectionManager:
    """Manager for SmartArt diagram connections."""

    def __init__(self, data_model: CT_DataModel) -> None:
        self._data_model = data_model

    def create_connection(
        self,
        cxnLst_elem: _Element,
        parent_node: CT_Pt,
        node_id: str,
        cxn_id: str,
        par_trans_id: str,
        sib_trans_id: str,
    ) -> None:
        """Create a connection from parent to new node."""
        existing_cxns = sum(
            1
            for cxn in cxnLst_elem.findall(qn("dgm:cxn"))
            if cxn.get("srcId") == parent_node.modelId and cxn.get("type") is None
        )
        cxn_elem = etree.SubElement(cxnLst_elem, qn("dgm:cxn"))
        cxn_elem.set("modelId", cxn_id)
        cxn_elem.set("srcId", parent_node.modelId)
        cxn_elem.set("destId", node_id)
        cxn_elem.set("srcOrd", str(existing_cxns))
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
        self, cxnLst_elem: _Element, editable_node_ids: set[str]
    ) -> CT_Pt | None:
        """Find parent node from existing connections."""
        for cxn in cxnLst_elem.findall(qn("dgm:cxn")):
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

    def add_node(self, text: str = "", parent: SmartArtNode | str | None = None) -> SmartArtNode:
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
        ptLst = self._get_or_create_pt_list()
        node_id, cxn_id, par_trans_id, sib_trans_id = self._generate_node_ids()

        pt = self._node_factory.create_node_with_transitions(
            ptLst, node_id, cxn_id, par_trans_id, sib_trans_id
        )

        cxnLst_elem = self._get_or_create_connection_list()
        parent_node = self._resolve_parent_node(parent, cxnLst_elem)

        if parent_node is not None:
            self._connection_manager.create_connection(
                cxnLst_elem, parent_node, node_id, cxn_id, par_trans_id, sib_trans_id
            )

        node = SmartArtNode(pt, self)
        if text:
            node.text = text

        return node

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

    def _get_or_create_connection_list(self) -> _Element:
        """Get or create the connection list element."""
        cxnLst_elem = self._data_model.find(qn("dgm:cxnLst"))
        if cxnLst_elem is None:
            cxnLst_elem = etree.SubElement(self._data_model, qn("dgm:cxnLst"))  # pyright: ignore[reportUnknownMemberType]
        return cxnLst_elem

    def _resolve_parent_node(
        self, parent: SmartArtNode | str | None, cxnLst_elem: _Element
    ) -> CT_Pt | None:
        """Resolve the parent node based on parent parameter."""
        if parent == "root":
            return self._data_model.get_doc_node()
        if isinstance(parent, SmartArtNode):
            return parent._element  # pyright: ignore[reportPrivateUsage]
        if parent is None and self.editable_nodes:
            editable_node_ids = {
                n._element.modelId
                for n in self.editable_nodes  # pyright: ignore[reportPrivateUsage]
            }
            return self._connection_manager.find_parent_from_connections(
                cxnLst_elem, editable_node_ids
            )
        return self._data_model.get_doc_node()


class SmartArtNode:
    """A single node (data point) in a SmartArt diagram.

    Each node can contain text and has properties like an ID and type.
    """

    def __init__(self, pt_element: CT_Pt, parent: SmartArt):
        self._element = pt_element
        self._parent = parent

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
