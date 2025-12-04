"""lxml custom element classes for diagram (SmartArt) XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from lxml import etree

from pptx.oxml.ns import qn
from pptx.oxml.simpletypes import XsdString
from pptx.oxml.xmlchemy import BaseOxmlElement, OptionalAttribute, ZeroOrMore

if TYPE_CHECKING:
    from typing import List


class CT_DataModel(BaseOxmlElement):
    """`dgm:dataModel` root element of diagramData part.

    Contains the node structure of a SmartArt diagram.
    """

    ptLst: CT_PtList | None = None  # Will be accessed via xpath

    @property
    def pt_lst(self) -> List[CT_Pt]:
        """List of `dgm:pt` (point/node) elements in this data model."""
        ptLst = self.find(qn("dgm:ptLst"))
        if ptLst is None:
            return []
        return cast("List[CT_Pt]", ptLst.findall(qn("dgm:pt")))

    @property
    def cxn_lst(self) -> CT_CxnList | None:
        """Get the `dgm:cxnLst` element, or None if it doesn't exist."""
        return cast("CT_CxnList | None", self.find(qn("dgm:cxnLst")))

    def get_doc_node(self) -> CT_Pt | None:
        """Get the document root node (type='doc')."""
        for pt in self.pt_lst:
            if pt.type == "doc":
                return pt
        return None

    def get_or_create_pt_list(self) -> CT_PtList:
        """Get or create the point list element."""
        ptLst = self.find(qn("dgm:ptLst"))
        if ptLst is None:
            ptLst = etree.SubElement(self, qn("dgm:ptLst"))  # pyright: ignore
        return cast("CT_PtList", ptLst)

    def get_or_create_cxn_list(self):
        """Get or create the connection list element."""
        cxnLst = self.find(qn("dgm:cxnLst"))
        if cxnLst is None:
            cxnLst = etree.SubElement(self, qn("dgm:cxnLst"))  # pyright: ignore
        return cxnLst


class CT_PtList(BaseOxmlElement):
    """`dgm:ptLst` element containing the list of diagram nodes."""

    pt_lst: list[CT_Pt] = ZeroOrMore("dgm:pt", successors=())  # pyright: ignore

    def add_pt(self, model_id: str, node_type: str = "node") -> CT_Pt:
        """Add a new `dgm:pt` element with the given modelId and type."""
        pt = cast("CT_Pt", etree.SubElement(self, qn("dgm:pt")))  # pyright: ignore
        pt.modelId = model_id
        pt.type = node_type
        return pt

    def remove_pt(self, pt: CT_Pt) -> None:
        """Remove the given `dgm:pt` element from this list."""
        self.remove(pt)


class CT_CxnList(BaseOxmlElement):
    """`dgm:cxnLst` element containing the list of connections."""

    def add_cxn(self, model_id: str, src_id: str, dest_id: str, src_ord: int) -> None:
        """Add a new `dgm:cxn` connection element."""
        cxn = etree.SubElement(self, qn("dgm:cxn"))  # pyright: ignore
        cxn.set("modelId", model_id)
        cxn.set("srcId", src_id)
        cxn.set("destId", dest_id)
        cxn.set("srcOrd", str(src_ord))
        cxn.set("destOrd", "0")

    def remove_cxn_by_dest(self, dest_id: str) -> None:
        """Remove all connections where destId matches."""
        for cxn in self.findall(qn("dgm:cxn")):
            if cxn.get("destId") == dest_id:
                self.remove(cxn)

    def remove_cxn_by_node(self, node_id: str) -> None:
        """Remove all connections where the node is source or destination."""
        for cxn in list(self.findall(qn("dgm:cxn"))):
            if cxn.get("srcId") == node_id or cxn.get("destId") == node_id:
                self.remove(cxn)

    def cleanup_orphaned_connections(self, valid_node_ids: set[str]) -> None:
        """Remove connections that reference non-existent nodes."""
        for cxn in list(self.findall(qn("dgm:cxn"))):
            src_id, dst_id = cxn.get("srcId"), cxn.get("destId")
            if src_id not in valid_node_ids or dst_id not in valid_node_ids:
                self.remove(cxn)


class CT_Pt(BaseOxmlElement):
    """`dgm:pt` element, representing a single node in the diagram.

    Each point/node has a modelId and can contain text and properties.
    """

    modelId: str = OptionalAttribute("modelId", XsdString, default="")  # pyright: ignore
    type: str = OptionalAttribute("type", XsdString, default="node")  # pyright: ignore

    @property
    def text_value(self) -> str:
        """Text content of this diagram node.

        Returns the text from the `dgm:t` element, or empty string if not present.
        """
        t_elem = self.find(qn("dgm:t"))
        if t_elem is not None:
            v_elem = t_elem.find(qn("dgm:v"))
            if v_elem is not None and v_elem.text:
                return v_elem.text
        return ""

    @property
    def connection_id(self) -> str | None:
        """Optional connection ID for this node."""
        return self.get("cxnId")

    @property
    def parent_id(self) -> str | None:
        """Optional parent node ID for this node."""
        prSet = self.find(qn("dgm:prSet"))
        if prSet is not None:
            return prSet.get("phldrT")
        return None

    def _ensure_text_structure(self) -> None:
        """Ensure this point has dgm:t/a:p structure for text content."""
        t_elem = self.find(qn("dgm:t"))
        if t_elem is None:
            # Create dgm:t element
            t_elem = etree.SubElement(self, qn("dgm:t"))
            # Add a:bodyPr (required by schema)
            etree.SubElement(t_elem, qn("a:bodyPr"))
            # Add a:lstStyle (required by schema)
            etree.SubElement(t_elem, qn("a:lstStyle"))
            # Create a:p (paragraph) inside dgm:t
            p_elem = etree.SubElement(t_elem, qn("a:p"))
            # Add a:endParaRPr for default formatting with lang attribute
            endParaRPr = etree.SubElement(p_elem, qn("a:endParaRPr"))
            endParaRPr.set("lang", "en-US")  # Default language
