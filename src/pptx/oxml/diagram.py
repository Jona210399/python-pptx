"""lxml custom element classes for diagram (SmartArt) XML elements."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

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


class CT_PtList(BaseOxmlElement):
    """`dgm:ptLst` element containing the list of diagram nodes."""

    pt_lst: list[CT_Pt] = ZeroOrMore("dgm:pt", successors=())  # pyright: ignore


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


def qn(tag: str) -> str:
    """Quick qualified name helper."""
    from pptx.oxml.ns import qn as _qn
    return _qn(tag)
