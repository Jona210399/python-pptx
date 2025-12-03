"""Diagram (SmartArt) part objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from pptx.diagram import SmartArt
from pptx.opc.package import XmlPart
from pptx.oxml import parse_xml

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DataModel


class DiagramDataPart(XmlPart):
    """A diagram data part.

    Contains the node structure and data for a SmartArt diagram.
    Corresponds to parts having partnames matching ppt/diagrams/data[1-9][0-9]*.xml
    """

    partname_template = "/ppt/diagrams/data%d.xml"

    @property
    def data_model(self) -> CT_DataModel:
        """The |CT_DataModel| root element of this diagram data part."""
        return self._element  # pyright: ignore

    @property
    def smartart(self) -> SmartArt:
        """The |SmartArt| object representing the diagram in this part."""
        return SmartArt(self.data_model, self)

    @classmethod
    def load(cls, partname, content_type, package, blob):  # type: ignore
        """Return a |DiagramDataPart| object loaded from `blob`.

        This overrides the base class to ensure proper parsing of diagram XML.
        """
        element = cast("CT_DataModel", parse_xml(blob))
        return cls(partname, content_type, package, element)
