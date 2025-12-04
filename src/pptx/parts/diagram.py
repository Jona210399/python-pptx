"""Diagram (SmartArt) part objects."""

from __future__ import annotations

from typing import TYPE_CHECKING, cast

from lxml import etree

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
        smartart_obj = SmartArt(self.data_model, self)

        # After the SmartArt object is created and initial fixes are applied,
        # capture the current blob as the "baseline" for future modifications.
        # This ensures that any modifications we make preserve this fixed state.
        if not hasattr(self, "_baseline_blob_set"):
            # Serialize the current state after fixes
            from lxml import etree

            if self._original_blob:
                # Use the original declaration format
                orig_decl_end = self._original_blob.find(b"?>") + 2
                original_decl = self._original_blob[:orig_decl_end]
                original_line_ending = b"\r\n" if b"\r\n" in self._original_blob else b"\n"

                # Serialize the fixed element
                serialized = etree.tostring(self._element, encoding="UTF-8")
                self._baseline_blob = original_decl + original_line_ending + serialized
            else:
                from pptx.opc.oxml import serialize_part_xml

                self._baseline_blob = serialize_part_xml(self._element)

            self._baseline_blob_set = True

        return smartart_obj

    @property
    def blob(self) -> bytes:
        """bytes XML serialization of this part, preserving original formatting.

        For SmartArt diagrams, we preserve the original XML formatting (quote style,
        line endings) from the template to avoid PowerPoint repair warnings.

        After initial access and fixes, we use the baseline to track modifications.
        """
        # Use baseline (fixed state) if available, otherwise fall back to original
        blob_to_use = getattr(self, "_baseline_blob", None) or self._original_blob

        if blob_to_use:
            # Extract the XML declaration from the baseline/original
            orig_decl_end = blob_to_use.find(b"?>") + 2
            original_decl = blob_to_use[:orig_decl_end]

            # Check if original used Windows line endings
            original_line_ending = b"\r\n" if b"\r\n" in blob_to_use else b"\n"

            # Serialize the root element (without adding a declaration)
            # When parsing with lxml, the root element serializes without a declaration
            serialized = etree.tostring(self._element, encoding="UTF-8")

            # Reconstruct with original declaration and line endings
            return original_decl + original_line_ending + serialized
        else:
            # Fallback to default serialization if no original blob
            from pptx.opc.oxml import serialize_part_xml

            return serialize_part_xml(self._element)

    @classmethod
    def load(cls, partname, content_type, package, blob):  # type: ignore
        """Return a |DiagramDataPart| object loaded from `blob`.

        This overrides the base class to ensure proper parsing of diagram XML.
        """
        element = cast("CT_DataModel", parse_xml(blob))
        instance = cls(partname, content_type, package, element)
        instance._original_blob = blob  # Store original for format preservation
        return instance
