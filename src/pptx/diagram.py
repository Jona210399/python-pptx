"""SmartArt-related objects.

A SmartArt graphic is a visual representation of information that can be created in PowerPoint.
"""

from __future__ import annotations

from typing import TYPE_CHECKING, Iterator

from pptx.shared import ParentedElementProxy

if TYPE_CHECKING:
    from pptx.oxml.diagram import CT_DataModel, CT_Pt
    from pptx.parts.diagram import DiagramDataPart
    from pptx.types import ProvidesPart


class SmartArt(ParentedElementProxy):
    """A SmartArt graphic in a presentation.

    Provides access to the nodes (data points) in the SmartArt diagram.
    """

    part: DiagramDataPart  # pyright: ignore

    def __init__(self, data_model: CT_DataModel, parent: ProvidesPart):
        super().__init__(data_model, parent)
        self._data_model = data_model

    @property
    def nodes(self) -> SmartArtNodes:
        """Collection of nodes in this SmartArt diagram.

        Returns a |SmartArtNodes| collection providing access to individual nodes.
        """
        return SmartArtNodes(self._data_model, self)


class SmartArtNode:
    """A single node (data point) in a SmartArt diagram.

    Each node can contain text and has properties like an ID and type.
    """

    def __init__(self, pt_element: CT_Pt):
        self._element = pt_element

    @property
    def text(self) -> str:
        """Text content of this node.

        Returns the text string associated with this diagram node, or an empty
        string if the node contains no text.
        """
        return self._element.text_value

    @property
    def node_id(self) -> str:
        """Unique identifier for this node within the diagram."""
        return self._element.modelId

    @property
    def node_type(self) -> str:
        """Type of this node (e.g., 'node', 'doc', 'asst')."""
        return self._element.type

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
        return SmartArtNode(pt_elements[index])

    def __iter__(self) -> Iterator[SmartArtNode]:
        """Enable iteration over nodes, e.g., `for node in nodes:`."""
        for pt_element in self._data_model.pt_lst:
            yield SmartArtNode(pt_element)

    def __len__(self) -> int:
        """Return the number of nodes in this collection."""
        return len(self._data_model.pt_lst)

    def __repr__(self) -> str:
        """Provide helpful representation for debugging."""
        return f"<SmartArtNodes count={len(self)}>"
