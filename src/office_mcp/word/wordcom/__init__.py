"""wordcom — drive a running Microsoft Word instance from Python.

Quick start:

    from office_mcp.word import wordcom as wl

    with wl.attach() as word:
        doc = word.documents.active
        with doc.edit("Update address"):
            doc.bookmarks["Address"].set_text("123 Main St")
"""

from __future__ import annotations

from importlib.metadata import PackageNotFoundError
from importlib.metadata import version as _pkg_version

from . import constants

try:
    __version__ = _pkg_version("office-mcp")
except PackageNotFoundError:  # running from a source tree with no install metadata
    __version__ = "0.0.0+unknown"
from .content._fields import Field, FieldCollection
from .content._hyperlinks import Hyperlink, HyperlinkCollection
from .content._notes import Endnote, EndnoteCollection, Footnote, FootnoteCollection
from .content._tables import Cell, ColumnAnchor, RowAnchor, Table, TableCollection
from .exceptions import (
    AmbiguousMatchError,
    AnchorNotFoundError,
    ComError,
    DocumentNotFoundError,
    EquationError,
    ExcelNotAvailableError,
    ImageSourceError,
    OpError,
    PathNotAllowedError,
    ReplaceVerificationError,
    SnapshotError,
    StyleNotFoundError,
    WordBusyError,
    WordcomError,
    WordNotRunningError,
)
from .model._anchors import (
    Anchor,
    Bookmark,
    BookmarkCollection,
    ChartAnchor,
    ChartCollection,
    ContentControl,
    EndAnchor,
    EquationAnchor,
    EquationCollection,
    Heading,
    HeadingCollection,
    ImageAnchor,
    ImageCollection,
    Paragraph,
    ParagraphCollection,
    RangeAnchor,
    ShapeAnchor,
    ShapeCollection,
    StartAnchor,
    TextBoxCollection,
)
from .model._app import Word, attach, connect
from .model._document import Document, DocumentCollection
from .model._edit import EditScope
from .model._lists import ListCollection
from .model._properties import PropertyCollection
from .model._sections import HeaderFooter, Section, SectionCollection
from .model._selection import Selection, SelectionSnapshot
from .model._styles import Style, StyleCollection
from .model._themes import DocumentTheme
from .model._variables import VariableCollection
from .references._citations import Bibliography, Citation
from .references._index import Index
from .references._sources import Source
from .references._toa import TableOfAuthorities
from .references._toc import TableOfFigures, Toc
from .render._snapshot import Snapshot
from .review._checkpoint import Checkpoint
from .review._comments import Comment, CommentCollection
from .review._linting import Finding
from .review._revisions import Revision, RevisionCollection

__all__ = [
    "AmbiguousMatchError",
    "Anchor",
    "AnchorNotFoundError",
    "Bibliography",
    "Bookmark",
    "BookmarkCollection",
    "Cell",
    "ChartAnchor",
    "ChartCollection",
    "ColumnAnchor",
    "Checkpoint",
    "Citation",
    "ComError",
    "Comment",
    "CommentCollection",
    "ContentControl",
    "Document",
    "DocumentCollection",
    "DocumentNotFoundError",
    "DocumentTheme",
    "EditScope",
    "EndAnchor",
    "Endnote",
    "EndnoteCollection",
    "EquationAnchor",
    "EquationCollection",
    "EquationError",
    "ExcelNotAvailableError",
    "Field",
    "FieldCollection",
    "Finding",
    "Footnote",
    "FootnoteCollection",
    "Heading",
    "HeadingCollection",
    "HeaderFooter",
    "Hyperlink",
    "HyperlinkCollection",
    "ImageAnchor",
    "ImageCollection",
    "ImageSourceError",
    "Index",
    "ListCollection",
    "OpError",
    "Paragraph",
    "ParagraphCollection",
    "PathNotAllowedError",
    "PropertyCollection",
    "RangeAnchor",
    "ReplaceVerificationError",
    "Revision",
    "RevisionCollection",
    "RowAnchor",
    "Section",
    "SectionCollection",
    "Selection",
    "SelectionSnapshot",
    "ShapeAnchor",
    "ShapeCollection",
    "Snapshot",
    "SnapshotError",
    "Source",
    "StartAnchor",
    "Style",
    "StyleCollection",
    "StyleNotFoundError",
    "Table",
    "TableCollection",
    "TableOfAuthorities",
    "TableOfFigures",
    "TextBoxCollection",
    "Toc",
    "VariableCollection",
    "Word",
    "WordBusyError",
    "WordNotRunningError",
    "WordcomError",
    "__version__",
    "attach",
    "connect",
    "constants",
]
