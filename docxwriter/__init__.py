"""
DocxWriter - A Python library for creating Word documents from JSON data.
"""

from docxwriter.document_creator import DocumentCreator
from docxwriter.components.styles import (
    COLORS,
    FONTS,
    SIZES,
    LINE_SPACING,
    WATERMARK_TEXT,
)

from docxwriter.compat import (
    doc,
    add_page_break,
    add_image,
    add_hyperlink,
    add_table,
    add_list,
    add_section,
    process_content,
)

__version__ = "0.1.0"
__author__ = "Karar Haider"
__license__ = "MIT"
__url__ = "https://github.com/karar-hayder/docxwriter"
__description__ = "A Python library for creating Word documents from JSON data"
__keywords__ = [
    "docx",
    "word",
    "document",
    "json",
    "writer",
    "create",
    "document",
    "creator",
]
__all__ = [
    "DocumentCreator",
    "COLORS",
    "FONTS",
    "SIZES",
    "LINE_SPACING",
    "WATERMARK_TEXT",
    "doc",
    "add_page_break",
    "add_image",
    "add_hyperlink",
    "add_table",
    "add_list",
    "add_section",
    "process_content",
]
