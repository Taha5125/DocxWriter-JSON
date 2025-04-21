"""
DocxWriter - A Python library for creating Word documents from JSON data.

This package provides tools for creating professionally formatted Word documents
from JSON data, with support for various document elements like headings,
paragraphs, tables, lists, images and more.
"""

__version__ = "0.1.0"

# Import main classes and functions for easier access
from docxwriter.document_creator import DocumentCreator
from docxwriter.components.styles import (
    COLORS,
    FONTS,
    SIZES,
    LINE_SPACING,
    WATERMARK_TEXT,
)

# For backward compatibility
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
