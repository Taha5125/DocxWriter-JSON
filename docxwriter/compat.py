"""Compatibility module for DocxWriter.

This module provides backward compatibility with the original writer.py code.
It exports all the functions and variables from the original code so that
existing import statements will continue to work.
"""

from pathlib import Path
from typing import Dict, Any, List, Optional

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docxwriter.components.styles import (
    COLORS,
    FONTS,
    SIZES,
    LINE_SPACING,
    WATERMARK_TEXT,
    create_style,
    apply_paragraph_format,
    apply_font_format,
)
from docxwriter.utils.logger import logger
from docxwriter.utils.paths import output_dir
from docxwriter.document_creator import DocumentCreator

# Create a document creator instance for global use
_creator = DocumentCreator(output_dir=output_dir, watermark_text=WATERMARK_TEXT)

# Export the document for backward compatibility
doc = _creator.doc


# Export the original functions with implementation in the DocumentCreator class
def add_page_break() -> None:
    """Add a page break to the document."""
    _creator.add_page_break()


def add_image(image_path: str, width: float = 6.0, height: float = 4.0) -> None:
    """Add an image to the document with specified dimensions."""
    _creator.add_image(image_path, width, height)


def add_hyperlink(paragraph, text: str, url: str) -> None:
    """Add a hyperlink to the document."""
    _creator.add_hyperlink(paragraph, text, url)


def set_cell_border(cell, **kwargs):
    """Set cell border properties."""
    _creator.set_cell_border(cell, **kwargs)


def add_table(
    data: List[List[str]], style: str = "Table Grid", header_rows: int = 1
) -> None:
    """Add a table to the document with the given data."""
    _creator.add_table(data, style, header_rows)


def add_list(items: List[str], list_type: str = "bullet", level: int = 0) -> None:
    """Add a bulleted or numbered list to the document."""
    _creator.add_list(items, list_type, level)


def add_section(heading: str, text: str, level: int = 1, style: str = None) -> None:
    """Add a section with heading and formatted paragraphs to the document."""
    _creator.add_section(heading, text, level, style)


def process_content(content: Dict[str, Any]) -> None:
    """Process the content dictionary and add it to the document."""
    _creator.process_content(content)


def main():
    """Main function to process the JSON data and create the document."""
    try:
        _creator.create_document_from_json("data.json")
    except Exception as e:
        # The create_document_from_json method already logs errors,
        # so we just re-raise them here
        raise


# If running directly, call the main function
if __name__ == "__main__":
    main()
