"""Styles module for backward compatibility.

This module provides a backward-compatible interface to the DocxWriter styles module.
It imports all the necessary functions and variables from the new modular structure
and exports them to maintain compatibility with existing code.
"""

# Import and export all styles-related functions and variables
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

# Import document from compat module
from docxwriter.compat import doc
