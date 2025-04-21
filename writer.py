"""Writer module for backward compatibility.

This module provides a backward-compatible interface to the DocxWriter package.
It imports all the necessary functions and variables from the new modular structure
and exports them to maintain compatibility with existing code.
"""

# Import and export all functions and variables from the compatibility module
from docxwriter.compat import (
    doc,
    output_dir,
    WATERMARK_TEXT,
    COLORS,
    FONTS,
    SIZES,
    LINE_SPACING,
    add_page_break,
    add_image,
    add_hyperlink,
    set_cell_border,
    add_table,
    add_list,
    add_section,
    process_content,
    main,
)

# Import logger for backward compatibility
from docxwriter.utils.logger import logger

# If running directly, call the main function
if __name__ == "__main__":
    main()
