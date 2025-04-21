"""Styles module for DocxWriter."""

from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH

# Default watermark text
WATERMARK_TEXT: str | None = (
    "\n\n\n\n\nTHIS WAS MADE BY Karar Haider - @kr__4r\n\n\n\n\n"
)

# Define color schemes
COLORS = {
    "primary": RGBColor(0, 0, 0),  # Black
    "secondary": RGBColor(128, 128, 128),  # Gray
    "accent": RGBColor(0, 112, 192),  # Blue
    "highlight": RGBColor(255, 255, 0),  # Yellow
    "error": RGBColor(255, 0, 0),  # Red
    "success": RGBColor(0, 176, 80),  # Green
    "warning": RGBColor(255, 192, 0),  # Orange
    "info": RGBColor(0, 176, 240),  # Light Blue
}

# Define font families
FONTS = {
    "serif": "Times New Roman",
    "sans": "Arial",
    "mono": "Courier New",
    "heading": "Times New Roman",
    "modern": "Calibri",
    "elegant": "Georgia",
    "technical": "Consolas",
}

# Define font sizes
SIZES = {
    "tiny": Pt(8),
    "small": Pt(10),
    "normal": Pt(12),
    "large": Pt(14),
    "huge": Pt(16),
    "title": Pt(18),
    "subtitle": Pt(14),
    "heading1": Pt(16),
    "heading2": Pt(14),
    "heading3": Pt(12),
}

# Define line spacing options
LINE_SPACING = {
    "single": 1.0,
    "one_point_five": 1.5,
    "double": 2.0,
    "triple": 3.0,
}


def create_style(
    doc,
    name: str,
    base_style: str = "Normal",
    style_type: str = WD_STYLE_TYPE.PARAGRAPH,
):
    """Create a new style with the given name and base style."""
    try:
        style = doc.styles.add_style(name, style_type)
        style.base_style = doc.styles[base_style]
        return style
    except ValueError:
        # Style already exists
        return doc.styles[name]


def apply_paragraph_format(style, **kwargs):
    """Apply paragraph formatting to a style."""
    for key, value in kwargs.items():
        setattr(style.paragraph_format, key, value)


def apply_font_format(style, **kwargs):
    """Apply font formatting to a style."""
    for key, value in kwargs.items():
        if key == "color":
            style.font.color.rgb = value
        else:
            setattr(style.font, key, value)


def apply_styles_to_document(doc):
    """Apply all styles to the document."""
    # Set up standard research paper styles
    # Page margins (1 inch all around)
    sections = doc.sections
    for section in sections:
        section.top_margin = Pt(72)  # 1 inch = 72 points
        section.bottom_margin = Pt(72)
        section.left_margin = Pt(72)
        section.right_margin = Pt(72)

    # Title style
    title_style = create_style(doc, "Title")
    apply_font_format(
        title_style, name=FONTS["heading"], size=SIZES["title"], bold=True
    )
    apply_paragraph_format(
        title_style, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(24)
    )

    # Subtitle style
    subtitle_style = create_style(doc, "Subtitle")
    apply_font_format(
        subtitle_style, name=FONTS["heading"], size=SIZES["subtitle"], italic=True
    )
    apply_paragraph_format(
        subtitle_style, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=Pt(24)
    )

    # Heading styles
    for level in range(1, 4):
        heading_style = doc.styles[f"Heading {level}"]
        apply_font_format(
            heading_style,
            name=FONTS["heading"],
            size=SIZES[f"heading{level}"],
            bold=True,
        )
        apply_paragraph_format(heading_style, space_before=Pt(18), space_after=Pt(12))

    # Normal paragraph style
    normal_style = doc.styles["Normal"]
    apply_font_format(normal_style, name=FONTS["serif"], size=SIZES["normal"])
    apply_paragraph_format(
        normal_style,
        line_spacing=LINE_SPACING["double"],
        space_after=Pt(0),
        first_line_indent=Pt(36),
    )

    # Abstract style
    abstract_style = create_style(doc, "Abstract")
    apply_font_format(abstract_style, name=FONTS["serif"], size=SIZES["normal"])
    apply_paragraph_format(
        abstract_style,
        line_spacing=LINE_SPACING["double"],
        space_before=Pt(0),
        space_after=Pt(24),
        first_line_indent=Pt(0),
    )

    # References style
    references_style = create_style(doc, "References")
    apply_font_format(references_style, name=FONTS["serif"], size=SIZES["normal"])
    apply_paragraph_format(
        references_style,
        line_spacing=LINE_SPACING["double"],
        space_after=Pt(12),
        first_line_indent=Pt(-36),
        left_indent=Pt(36),
    )

    # Quote style
    quote_style = create_style(doc, "Quote")
    apply_font_format(
        quote_style, name=FONTS["serif"], size=SIZES["normal"], italic=True
    )
    apply_paragraph_format(
        quote_style,
        left_indent=Pt(36),
        right_indent=Pt(36),
        space_before=Pt(12),
        space_after=Pt(12),
    )

    # List styles
    list_bullet_style = create_style(doc, "List Bullet")
    apply_font_format(list_bullet_style, name=FONTS["serif"], size=SIZES["normal"])
    apply_paragraph_format(
        list_bullet_style, left_indent=Pt(36), first_line_indent=Pt(-18)
    )

    list_number_style = create_style(doc, "List Number")
    apply_font_format(list_number_style, name=FONTS["serif"], size=SIZES["normal"])
    apply_paragraph_format(
        list_number_style, left_indent=Pt(36), first_line_indent=Pt(-18)
    )

    # Table styles
    table_style = create_style(doc, "Table Grid", style_type=WD_STYLE_TYPE.TABLE)
    apply_font_format(table_style, name=FONTS["serif"], size=SIZES["normal"])
    apply_paragraph_format(table_style, alignment=WD_ALIGN_PARAGRAPH.LEFT)

    # Caption style
    caption_style = create_style(doc, "Caption")
    apply_font_format(
        caption_style, name=FONTS["serif"], size=SIZES["small"], italic=True
    )
    apply_paragraph_format(
        caption_style,
        alignment=WD_ALIGN_PARAGRAPH.CENTER,
        space_before=Pt(6),
        space_after=Pt(6),
    )

    # Footnote style
    footnote_style = create_style(doc, "Footnote")
    apply_font_format(footnote_style, name=FONTS["serif"], size=SIZES["tiny"])
    apply_paragraph_format(
        footnote_style, left_indent=Pt(36), first_line_indent=Pt(-18)
    )

    # Header style
    header_style = create_style(doc, "Header")
    apply_font_format(header_style, name=FONTS["serif"], size=SIZES["small"])
    apply_paragraph_format(header_style, alignment=WD_ALIGN_PARAGRAPH.RIGHT)

    # Footer style
    footer_style = create_style(doc, "Footer")
    apply_font_format(footer_style, name=FONTS["serif"], size=SIZES["small"])
    apply_paragraph_format(footer_style, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    # Code style
    code_style = create_style(doc, "Code")
    apply_font_format(code_style, name=FONTS["mono"], size=SIZES["normal"])
    apply_paragraph_format(
        code_style,
        left_indent=Pt(36),
        right_indent=Pt(36),
        space_before=Pt(12),
        space_after=Pt(12),
    )

    # Warning style
    warning_style = create_style(doc, "Warning")
    apply_font_format(warning_style, name=FONTS["serif"], size=SIZES["normal"])
    warning_style.font.color.rgb = COLORS["warning"]
    apply_paragraph_format(
        warning_style,
        left_indent=Pt(36),
        right_indent=Pt(36),
        space_before=Pt(12),
        space_after=Pt(12),
    )

    # Error style
    error_style = create_style(doc, "Error")
    apply_font_format(error_style, name=FONTS["serif"], size=SIZES["normal"])
    error_style.font.color.rgb = COLORS["error"]
    apply_paragraph_format(
        error_style,
        left_indent=Pt(36),
        right_indent=Pt(36),
        space_before=Pt(12),
        space_after=Pt(12),
    )

    # Success style
    success_style = create_style(doc, "Success")
    apply_font_format(success_style, name=FONTS["serif"], size=SIZES["normal"])
    success_style.font.color.rgb = COLORS["success"]
    apply_paragraph_format(
        success_style,
        left_indent=Pt(36),
        right_indent=Pt(36),
        space_before=Pt(12),
        space_after=Pt(12),
    )

    # Hidden style for watermark
    hidden_style = create_style(doc, "Hidden")
    apply_font_format(hidden_style, name=FONTS["heading"], size=SIZES["normal"])
    hidden_style.font.color.rgb = RGBColor(255, 255, 255)  # White color
    apply_paragraph_format(hidden_style, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    return doc
