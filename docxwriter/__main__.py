"""Main entry point for DocxWriter package.

This module is executed when the package is run with python -m docxwriter.
"""

import sys
import argparse
from pathlib import Path

from docxwriter.document_creator import DocumentCreator
from docxwriter.components.styles import WATERMARK_TEXT
from docxwriter.utils.logger import logger


def parse_args():
    """Parse command line arguments."""
    parser = argparse.ArgumentParser(
        description="Create Word documents from JSON data."
    )
    parser.add_argument(
        "-i",
        "--input",
        default="data.json",
        help="Input JSON file (default: data.json)",
    )
    parser.add_argument(
        "-o", "--output-dir", default="data", help="Output directory (default: data)"
    )
    parser.add_argument(
        "-w",
        "--watermark",
        default=WATERMARK_TEXT,
        help="Custom watermark text (default: predefined text)",
    )
    parser.add_argument("--no-watermark", action="store_true", help="Disable watermark")
    return parser.parse_args()


def main():
    """Main function executed when the module is run."""
    args = parse_args()

    # Set watermark text (None if no watermark)
    watermark_text = None if args.no_watermark else args.watermark

    # Create output directory
    output_dir = Path(args.output_dir)

    # Create document
    creator = DocumentCreator(output_dir=output_dir, watermark_text=watermark_text)

    try:
        # Create document from JSON
        output_path = creator.create_document_from_json(args.input)
        logger.info(f"Document created successfully: {output_path}")
        return 0
    except FileNotFoundError:
        logger.error(f"Input file not found: {args.input}")
        return 1
    except Exception as e:
        logger.error(f"Error creating document: {str(e)}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
