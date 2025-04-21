"""Logging utilities for DocxWriter."""

import logging

# Configure the logger
logger = logging.getLogger("docxwriter")


def setup_logging(level=logging.INFO):
    """Set up the logger configuration.

    Args:
        level: The logging level to use
    """
    logging.basicConfig(level=level, format="%(asctime)s - %(levelname)s - %(message)s")
    return logger


# Create a default logger
logger = setup_logging()
