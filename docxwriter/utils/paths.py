"""Path utility functions for DocxWriter."""

from pathlib import Path

# Default output directory
DEFAULT_OUTPUT_DIR = Path("output")


def ensure_dir(directory: Path) -> Path:
    """Ensure a directory exists.

    Args:
        directory: The directory path to ensure exists

    Returns:
        The directory path
    """
    directory.mkdir(exist_ok=True, parents=True)
    return directory


# Create default output directory
output_dir = ensure_dir(Path("data"))
