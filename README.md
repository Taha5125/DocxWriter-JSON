# DocxWriter

A powerful Python library for creating professionally formatted Word documents from JSON data. This tool allows you to generate complex documents with various elements like headings, paragraphs, tables, lists, images, and more.

## Features

- Create professionally formatted Word documents
- Support for multiple document elements
- Consistent styling and formatting
- Customizable styles and themes
- Hidden watermark support
- Error handling and logging
- Type hints for better code maintainability

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/karar-hayder/DocxWriter-JSON.git
    cd DocxWriter
    ```

2. Create and activate a virtual environment:

    ```bash
    python -m venv .venv
    source .venv/bin/activate  # On Windows: .venv\Scripts\activate
    ```

3. Install dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## Project Structure

DocxWriter has been restructured into a modular package for better maintainability and extensibility:

```txt
DocxWriter/
├── docxwriter/              # Main package directory
│   ├── __init__.py          # Package initialization
│   ├── __main__.py          # Entry point for command-line usage
│   ├── compat.py            # Compatibility layer for backward compatibility
│   ├── document_creator.py  # Main document creation class
│   ├── components/          # Document components
│   │   ├── __init__.py
│   │   └── styles.py        # Document styles
│   └── utils/               # Utility modules
│       ├── __init__.py
│       ├── logger.py        # Logging utilities
│       └── paths.py         # Path utilities
├── writer.py                # Backward compatibility wrapper
├── styles.py                # Backward compatibility wrapper
├── tests.py                 # Test suite
├── setup.py                 # Package setup script
├── requirements.txt         # Dependencies
├── examples/                # Example data files
├── README.md                # This file
└── CONTRIBUTING.md          # Contribution guidelines
```

## Usage Options

DocxWriter can be used in several ways:

1. **Direct script execution (backward compatible):**

   ```python
   python writer.py
   ```

2. **As a module:**

   ```python
   python -m docxwriter -i data.json -o output
   ```

3. **As a library in your code:**

   ```python
   from docxwriter.document_creator import DocumentCreator
   
   # Create a document creator
   creator = DocumentCreator(output_dir="output")
   
   # Create a document from JSON file
   creator.create_document_from_json("data.json")
   ```

## Command Line Options

When used as a module, DocxWriter supports several command line options:

```txt
usage: python -m docxwriter [-h] [-i INPUT] [-o OUTPUT_DIR] [-w WATERMARK] [--no-watermark]

Create Word documents from JSON data.

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Input JSON file (default: data.json)
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        Output directory (default: data)
  -w WATERMARK, --watermark WATERMARK
                        Custom watermark text (default: predefined text)
  --no-watermark        Disable watermark
```

## Usage

### Basic Structure

Create a `data.json` file with your document content. The basic structure is:

```json
{
    "title": "Document Title",
    "file_name": "output.docx",
    "content": {
        "Section 1": "Regular paragraph text",
        "Section 2": {
            "table": [
                ["Header 1", "Header 2"],
                ["Data 1", "Data 2"]
            ]
        }
    }
}
```

### Document Elements

#### 1. Title

The document title is specified in the root of the JSON:

```json
{
    "title": "Your Document Title",
    "file_name": "output.docx",
    "content": { ... }
}
```

#### 2. Sections

Regular sections with headings and text:

```json
{
    "Section Name": "Section content text. This can be multiple paragraphs separated by double newlines."
}
```

#### 3. Tables

Create tables with headers and data:

```json
{
    "Table Section": {
        "table": [
            ["Header 1", "Header 2", "Header 3"],
            ["Row 1 Col 1", "Row 1 Col 2", "Row 1 Col 3"],
            ["Row 2 Col 1", "Row 2 Col 2", "Row 2 Col 3"]
        ],
        "style": "Table Grid",
        "header_rows": 1
    }
}
```

#### 4. Lists

Create bulleted or numbered lists:

```json
{
    "List Section": {
        "list": [
            "First item",
            "Second item",
            "Third item"
        ],
        "list_type": "bullet",  // or "number"
        "level": 0  // 0 for top level, 1 for nested, etc.
    }
}
```

#### 5. Images

Add images to your document:

```json
{
    "Image Section": {
        "image": "path/to/your/image.jpg",
        "width": 6.0,  // in inches
        "height": 4.0  // in inches
    }
}
```

#### 6. Page Breaks

Add page breaks:

```json
{
    "Page Break Section": {
        "page_break": true
    }
}
```

#### 7. Styled Text

Use different styles for text:

```json
{
    "Warning Section": {
        "style": "Warning",
        "text": "This is a warning message"
    },
    "Error Section": {
        "style": "Error",
        "text": "This is an error message"
    },
    "Success Section": {
        "style": "Success",
        "text": "This is a success message"
    },
    "Code Section": {
        "style": "Code",
        "text": "This is code text"
    }
}
```

### Available Styles

1. Text Styles:
   - Normal
   - Title
   - Subtitle
   - Heading 1-3
   - Abstract
   - Quote
   - Code
   - Warning
   - Error
   - Success

2. List Styles:
   - List Bullet
   - List Number

3. Table Styles:
   - Table Grid

4. Special Styles:
   - Caption
   - Footnote
   - Header
   - Footer
   - Hidden (for watermark)

### Example JSON

Here's a complete example showing various elements:

```json
{
    "title": "Sample Document",
    "file_name": "sample.docx",
    "content": {
        "Introduction": "This is the introduction section with regular paragraph text.",
        
        "Data Table": {
            "table": [
                ["Name", "Age", "City"],
                ["John", "25", "New York"],
                ["Jane", "30", "Los Angeles"]
            ],
            "style": "Table Grid",
            "header_rows": 1
        },
        
        "Features": {
            "list": [
                "Feature 1",
                "Feature 2",
                "Feature 3"
            ],
            "list_type": "bullet"
        },
        
        "Important Notice": {
            "style": "Warning",
            "text": "This is an important warning message!"
        },
        
        "Code Example": {
            "style": "Code",
            "text": "def hello_world():\n    print('Hello, World!')"
        },
        
        "New Page": {
            "page_break": true
        },
        
        "Image Section": {
            "image": "path/to/image.jpg",
            "width": 6.0,
            "height": 4.0
        }
    }
}
```

## Running the Script

Simply run the script with Python:

```bash
python writer.py
```

The script will:

1. Read the data.json file
2. Create a Word document with the specified content
3. Apply all formatting and styles
4. Add a watermark
5. Save the document in the `data` directory

## Output

The generated document will be saved in the `data` directory with the filename specified in your JSON file. The document will include:

- Professional formatting
- Consistent styling
- All specified elements (tables, lists, images, etc.)
- A hidden watermark
- Proper spacing and margins

## Error Handling

The script includes comprehensive error handling for:

- Missing JSON file
- Invalid JSON format
- Missing required fields
- Image loading errors
- Style application errors

All errors are logged with timestamps and appropriate error messages.

## Contributing

Feel free to submit issues and enhancement requests!
