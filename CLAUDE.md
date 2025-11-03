# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

MDtoDOCX is a command-line tool that converts Markdown files to Microsoft Word DOCX format. It supports basic Markdown formatting, tables, and images.

## Architecture

The converter uses a two-stage conversion process:

1. **Markdown → HTML**: Uses the `markdown` library with extensions (extra, codehilite, tables, fenced_code) to convert Markdown to HTML
2. **HTML → DOCX**: Custom `HTMLParser` subclass (`MarkdownToDocxConverter`) parses the HTML and builds a DOCX document using `python-docx`

### Key Components

- **md_to_docx.py**: Main script containing:
  - `MarkdownToDocxConverter`: HTML parser that converts parsed elements to DOCX structures
  - `convert_markdown_to_docx()`: Core conversion function
  - CLI interface with argparse

### Supported Features

#### Basic Features
- Headings (H1-H6)
- Text formatting: bold, italic, inline code
- Lists: ordered and unordered (with nesting)
- Code blocks with monospace font
- Links (text + URL in parentheses)
- Tables (with header row styling)
- Images (local files and URLs, centered with optional captions)

#### Advanced Features (New!)
- **Blockquotes**: `> text` - Styled with 'Intense Quote' style
- **Horizontal rules**: `---` - Visual separator lines
- **Strikethrough**: `~~text~~` - Cross out text
- **Superscript**: `text^super^` - For exponents and footnotes
- **Subscript**: `text~sub~` - For chemical formulas
- **Highlighting**: `==text==` - Yellow background highlight
- **Task lists**: `- [ ]` unchecked, `- [x]` checked - With checkbox symbols (☐/☑)
- **Page breaks**: `<!--pagebreak-->`, `\pagebreak`, or `<pagebreak/>` - Force new page

### State Management in Parser

The `MarkdownToDocxConverter` maintains:
- Current paragraph and run objects for text insertion
- Formatting state: bold, italic, code, strikethrough, superscript, subscript, highlight
- List nesting level and counters
- Blockquote state
- Table data collection before rendering
- Image handling with error fallback

## Development Commands

### Setup
```bash
pip install -r requirements.txt
```

### Run Converter
```bash
# Basic usage
python md_to_docx.py input.md

# Specify output file
python md_to_docx.py input.md -o output.docx

# Test with example
python md_to_docx.py example.md
```

### Dependencies
- `python-docx>=1.1.0`: DOCX document creation
- `markdown>=3.5`: Markdown parsing
- `Pillow>=10.0.0`: Image processing

## Adding New Features

### To add support for a new Markdown element:

1. **If the element is not natively supported by Markdown**: Add a preprocessing function to convert custom syntax to HTML
2. Add handling in `handle_starttag()` and `handle_endtag()` methods
3. Update `handle_data()` if special text processing is needed
4. Add state variables to track the element context
5. Create helper methods (like `_handle_image()`, `_create_table()`) for complex elements

### Example 1: Adding Blockquote Support (Native HTML)

```python
# In __init__:
self.in_blockquote = False

# In handle_starttag:
elif tag == 'blockquote':
    self.in_blockquote = True
    self.current_paragraph = self.doc.add_paragraph()
    self.current_paragraph.style = 'Intense Quote'

# In handle_endtag:
elif tag == 'blockquote':
    self.in_blockquote = False
```

### Example 2: Adding Custom Syntax (Requires Preprocessing)

```python
# Add preprocessing function:
def _preprocess_custom_syntax(content):
    """Convert custom syntax to HTML"""
    content = re.sub(r'@@([^@]+)@@', r'<custom>\1</custom>', content)
    return content

# Call in convert_markdown_to_docx():
md_content = _preprocess_custom_syntax(md_content)

# Then handle in parser as normal HTML tags
```

### Preprocessing Functions

The converter uses preprocessing to support non-standard Markdown syntax:

- `_preprocess_special_formatting()`: Handles strikethrough, highlighting, superscript, subscript
- `_preprocess_task_lists()`: Converts `[ ]` and `[x]` to checkbox symbols
- `_preprocess_page_breaks()`: Converts various page break syntaxes to HTML

**Important**: Order matters! Strikethrough (`~~`) must be processed before subscript (`~`).
