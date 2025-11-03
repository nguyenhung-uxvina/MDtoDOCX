# MDtoDOCX

A powerful command-line tool to convert Markdown files to Microsoft Word DOCX format with extensive formatting support.

## Quick Start

**Want to try it immediately?**

```bash
# 1. Install dependencies
pip install python-docx markdown Pillow

# 2. Convert a file
python md_to_docx.py your-file.md

# 3. Done! You'll get your-file.docx
```

**That's it!** The tool works without any complex setup. Just have Python installed and run the script.

## Features

### Basic Formatting
- Headings (H1-H6)
- Text formatting: **bold**, *italic*, `inline code`
- Lists: ordered and unordered with nesting support
- Code blocks with syntax highlighting
- Links with URL display
- Tables with styled headers
- Image embedding (local files and URLs)

### Advanced Features
- **Blockquotes** - Styled quotations
- **Horizontal rules** - Visual section separators
- **Strikethrough text** - ~~deleted content~~
- **Superscript** - For math formulas and footnotes (x^2^)
- **Subscript** - For chemical formulas (H~2~O)
- **Highlighting** - ==Important text== with yellow background
- **Task lists** - Checkboxes for TODO items
- **Page breaks** - Control document pagination

## Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package installer)

### Step-by-Step Installation

1. **Clone or download this repository**
   ```bash
   git clone <repository-url>
   cd MDtoDOCX
   ```

   Or download and extract the ZIP file to your preferred location.

2. **Install required dependencies**
   ```bash
   pip install -r requirements.txt
   ```

   This will install:
   - `python-docx` - For creating DOCX documents
   - `markdown` - For parsing Markdown syntax
   - `Pillow` - For image processing

3. **Verify installation**
   ```bash
   python md_to_docx.py example.md
   ```

   If successful, you'll see:
   ```
   Converting 'example.md'... [OK]
   Output: example.docx
   ```

### Alternative: Install in Virtual Environment (Recommended)

Using a virtual environment keeps dependencies isolated:

```bash
# Create virtual environment
python -m venv venv

# Activate it
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt
```

### Troubleshooting

**"python: command not found"**
- Try `python3` instead of `python`
- Ensure Python is installed and added to your PATH

**"pip: command not found"**
- Try `python -m pip install -r requirements.txt`
- Or install pip: https://pip.pypa.io/en/stable/installation/

**Permission errors on installation**
- Use `pip install --user -r requirements.txt`
- Or run with administrator/sudo privileges (not recommended)

## Usage

### Single File Conversion

Convert a Markdown file to DOCX:
```bash
python md_to_docx.py input.md
```

Specify a custom output filename:
```bash
python md_to_docx.py input.md -o output.docx
```

### Batch Conversion (Multiple Files)

Convert multiple files at once:
```bash
python md_to_docx.py file1.md file2.md file3.md
```

Use wildcards to convert all Markdown files in a directory:
```bash
python md_to_docx.py *.md
```

**Note:** The `-o` option is only available for single file conversion. Batch operations automatically generate output files with the same names.

### Making It Easier to Use

#### Option 1: Create an Alias (Recommended for Regular Use)

**On Windows (PowerShell):**
```powershell
# Add to your PowerShell profile
notepad $PROFILE

# Add this line (update path to your installation):
function md2docx { python "C:\path\to\MDtoDOCX\md_to_docx.py" $args }

# Now you can use:
md2docx document.md
```

**On macOS/Linux (Bash/Zsh):**
```bash
# Add to ~/.bashrc or ~/.zshrc
alias md2docx='python /path/to/MDtoDOCX/md_to_docx.py'

# Reload your shell
source ~/.bashrc  # or ~/.zshrc

# Now you can use:
md2docx document.md
```

#### Option 2: Add to System PATH

**Windows:**
1. Copy `md_to_docx.py` to a folder in your PATH (e.g., `C:\Tools\`)
2. Create a batch file `md2docx.bat` in the same folder:
   ```batch
   @echo off
   python "%~dp0md_to_docx.py" %*
   ```
3. Now run from anywhere: `md2docx document.md`

**macOS/Linux:**
```bash
# Create a symbolic link in a PATH directory
sudo ln -s /path/to/MDtoDOCX/md_to_docx.py /usr/local/bin/md2docx
chmod +x /path/to/MDtoDOCX/md_to_docx.py

# Now run from anywhere
md2docx document.md
```

#### Option 3: Run from Any Directory

Always specify the full path:
```bash
python "C:\path\to\MDtoDOCX\md_to_docx.py" document.md
```

### Common Usage Patterns

**Convert all docs in current folder:**
```bash
python md_to_docx.py *.md
```

**Convert docs in a specific folder:**
```bash
python md_to_docx.py /path/to/docs/*.md
```

**Convert with custom name:**
```bash
python md_to_docx.py draft.md -o "Final Report.docx"
```

**Convert multiple specific files:**
```bash
python md_to_docx.py chapter1.md chapter2.md chapter3.md
```

**Drag and drop (Windows):**
1. Create a batch file `convert.bat`:
   ```batch
   @echo off
   python "C:\path\to\MDtoDOCX\md_to_docx.py" %*
   pause
   ```
2. Drag .md files onto the batch file to convert them

## Supported Markdown Elements

### Text Formatting
| Markdown | Result |
|----------|--------|
| `**bold**` | **bold** |
| `*italic*` | *italic* |
| `` `code` `` | `code` |
| `~~strikethrough~~` | ~~strikethrough~~ |
| `==highlight==` | ==highlight== |
| `x^2^` | superscript |
| `H~2~O` | subscript |

### Structure
- **Headings**: `# H1` through `###### H6`
- **Lists**: Ordered (`1.`) and unordered (`-` or `*`) with nesting
- **Task lists**: `- [ ]` unchecked, `- [x]` checked
- **Blockquotes**: `> quote text`
- **Horizontal rules**: `---` or `***`
- **Page breaks**: `<!--pagebreak-->`, `\pagebreak`, or `<pagebreak/>`

### Media
- **Images**: `![alt text](image.png)` - Supports local and remote URLs
- **Links**: `[text](url)` - Displays text with URL in parentheses

### Tables
```markdown
| Header 1 | Header 2 |
|----------|----------|
| Cell 1   | Cell 2   |
```

### Code
- **Inline code**: `` `code` ``
- **Code blocks**:
  ````markdown
  ```python
  def hello():
      print("Hello!")
  ```
  ````

## Examples

### Quick Start
Try it with the included examples:
```bash
# Basic features
python md_to_docx.py example.md

# All features demo
python md_to_docx.py features_test.md
```

### Real-World Usage
```bash
# Convert a document
python md_to_docx.py my_document.md

# Specify custom output
python md_to_docx.py report.md -o monthly_report.docx

# Convert all markdown files in current directory
python md_to_docx.py *.md

# Convert specific set of files
python md_to_docx.py chapter1.md chapter2.md chapter3.md
```

### Batch Conversion Output
When converting multiple files, you'll see a progress indicator and summary:
```
Converting 'file1.md'... [OK]
Converting 'file2.md'... [OK]
Converting 'file3.md'... [OK]

==================================================
Conversion Summary:
  Successful: 3/3
==================================================
```

## Feature Examples

### Task Lists
```markdown
- [x] Design the converter
- [x] Implement core features
- [x] Add advanced formatting
- [ ] Add custom themes
```

### Scientific Notation
```markdown
Einstein's equation: E = mc^2^
Water formula: H~2~O
```

### Page Breaks
```markdown
Section 1 content here...

<!--pagebreak-->

Section 2 starts on a new page
```

## How It Works

MDtoDOCX uses a two-stage conversion process:

1. **Markdown → HTML**: The `markdown` library parses your .md file and converts it to HTML
2. **HTML → DOCX**: A custom parser reads the HTML and builds a properly formatted Word document using `python-docx`

**What happens during conversion:**
- Text formatting (bold, italic, etc.) is applied to text runs
- Headings use Word's built-in heading styles (H1-H6)
- Lists maintain proper indentation and numbering
- Images are embedded (downloaded if from URLs)
- Tables are created with styled headers
- Special features (strikethrough, highlighting, etc.) are preprocessed before conversion

**File handling:**
- Supports UTF-8, Windows (cp1252), and other common encodings
- Handles files with or without BOM (Byte Order Mark)
- Images can be local files or URLs (automatically downloaded)
- Output files use the same name as input with .docx extension (unless specified with `-o`)

## Tips & Best Practices

### Writing Markdown for Best Results

**✓ DO:**
- Use standard Markdown syntax for maximum compatibility
- Save files in UTF-8 encoding
- Use descriptive alt text for images: `![Chart showing sales data](chart.png)`
- Test with the included `features_test.md` to see all supported features
- Use relative paths for local images when possible

**✗ AVOID:**
- Very large images (will be resized to fit the page)
- Extremely nested lists (stick to 2-3 levels max)
- Special characters in filenames (use alphanumeric and hyphens)
- Opening the DOCX file while conversion is running

### Performance Tips

- Batch conversion is faster than running the tool multiple times
- Images from URLs may take longer (downloaded during conversion)
- Large files with many images will take longer to process

### Formatting Tips

**For Professional Documents:**
- Use page breaks to start new sections: `<!--pagebreak-->`
- Use blockquotes for important callouts: `> Important note`
- Use task lists for action items: `- [ ] Todo item`
- Add horizontal rules between major sections: `---`

**For Technical Documents:**
- Use code blocks with language hints: ` ```python `
- Use inline code for technical terms: `` `variable_name` ``
- Use superscript for math: `E = mc^2^`
- Use subscript for formulas: `H~2~O`

**For Highlighting:**
- Use **bold** for emphasis
- Use ==highlighting== for critical information
- Use ~~strikethrough~~ for deleted/outdated content
- Use *italic* for subtle emphasis

### Common Issues and Solutions

**"Output file looks different from Markdown"**
- Word uses its own styling system - some visual differences are normal
- Try different Word themes if colors don't match your preference

**"Images not showing up"**
- Check that image paths are correct relative to the .md file
- For URLs, ensure you have internet connection
- Some image formats may not be supported (stick to PNG, JPG, GIF)

**"Conversion is slow"**
- Large images slow down conversion - optimize images before converting
- URL-based images require downloading - use local files when possible

**"Special characters look wrong"**
- Save your Markdown file as UTF-8 in your text editor
- Most modern text editors default to UTF-8

### File Organization

For large projects, organize your files:
```
project/
├── docs/
│   ├── chapter1.md
│   ├── chapter2.md
│   └── images/
│       ├── diagram1.png
│       └── diagram2.png
└── output/
    ├── chapter1.docx
    └── chapter2.docx
```

Then convert with:
```bash
python md_to_docx.py docs/*.md
# Outputs will be in same folder as source files
```

## License

This project is open source. Feel free to use, modify, and distribute.

## Contributing

Contributions are welcome! If you find bugs or have feature requests, please open an issue.
