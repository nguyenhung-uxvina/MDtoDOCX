# MDtoDOCX - Complete Feature Test

This document demonstrates all the features supported by the MDtoDOCX converter, including both original and newly added features.

## Text Formatting

### Basic Formatting
- **Bold text** using double asterisks
- *Italic text* using single asterisks
- ***Bold and italic*** using triple asterisks
- `Inline code` using backticks

### New Formatting Features
- ~~Strikethrough text~~ using double tildes
- Regular text with superscript: E = mc^2^ (for mathematical notation)
- Chemical formulas with subscript: H~2~O
- ==Highlighted text== for emphasis

## Headings

# Heading 1
## Heading 2
### Heading 3
#### Heading 4
##### Heading 5
###### Heading 6

## Lists

### Unordered List
- First item
- Second item
- Third item with nested items:
  - Nested item A
  - Nested item B
    - Double nested item
- Fourth item

### Ordered List
1. First step
2. Second step
3. Third step
4. Fourth step

### Task Lists (New Feature!)
- [x] Completed task
- [x] Another completed task
- [ ] Pending task
- [ ] Another pending task
- [x] Design the feature
- [ ] Implement the feature
- [ ] Test the feature

## Code Blocks

Here's a Python code example:

```python
def factorial(n):
    """Calculate factorial of n"""
    if n <= 1:
        return 1
    return n * factorial(n - 1)

# Example usage
result = factorial(5)
print(f"5! = {result}")
```

And here's a JavaScript example:

```javascript
function greet(name) {
    return `Hello, ${name}!`;
}

console.log(greet("World"));
```

## Links

Visit [GitHub](https://github.com) for code hosting.

Check out the [Python Documentation](https://docs.python.org) for learning Python.

Learn about [Markdown](https://www.markdownguide.org) syntax.

## Blockquotes (New Feature!)

> This is a blockquote. It's perfect for highlighting important information,
> quotes from other sources, or emphasizing key points in your document.

> "The best way to predict the future is to invent it."
> - Alan Kay

## Tables

### Simple Table

| Feature | Status | Priority |
|---------|--------|----------|
| Blockquotes | ✓ Done | High |
| Strikethrough | ✓ Done | Medium |
| Task Lists | ✓ Done | High |
| Page Breaks | ✓ Done | Medium |

### Complex Table

| Language | Year | Type | Popular Use Cases |
|----------|------|------|-------------------|
| Python | 1991 | Interpreted | Data Science, Web, Automation |
| JavaScript | 1995 | Interpreted | Web Development, Node.js |
| Rust | 2010 | Compiled | Systems Programming, WebAssembly |
| Go | 2009 | Compiled | Cloud Services, CLI Tools |

## Horizontal Rules (New Feature!)

You can separate sections with horizontal rules:

---

Content above the line.

Content below the line.

---

## Images

Images are supported (both local and remote):

![Placeholder Image](https://via.placeholder.com/400x200.png?text=Sample+Image)

## Mixed Content Example

You can combine **bold**, *italic*, `code`, ~~strikethrough~~, ==highlight==, and even super^script^ or sub~script~ all in the same paragraph. This makes it incredibly flexible for technical writing!

The formula for water is H~2~O, and Einstein's famous equation is E = mc^2^.

## Scientific Notation

Some examples with superscript and subscript:

- Mathematical: x^2^ + y^2^ = z^2^
- Chemical: CO~2~, H~2~SO~4~, CH~3~COOH
- Footnote reference^[1]^

## Page Break Example

The next section will appear on a new page when converted to DOCX.

<!--pagebreak-->

## New Page!

This content appears after a page break. This is useful for:
- Creating chapter breaks
- Separating major sections
- Starting new topics on fresh pages
- Organizing long documents

## Real-World Use Case

Here's a practical example combining multiple features:

### Project Status Report

**Project:** MDtoDOCX Converter
**Date:** November 3, 2025
**Status:** ==In Progress==

#### Completed Tasks
- [x] Basic Markdown parsing
- [x] Table support
- [x] Image handling
- [x] Blockquote formatting
- [x] Strikethrough text
- [x] Task list checkboxes

#### Pending Tasks
- [ ] Add footnote support
- [ ] Implement custom themes
- [ ] Add table of contents generation
- [ ] Support for embedded videos

#### Key Metrics

| Metric | Value | Change |
|--------|-------|--------|
| Lines of Code | 450 | +150 |
| Features | 15 | +8 |
| Test Coverage | 85% | +10% |

> **Note:** This project has exceeded initial expectations with the addition of many user-requested features!

---

## Conclusion

The MDtoDOCX converter now supports a comprehensive set of Markdown features, making it suitable for:

1. Technical documentation
2. Project reports
3. Academic papers
4. User manuals
5. Personal notes and wikis

With features like ~~outdated text~~, ==highlighted important notes==, task lists, and proper page breaks, you can create professional documents with ease!

---

*Document generated with MDtoDOCX - Version 2.0*
