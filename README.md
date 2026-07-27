# blog2html

Converts a Microsoft Word document (.docx) containing blog posts into a styled HTML page with collapsible sections.

## Overview

`BloggConverter.py` reads a Word document (`blog published.docx`) that contains dated blog entries, parses each entry by detecting date lines (e.g., `MM/DD/YY` or "This entry was posted in January 15, 2024"), and generates a single self-contained HTML file. Each blog post appears as a collapsible card showing the date and a preview line in the header; clicking the header expands the full post content.

The generated HTML includes:
- Responsive layout with Google Fonts (Inter)
- Collapsible blog post cards with expand/collapse animation
- Date and first-line preview in each card header
- Hover effects and visual styling

## Prerequisites

- Python 3.6+
- `python-docx` library

## Setup

```bash
# Clone the repository
git clone https://github.com/biofool/blog2html.git
cd blog2html

# Install the dependency
pip install python-docx
```

## How to Run

1. Place your Word document named `blog published.docx` in the same directory as the script (or update the path in the `__main__` block).
2. Run the script:

```bash
python BloggConverter.py
```

3. The output file `blog_posts.html` will be created in the same directory.
4. Open `blog_posts.html` in any web browser to view the blog posts.

## Project Structure

```
blog2html/
├── BloggConverter.py    # Main script: converts .docx to HTML
└── blog published.docx  # Input Word document (not included in repo)
```

## Notes

- The script expects the input file to be named `blog published.docx` by default. To use a different filename, modify the `docx_file_path` variable in the `__main__` block.
- Formatting such as bold, italic, and other rich text attributes from the Word document are not preserved; only paragraph text is extracted.
- Empty paragraphs in the source document are rendered as non-breaking spaces to maintain spacing.
