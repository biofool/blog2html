# blog2html — Technical Marketing Summary

## One-Line Positioning

A Python utility that transforms a Word document of dated blog posts into a clean, interactive HTML page with collapsible sections.

## Target Users / Personas

- **Bloggers and content creators** who maintain blog posts in Microsoft Word and want to publish them as a simple, self-contained HTML page.
- **Non-technical users** who need a one-step conversion from `.docx` to web-ready HTML without setting up a CMS or static site generator.

## Key Features (Grounded in Code)

- **Date-based parsing**: Detects dates in `MM/DD/YY` format and "This entry was posted in [Month] [Day], [Year]" format using regex to split the document into individual blog posts (`BloggConverter.py`, `DATE_REGEX`).
- **Collapsible HTML output**: Each blog post is rendered as a card with a clickable header that expands/collapses the full content, implemented with vanilla JavaScript (`BloggConverter.py`, `HTML_TEMPLATE_START`/`HTML_TEMPLATE_END`).
- **Preview lines**: Each collapsed card shows the date and the first non-empty line of the post as a preview (`BloggConverter.py`, `preview_line` logic).
- **Self-contained HTML**: All CSS and JavaScript are embedded in the output file — no external dependencies beyond a Google Fonts link.
- **Responsive styling**: Uses the Inter font family, max-width layout, and hover animations for a modern look.

## Technical Differentiators

- **Zero-config conversion**: No configuration files, no build tools, no CMS — just point it at a `.docx` file and get HTML.
- **Minimal dependencies**: Only requires `python-docx`; the output HTML uses no JavaScript frameworks.
- **Date-aware splitting**: Automatically segments a single Word document into multiple blog posts based on date markers, rather than treating it as one monolithic document.

## Use Cases

- Converting an archive of blog posts stored in a single Word document into a browsable HTML page.
- Creating a quick portfolio or archive page from existing Word-formatted content.
- Generating a shareable HTML summary of dated journal entries or newsletters.

## Benefits / Value Proposition

- Eliminates manual copy-paste from Word to HTML.
- Produces a professional, interactive result with no web development knowledge required.
- Output is a single file that can be emailed, hosted anywhere, or opened locally.

## Tech Stack

- **Language**: Python 3.6+
- **Library**: `python-docx` (for reading Word documents)
- **Output**: Self-contained HTML5 with embedded CSS and vanilla JavaScript

## Known Limitations

- Input filename is hardcoded to `blog published.docx` (must be edited in source to change).
- Rich text formatting (bold, italic, fonts, colors) from the Word document is not preserved — only plain paragraph text is extracted.
- No command-line arguments for specifying input/output paths at runtime.
- Date detection relies on specific formats; non-standard date formats may not be recognized.
