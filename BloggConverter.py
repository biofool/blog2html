import re
from docx import Document
import os


def convert_docx_to_html(docx_path="blog_posts.html", output_html_path="blog_posts.html"):
    """
    Converts a Word document with blog posts into an HTML file with collapsible sections.
    Each section previews the date and the first line of the blog post.

    Args:
        docx_path (str): The file path to the input .docx document.
        output_html_path (str): The file path where the generated HTML will be saved.
    """
    # HTML template for the output
    HTML_TEMPLATE_START = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>My Blog Posts</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600&display=swap" rel="stylesheet">
    <style>
        body {
            font-family: 'Inter', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 900px;
            margin: 20px auto;
            padding: 0 15px;
            background-color: #f8f8f8;
            border-radius: 8px;
            box-shadow: 0 4px 10px rgba(0,0,0,0.05);
        }
        h1 {
            text-align: center;
            color: #2c3e50;
            margin-bottom: 30px;
            font-weight: 600;
        }
        .blog-post-container {
            background-color: #fff;
            border: 1px solid #e0e0e0;
            margin-bottom: 15px;
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 2px 5px rgba(0,0,0,0.08);
            transition: all 0.3s ease;
        }
        .blog-post-container:hover {
            box-shadow: 0 4px 10px rgba(0,0,0,0.12);
        }
        .blog-post-header {
            cursor: pointer;
            font-weight: bold;
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 15px 20px;
            background-color: #e9f5ff; /* Light blue for header */
            color: #3498db; /* Darker blue text */
            border-bottom: 1px solid #d0e9ff;
            transition: background-color 0.3s ease;
            border-radius: 8px 8px 0 0;
        }
        .blog-post-header:hover {
            background-color: #dbeaff;
        }
        .blog-date {
            font-size: 0.9em;
            color: #666;
            flex-shrink: 0;
            margin-right: 15px;
            font-weight: 400;
        }
        .blog-preview-line {
            flex-grow: 1;
            font-size: 1.1em;
            overflow: hidden;
            white-space: nowrap;
            text-overflow: ellipsis;
            font-weight: 600;
        }
        .blog-post-content {
            display: none; /* Hidden by default */
            padding: 20px;
            border-top: 1px solid #f0f0f0;
            animation: fadeIn 0.5s ease-out;
        }
        .blog-post-header:after {
            content: '+'; /* Plus sign when collapsed */
            margin-left: 15px;
            font-size: 1.2em;
            transition: transform 0.3s ease;
            color: #3498db;
        }
        .blog-post-header.active:after {
            content: '-'; /* Minus sign when expanded */
            transform: rotate(0deg); /* No rotation needed for minus */
        }
        .blog-post-content p {
            margin-bottom: 1em;
        }
        .blog-post-content p:last-child {
            margin-bottom: 0;
        }
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-10px); }
            to { opacity: 1; transform: translateY(0); }
        }
    </style>
</head>
<body>

    <h1>My Blog Posts</h1>

    <div id="blog-posts-wrapper">
    """

    HTML_TEMPLATE_END = """
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function() {
            const headers = document.querySelectorAll('.blog-post-header');

            headers.forEach(header => {
                header.addEventListener('click', function() {
                    this.classList.toggle('active');
                    const content = this.nextElementSibling;
                    if (content.style.display === "block") {
                        content.style.display = "none";
                    } else {
                        content.style.display = "block";
                    }
                });
            });
        });
    </script>

</body>
</html>
    """

    # Regex to find dates in MM/DD/YY format at the beginning of a line
    DATE_REGEX = re.compile(
        r"(This entry was posted in .*? ([A-Za-z]+) (\d{1,2}), (\d{4}))|(\b\d{1,2}/\d{1,2}/\d{2,4}\b)"
    )
    try:
        document = Document("blog published.docx")
    except Exception as e:
        print(f"Error: Could not open or read the Word document. Please ensure it's a valid .docx file. Error: {e}")
        return

    blog_posts_data = []
    current_post = None

    # State variable to track if we are currently parsing content for a post
    # after its date and preview line have been identified.
    in_post_content_collection = False

    for i, paragraph in enumerate(document.paragraphs):
        text = paragraph.text.strip()

        if "This entry was posted" in text:
            cleaned_text = re.sub(r"This entry was posted (in|on)\s", "", text)
            text= cleaned_text

        # Check if the current paragraph is a date, and if we are not already collecting content
        if DATE_REGEX.match(text) and not in_post_content_collection:
            # If there was a previous post being built, save it
            if current_post:
                blog_posts_data.append(current_post)

            # Start a new post
            current_post = {
                'date': text,
                'preview_line': '',
                'content': []
            }
            in_post_content_collection = True  # Now we expect the preview line and then content
            continue  # Move to the next paragraph

        # If we are in the process of collecting content for a post
        if in_post_content_collection:
            if DATE_REGEX.match(text):  # Found a new date, means the previous post has ended
                if current_post:
                    blog_posts_data.append(current_post)
                # Start a new post with this date
                current_post = {
                    'date': text,
                    'preview_line': '',
                    'content': []
                }
                # Keep in_post_content_collection as True as we're starting a new post
                continue

            # If the preview line hasn't been set yet and we have text
            if current_post['preview_line'] == '' and text:
                current_post['preview_line'] = text
                # Add the preview line to the content as well, as it's part of the full post
                current_post['content'].append(f"<p>{text}</p>")
            elif text:
                # Add subsequent non-empty paragraphs to the current post's content
                # Note: This simple approach does not preserve bold/italic/etc.
                # For more advanced formatting, you'd need to iterate through paragraph.runs
                current_post['content'].append(f"<p>{text}</p>")
            else:
                # Add a non-breaking space for empty paragraphs to maintain spacing
                current_post['content'].append("<p>&nbsp;</p>")

    # Add the last post if it exists
    if current_post:
        blog_posts_data.append(current_post)

    # Generate the HTML content for all blog posts
    blog_posts_html_sections = []
    for post in blog_posts_data:
        # Ensure preview_line is not empty, if it is, use a placeholder or the first few words of content
        if not post['preview_line'] and post['content']:
            # Extract plain text from the first content paragraph for preview
            first_content_text = re.sub(r'<[^>]+>', '', post['content'][0])  # Remove HTML tags
            post['preview_line'] = (first_content_text[:100] + '...') if len(
                first_content_text) > 100 else first_content_text
            if not post['preview_line']:  # Fallback if even after stripping tags, it's empty
                post['preview_line'] = "No preview available"

        content_html = "\n".join(post['content'])
        blog_posts_html_sections.append(f"""
        <div class="blog-post-container">
            <div class="blog-post-header">
                <span class="blog-date">{post['date']}</span>
                <span class="blog-preview-line">{post['preview_line']}</span>
            </div>
            <div class="blog-post-content">
                {content_html}
            </div>
        </div>
        """)

    full_html = HTML_TEMPLATE_START + "\n".join(blog_posts_html_sections) + HTML_TEMPLATE_END

    # Save the generated HTML to a file
    try:
        with open(output_html_path, 'w', encoding='utf-8') as f:
            f.write(full_html)
        print(f"Conversion successful! HTML saved to {output_html_path}")
    except Exception as e:
        print(f"Error: Could not write the HTML file. Error: {e}")


# --- How to use the script ---
# 1. Save the above code as a Python file (e.g., `convert_blog.py`).
# 2. Make sure your Word document `blog published.docx` is in the same directory,
#    or update `docx_file_path` to the correct path.
# 3. Run the script from your terminal: `python convert_blog.py`
# 4. An `blog_posts.html` file will be created in the same directory.
# 5. Open `blog_posts.html` in your web browser to view the result.

# Example usage:
if __name__ == "__main__":
    docx_file_path = "blog published.docx"  # Make sure this file is in the same directory
    if os.path.exists(docx_file_path):
        convert_docx_to_html(docx_file_path)
    else:
        print(f"Error: The file '{docx_file_path}' was not found.")
        print("Please ensure the Word document is in the same directory as the script, or provide the full path.")
