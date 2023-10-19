import os
import subprocess  # Ensure you import the subprocess module
from docx2tex import docx2tex

# Path to the input DOCX file
docx_file = r'C:\Users\nayee\OneDrive\Desktop\EXAMINATION\Question.docx'

# Path to the output HTML file
html_file = r'C:\Users\nayee\OneDrive\Desktop\EXAMINATION\output.html'

# Convert DOCX to LaTeX
latex_content = docx2tex(docx_file)

# Create a LaTeX template for HTML rendering with MathJax
latex_template = r"""
<!DOCTYPE html>
<html>
<head>
    <script type="text/javascript" async
        src="https://cdnjs.cloudflare.com/ajax/libs/mathjax/2.7.7/MathJax.js?config=TeX-MML-AM_CHTML">
    </script>
</head>
<body>
{latex_content}
</body>
</html>
"""

# Replace {latex_content} with the converted LaTeX content
html_content = latex_template.format(latex_content=latex_content)

# Write the HTML content to the output file
with open(html_file, 'w', encoding='utf-8') as html_output:
    html_output.write(html_content)

# Optionally, you can print a message to confirm the HTML file creation
print(f"HTML file '{html_file}' has been created from the DOCX file.")
