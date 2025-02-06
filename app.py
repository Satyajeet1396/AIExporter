import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from markdown import markdown
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatters import HtmlFormatter
import logging

logging.basicConfig(level=logging.INFO)

def html_to_docx(html_content):
    if not html_content.strip():
        logging.warning("Empty HTML content received for DOCX conversion.")
        return None
    
    logging.info(f"HTML Content Received: {html_content[:500]}")  # Log first 500 characters of HTML

    soup = BeautifulSoup(html_content, "html.parser")
    
    document = Document()
    paragraphs = soup.find_all('p')
    headings = soup.find_all(['h1', 'h2', 'h3'])
    
    doc = document.add_paragraph
    current_paragraph = doc
    for paragraph in paragraphs:
        current_paragraph = doc.add_paragraph(paragraph.get_text())
        # Apply formatting if needed (e.g., bold, italic)
        if "strong" in paragraph.get("class", []):
            run = current_paragraph.get_run()
            run.bold = True
        elif "em" in paragraph.get("class", []):
            run = current_paragraph.get_run()
            run.italic = True
    
    for heading in headings:
        doc.add_paragraph(heading, style='Heading 1')
    
    # Handle code blocks within elements with 'code' class
    for element in soup.find_all(True):
        if element.name == "span" and "math" in element.get("class", []):
            content = element.get_text()
            math_omml = convert_latex_to_omml(content)
            p = document.add_paragraph()
            p._element.append(parse_xml(math_omml))
    
    # Handle tables
    for table in soup.find_all('table'):
        rows = table.find_all("tr")
        if not rows:
            continue
        cols = len(rows[0].find_all(["td", "th"]))
        document.add_table(rows, cols)
    
    logging.info(f"Document created with {len(document.paragraphs)} paragraphs")
    
    if len(document.paragraphs) == 0:
        logging.warning("No paragraphs were added to the document.")
    else:
        logging.info(f"After processing HTML: {len(document.paragraphs)} paragraphs in DOCX")

def convert_latex_to_omml(latex_code):
    # Example of conversion with more accurate formatting
    lexer = get_lexer_by_name("python", stripall=True)
    formatter = HtmlFormatter(style="colorful")
    highlighted_code = highlighter(lexer, "python")(
        latex_code,
        formatter,
    )
    return highlighter(lexer, "python")(highlighted_code)

# Streamlit GUI

st.title('HTML to DOCX Converter')

# File uploader for HTML content
uploaded_file = st.file_uploader("Upload HTML file", type=["html"])

# Text area for HTML input (if not uploading a file)
html_content = st.text_area("Or enter HTML content", height=300)

if uploaded_file:
    html_content = uploaded_file.read().decode("utf-8")
    st.success("HTML file uploaded successfully.")
elif html_content:
    st.success("HTML content provided.")
else:
    st.warning("Please upload or provide HTML content to convert.")

# If the user uploads a file or enters HTML content
if uploaded_file:
    doc = html_to_docx(html_content)
    
    if doc:
        # Save DOCX to a file
        doc_path = "output.docx"
        doc.save(doc_path)
        st.success(f"Document created: {doc_path}")
        
        # Provide a download link for the DOCX file
        with open(doc_path, "rb") as f:
            st.download_file(
                'output.docx',
                'output.docx',
                ['output.docx']
            )
    else:
        st.error("Failed to convert HTML to DOCX.")
else:
    st.warning("Please upload or provide HTML content to convert.")
