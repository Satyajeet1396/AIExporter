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

    # Parse the HTML
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
        if element.name == "code":
            text = element.get("class", []).get('text', '')
            p = document.add_paragraph()
            run = p.get_run()
            
            # Apply specific font settings for code
            lexer = get_lexer_by_name("python")
            formatter = HtmlFormatter(style="colorful")
            highlighted_code = highlighter(lexer, "python")(
                text,
                formatter,
            )
            
            doc.add_element(run, 'code', highlighted_code)
    
    # Handle spans with math tags
    for element in soup.find_all(True):
        if element.name == "span" and "math" in element.get("class", []):
            content = element.get_text()
            math_omml = convert_latex_to_omml(content)
            p = document.add_paragraph()
            p._element.append(parse_xml(math_omml))
    
    # Process tables
    for table in soup.find_all('table'):
        rows = table.find_all("tr")
        if not rows:
            continue
        cols = len(rows[0].find_all(["td", "th"]))
        doc.add_table(rows, cols)
    
    logging.info(f"Document created with {len(document.paragraphs)} paragraphs")
    
    if len(document.paragraphs) == 0:
        logging.warning("No paragraphs were added to the document.")
    else:
        logging.info(f"After processing HTML: {len(document.paragraphs)} paragraphs in DOCX")
        
    return document

def convert_latex_to_omml(latex_code):
    # Example of conversion with more accurate formatting
    lexer = get_lexer_by_name("python", stripall=True)
    formatter = HtmlFormatter(style="colorful")
    highlighted_code = highlighter(lexer, "python")(
        latex_code,
        formatter,
    )
    return highlighter(lexer, "python")(highlighted_code)

def main():
    st.title('HTML to DOCX Converter')

    file uploader for HTML content
    uploaded_file = st.file_uploader("Upload HTML file", type=["html"])
    
    html_content = None
    if uploaded_file:
        uploaded_file_content = uploaded_file.read().decode('utf-8')
        html_content = document.add_paragraph()
        html_content._element.append(document.parse_xml(uploaded_file.read().decode('utf-8')))
        html_content.get_text() = uploaded_file_content
    
    # Prepare HTML content
    if not html_content:
        st.error("Empty HTML file or no HTML content was uploaded.")
        return
    
    st.info("Converting HTML to DOCX...")
    
    try:
        doc = html_to_docx(html_content)
        logging.info(f"Conversion completed: {doc.__class__.__name__}")
        if doc:
            logging.info(f"Document created with {len(doc.paragraphs)} paragraphs")
            
            st.success("DOCX Document Created.")
            doc.save('output.docx')
            
            # Provide a download link
            with open('output.docx', 'rb') as f:
                st.download_file(
                    'output.docx',
                    'output.docx',
                    ['output.docx']
                )
        
        logging.info("Conversion failed or didn't produce proper DOCX file.")
    except Exception as e:
        logging.error(f"Error converting HTML to DOCX: {str(e)}")
        st.error("Failed to convert HTML. Please check your input.")

if __name__ == "__main__":
    main()
