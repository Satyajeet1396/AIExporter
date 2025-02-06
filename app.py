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
    
    # Convert Markdown-style syntax to HTML (if needed)
    html_content = markdown(html_content)
    
    soup = BeautifulSoup(html_content, "html.parser")
    logging.info(f"Parsed HTML content: {soup.prettify()[:1000]}")  # Log parsed HTML to debug
    
    doc = Document()
    
    # Process the HTML and convert it into DOCX
    for element in soup.find_all(True):
        logging.info(f"Processing Element: {element.name}, Content: {element.get_text()}")  # Log each element being processed
        
        if element.name == "p":
            p = doc.add_paragraph()
            for content in element.contents:
                if content.name == "strong":
                    run = p.add_run(content.get_text())
                    run.bold = True
                elif content.name == "em":
                    run = p.add_run(content.get_text())
                    run.italic = True
                elif content.name == "span":
                    run = p.add_run(content.get_text())
                    if "style" in content.attrs:
                        # Apply styles (e.g., color, font size)
                        pass
                else:
                    p.add_run(str(content))
        
        elif element.name in ["h1", "h2", "h3"]:
            doc.add_paragraph(element.get_text(), style=f'Heading {element.name[1]}')
        
        elif element.name == "code":
            p = doc.add_paragraph()
            run = p.add_run(element.get_text())
            run.font.name = "Courier New"
            run.font.size = Pt(10)  # Set the font size for code
            run.font.color.rgb = RGBColor(0, 0, 255)  # Set font color for code (blue)
        
        elif element.name == "pre":
            lexer = get_lexer_by_name("python", stripall=True)
            formatter = HtmlFormatter(style="colorful")
            highlighted_code = highlight(element.get_text(), lexer, formatter)
            code_soup = BeautifulSoup(highlighted_code, "html.parser")
            p = doc.add_paragraph()
            run = p.add_run(code_soup.get_text())
            run.font.name = "Courier New"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0, 0, 255)
        
        elif element.name == "table":
            rows = element.find_all("tr")
            if rows:
                cols = len(rows[0].find_all(["td", "th"]))
                table = doc.add_table(rows=0, cols=cols)
                for row in rows:
                    cells = row.find_all(["td", "th"])
                    row_cells = table.add_row().cells
                    for i, cell in enumerate(cells):
                        row_cells[i].text = cell.get_text()
        
        elif element.name == "span" and "math" in element.get("class", []):
            latex_code = element.get_text()
            math_omml = convert_latex_to_omml(latex_code)
            p = doc.add_paragraph()
            p._element.append(parse_xml(math_omml))
    
    logging.info(f"Document created with {len(doc.paragraphs)} paragraphs")
    
    if len(doc.paragraphs) == 0:
        logging.warning("No paragraphs were added to the document.")

    return doc

def convert_latex_to_omml(latex_code):
    # Placeholder for LaTeX-to-OMML conversion
    return f'<m:oMath><m:t>{latex_code}</m:t></m:oMath>'

# Streamlit GUI
st.title('HTML to DOCX Converter')
uploaded_file = st.file_uploader("Upload HTML file", type=["html"])
html_content = st.text_area("Or enter HTML content", height=300)

if uploaded_file:
    html_content = uploaded_file.read().decode("utf-8")
    st.success(f"HTML file uploaded successfully. Length: {len(html_content)} characters.")
elif html_content:
    st.success("HTML content provided.")
else:
    st.warning("Please upload an HTML file or provide HTML content.")

if st.button("Generate DOCX"):
    if html_content:
        st.info("Converting HTML to DOCX...")
        doc = html_to_docx(html_content)
        
        if doc:
            doc_path = "output.docx"
            doc.save(doc_path)
            st.success(f"Document successfully created: {doc_path}")
            
            with open(doc_path, "rb") as f:
                st.download_button(
                    label="Download DOCX",
                    data=f,
                    file_name="output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        else:
            st.error("Failed to convert HTML to DOCX.")
    else:
        st.warning("Please upload or provide HTML content to convert.")
