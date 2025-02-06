import streamlit as st
import logging
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from latex2mathml.converter import convert
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatters import HtmlFormatter
from fpdf import FPDF
from pptx import Presentation
import io

# Setup logging
logging.basicConfig(level=logging.INFO)

def convert_latex_to_omml(latex):
    try:
        mathml = convert(latex)
        return f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{mathml}</m:oMath>'
    except Exception as e:
        logging.error(f"Error converting LaTeX: {e}")
        return f"[Math: {latex}]"

def html_to_docx(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    doc = Document()
    
    for element in soup.find_all(True):
        if element.name == "p":
            doc.add_paragraph(element.get_text())
        elif element.name in ["h1", "h2", "h3"]:
            doc.add_paragraph(element.get_text(), style=element.name.capitalize())
        elif element.name == "code":
            p = doc.add_paragraph()
            run = p.add_run(element.get_text())
            run.font.name = "Courier New"
        elif element.name == "pre":
            lexer = get_lexer_by_name("python", stripall=True)
            formatter = HtmlFormatter(style="colorful")
            highlighted_code = highlight(element.get_text(), lexer, formatter)
            doc.add_paragraph(highlighted_code)
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
    
    return doc

def html_to_pdf(html_content):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.multi_cell(0, 10, html_content)
    return pdf

def html_to_ppt(html_content):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])
    textbox = slide.shapes.add_textbox(50, 50, 600, 400)
    tf = textbox.text_frame
    tf.text = html_content
    return prs

# Streamlit UI
st.title("ChatGPT Export Tool")
st.write("Convert ChatGPT outputs into DOCX, PDF, or PowerPoint.")

uploaded_file = st.file_uploader("Upload HTML File", type=["html"], accept_multiple_files=False)
input_text = st.text_area("Paste HTML Content")
export_format = st.radio("Select Export Format", ["DOCX", "PDF", "PowerPoint"])

if st.button("Convert & Download"):
    if uploaded_file:
        html_content = uploaded_file.read().decode("utf-8")
    elif input_text.strip():
        html_content = input_text.strip()
    else:
        st.error("Please provide input!")
        st.stop()
    
    if export_format == "DOCX":
        doc = html_to_docx(html_content)
        doc_path = "output.docx"
        doc.save(doc_path)
        with open(doc_path, "rb") as f:
            st.download_button("Download DOCX", f, file_name="output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    elif export_format == "PDF":
        pdf = html_to_pdf(html_content)
        pdf_output = io.BytesIO()
        pdf.output(pdf_output, dest='S')
        pdf_output.seek(0)
        st.download_button("Download PDF", pdf_output, file_name="output.pdf", mime="application/pdf")
    elif export_format == "PowerPoint":
        ppt = html_to_ppt(html_content)
        ppt_path = "output.pptx"
        ppt.save(ppt_path)
        with open(ppt_path, "rb") as f:
            st.download_button("Download PowerPoint", f, file_name="output.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
