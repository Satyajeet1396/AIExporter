import streamlit as st
import re
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatters import HtmlFormatter
from latex2mathml.converter import convert
import zipfile
import io

# Convert LaTeX to OMML (Word equation format)
def convert_latex_to_omml(latex):
    try:
        mathml = convert(latex)
        return f'<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{mathml}</m:oMath>'
    except Exception:
        return f"[Math: {latex}]"

# Convert HTML to DOCX
def html_to_docx(html_content):
    soup = BeautifulSoup(html_content, "html.parser")
    doc = Document()

    for element in soup.children:
        if element.name == "p":
            doc.add_paragraph(element.get_text())
        elif element.name in ["h1", "h2", "h3"]:
            doc.add_paragraph(element.get_text(), style=element.name.capitalize())
        elif element.name == "code":
            p = doc.add_paragraph()
            run = p.add_run(element.get_text())
            run.font.name = "Courier New"
        elif element.name == "pre":  # Code block
            code_text = element.get_text()
            lexer = get_lexer_by_name("python", stripall=True)
            formatter = HtmlFormatter(style="colorful")
            highlighted_code = highlight(code_text, lexer, formatter)
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
        elif element.name == "img":  # Handle images
            img_src = element.get("src")
            if img_src:
                try:
                    doc.add_picture(img_src)
                except Exception as e:
                    st.error(f"Failed to add image: {e}")
        elif element.name == "a":  # Handle links
            link_text = element.get_text()
            link_url = element.get("href")
            if link_url:
                p = doc.add_paragraph()
                run = p.add_run(link_text)
                run.add_hyperlink(link_url)

    return doc

# Streamlit UI
st.title("ChatGPT to DOCX Converter")
st.write("Paste ChatGPT output (HTML) below or upload an HTML file.")

uploaded_file = st.file_uploader("Upload an HTML file", type=["html"], accept_multiple_files=True)
input_text = st.text_area("Paste ChatGPT output (HTML format)")

if st.button("Convert to DOCX"):
    if uploaded_file:
        doc_files = []
        for file in uploaded_file:
            html_content = file.read().decode("utf-8")
            doc = html_to_docx(html_content)
            doc_path = f"converted_output_{file.name}.docx"
            doc.save(doc_path)
            doc_files.append(doc_path)
        
        if len(doc_files) > 1:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for doc_file in doc_files:
                    zip_file.write(doc_file)
            zip_buffer.seek(0)
            st.download_button("Download All DOCX", zip_buffer, file_name="ChatGPT_outputs.zip", mime="application/zip")
        else:
            with open(doc_files[0], "rb") as f:
                st.download_button("Download DOCX", f, file_name="ChatGPT_output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    elif input_text:
        html_content = input_text
        doc = html_to_docx(html_content)
        doc_path = "converted_output.docx"
        doc.save(doc_path)

        with open(doc_path, "rb") as f:
            st.download_button("Download DOCX", f, file_name="ChatGPT_output.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        st.error("Please provide input!")
        st.stop()
