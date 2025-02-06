import streamlit as st
import markdown2
from fpdf import FPDF
from docx import Document
import io
from markdown import markdown
from bs4 import BeautifulSoup

# ✅ Ensure session state for auto-rendering
if "markdown_input" not in st.session_state:
    st.session_state.markdown_input = ""

def parse_markdown(markdown_text):
    """Convert Markdown to BeautifulSoup for parsing."""
    html = markdown(markdown_text)
    return BeautifulSoup(html, "html.parser")

def convert_markdown_to_pdf(markdown_text):
    """Convert Markdown text to a PDF file (black text only)."""
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    # ✅ No colors, only black text
    pdf.set_font("Arial", size=12)

    soup = parse_markdown(markdown_text)
    for tag in soup.find_all():
        if tag.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            pdf.set_font("Arial", style="B", size=16 - (int(tag.name[1]) * 2))
        elif tag.name == "p":
            pdf.set_font("Arial", size=12)
        elif tag.name == "strong":
            pdf.set_font("Arial", style="B", size=12)
        elif tag.name == "em":
            pdf.set_font("Arial", style="I", size=12)

        pdf.multi_cell(0, 10, tag.get_text())

    return pdf.output(dest="S").encode("latin1")

def convert_markdown_to_docx(markdown_text):
    """Convert Markdown text to a DOCX file (black text only)."""
    doc = Document()
    soup = parse_markdown(markdown_text)

    for tag in soup.find_all():
        if tag.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            doc.add_heading(tag.get_text(), level=int(tag.name[1]))
        elif tag.name == "p":
            doc.add_paragraph(tag.get_text())
        elif tag.name == "strong":
            p = doc.add_paragraph()
            run = p.add_run(tag.get_text())
            run.bold = True  # ✅ Keep bold but black
        elif tag.name == "em":
            p = doc.add_paragraph()
            run = p.add_run(tag.get_text())
            run.italic = True  # ✅ Keep italic but black
        elif tag.name == "ul":
            for li in tag.find_all("li"):
                doc.add_paragraph(li.get_text(), style="List Bullet")
        elif tag.name == "ol":
            for li in tag.find_all("li"):
                doc.add_paragraph(li.get_text(), style="List Number")

    doc_bytes = io.BytesIO()
    doc.save(doc_bytes)
    doc_bytes.seek(0)
    return doc_bytes

def main():
    st.title("LLM Markdown to DOCX & PDF Converter")

    # ✅ Auto-render when typing
    markdown_text = st.text_area(
        "Paste your copied markdown below:",
        value=st.session_state.markdown_input,
        key="markdown_input"
    )

    # Only update the session state when the text changes
    st.session_state.markdown_input = markdown_text

    if markdown_text:
        html_text = markdown2.markdown(markdown_text)
        st.markdown(html_text, unsafe_allow_html=True)

        pdf_bytes = convert_markdown_to_pdf(markdown_text)
        docx_bytes = convert_markdown_to_docx(markdown_text)

        col1, col2 = st.columns(2)

        with col1:
            st.download_button(
                label="Download DOCX",
                data=docx_bytes,
                file_name="converted.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

        with col2:
            st.download_button(
                label="Download PDF",
                data=pdf_bytes,
                file_name="converted.pdf",
                mime="application/pdf"
            )

if __name__ == "__main__":
    main()
