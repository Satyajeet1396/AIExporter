import streamlit as st
import markdown2
from fpdf import FPDF
from docx import Document
import tempfile
import os
from markdown import markdown
from bs4 import BeautifulSoup

def convert_markdown_to_pdf(markdown_text, pdf_filename):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    lines = markdown_text.split("\n")
    for line in lines:
        pdf.multi_cell(0, 10, line)
    
    pdf.output(pdf_filename)

def convert_markdown_to_docx(markdown_text, docx_filename):
    doc = Document()
    html = markdown(markdown_text)
    soup = BeautifulSoup(html, "html.parser")
    
    for tag in soup.find_all():
        if tag.name in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            doc.add_heading(tag.get_text(), level=int(tag.name[1]))
        elif tag.name == "p":
            doc.add_paragraph(tag.get_text())
        elif tag.name == "strong":
            p = doc.add_paragraph()
            p.add_run(tag.get_text()).bold = True
        elif tag.name == "em":
            p = doc.add_paragraph()
            p.add_run(tag.get_text()).italic = True
        elif tag.name == "ul":
            for li in tag.find_all("li"):
                doc.add_paragraph(li.get_text(), style="List Bullet")
        elif tag.name == "ol":
            for li in tag.find_all("li"):
                doc.add_paragraph(li.get_text(), style="List Number")
    
    doc.save(docx_filename)

def main():
    st.title("LLM Markdown to DOCX & PDF Converter")
    
    markdown_text = st.text_area("Paste your copied markdown below:")
    
    if st.button("Render Markdown"):
        if markdown_text:
            html_text = markdown2.markdown(markdown_text)
            st.markdown(html_text, unsafe_allow_html=True)
        else:
            st.warning("Please paste some markdown text.")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Convert to DOCX"):
            if markdown_text:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmpfile:
                    convert_markdown_to_docx(markdown_text, tmpfile.name)
                    st.download_button(label="Download DOCX", data=open(tmpfile.name, "rb").read(), file_name="converted.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                    os.unlink(tmpfile.name)
            else:
                st.warning("Please paste some markdown text.")
    
    with col2:
        if st.button("Convert to PDF"):
            if markdown_text:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmpfile:
                    convert_markdown_to_pdf(markdown_text, tmpfile.name)
                    st.download_button(label="Download PDF", data=open(tmpfile.name, "rb").read(), file_name="converted.pdf", mime="application/pdf")
                    os.unlink(tmpfile.name)
            else:
                st.warning("Please paste some markdown text.")
    
if __name__ == "__main__":
    main()
