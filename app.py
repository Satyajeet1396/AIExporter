from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml import parse_xml
from markdown import markdown
import logging
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatters import HtmlFormatter

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
    
    # Check if we can find the expected HTML elements
    paragraphs = soup.find_all('p')
    headings = soup.find_all(['h1', 'h2', 'h3'])
    
    logging.info(f"Found {len(paragraphs)} <p> tags and {len(headings)} heading tags")

    # Process the HTML and convert it into DOCX
    for element in soup.find_all(True):
        logging.info(f"Processing Element: {element.name}")  # Log each element being processed
        
        if element.name == "p":
            p = doc.add_paragraph(element.get_text())
            # Apply formatting if needed (e.g., bold, italic)
            if element.find("strong"):  # Bold text
                for run in p.runs:
                    run.bold = True
            if element.find("em"):  # Italic text
                for run in p.runs:
                    run.italic = True
        
        elif element.name == "h1":
            doc.add_paragraph(element.get_text(), style='Heading 1')
        
        elif element.name == "h2":
            doc.add_paragraph(element.get_text(), style='Heading 2')
        
        elif element.name == "h3":
            doc.add_paragraph(element.get_text(), style='Heading 3')  # Correct style name
        
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
    
    # After the loop, log the length of the document content
    logging.info(f"Document created with {len(doc.paragraphs)} paragraphs")
    
    # If no paragraphs were added, log a warning
    if len(doc.paragraphs) == 0:
        logging.warning("No paragraphs were added to the document.")

    return doc

def convert_latex_to_omml(latex_code):
    # Example of a simple LaTeX to OMML (Office Math Markup Language) conversion
    # Replace this with actual LaTeX-to-OMML conversion logic (e.g., using a library like python-docx-oxml)
    # Here, we are just wrapping LaTeX in a placeholder for demo purposes.
    return f'<m:oMath><m:t>{latex_code}</m:t></m:oMath>'

# Example usage:

if __name__ == "__main__":
    html_content = """
    <h1>Heading 1</h1>
    <p>This is a <strong>bold</strong> paragraph with <em>italic</em> text.</p>
    <pre>
    def example_function():
        return "Hello, World!"
    </pre>
    <p>Here is a math formula: <span class="math">E = mc^2</span></p>
    <table>
        <tr><th>Header 1</th><th>Header 2</th></tr>
        <tr><td>Row 1 Col 1</td><td>Row 1 Col 2</td></tr>
        <tr><td>Row 2 Col 1</td><td>Row 2 Col 2</td></tr>
    </table>
    """
    
    # Convert HTML to DOCX
    doc = html_to_docx(html_content)
    
    if doc:
        # Save DOCX to a file
        doc.save("output.docx")
        print("Document saved as 'output.docx'.")
    else:
        print("Failed to create document.")
