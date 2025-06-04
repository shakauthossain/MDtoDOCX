from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from io import BytesIO
import markdown2
from html2docx import html2docx
from bs4 import BeautifulSoup
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

app = FastAPI()

def remove_empty_paragraphs_around(soup, tag_names):
    for tag_name in tag_names:
        for tag in soup.find_all(tag_name):
            # Remove empty <p> before the tag
            for prev in tag.find_previous_siblings('p'):
                if not prev.text.strip():
                    prev.decompose()
                    break
                elif prev.name not in ["p", "br", None]:
                    break

            # Remove empty <p> after the tag
            for next_ in tag.find_next_siblings('p'):
                if not next_.text.strip():
                    next_.decompose()
                    break
                elif next_.name not in ["p", "br", None]:
                    break

@app.post("/convert-md-to-docx")
async def convert_md_to_docx(request: Request):
    data = await request.json()
    md_text = data.get("markdown", "")
    client_name = data.get("client_name", "Client").strip()
    
    if not md_text:
        return {"error": "No markdown text provided"}
    
    # Convert Markdown to HTML
    html = markdown2.markdown(md_text, extras=[
        "tables", 
        "fenced-code-blocks", 
        "cuddled-lists", 
        "footnotes"
    ])
    
    # Clean up spacing
    soup = BeautifulSoup(html, "html.parser")
    remove_empty_paragraphs_around(soup, ["table", "img", "h1", "h2", "h3", "h4", "h5", "h6"])
    cleaned_html = str(soup)
    
    # Convert to DOCX
    document = Document()
    for paragraph in cleaned_html.split("<p>"):
        if paragraph:
            run = document.add_paragraph()
            run.add_run(paragraph)
            run.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.left
            for table in paragraph.split("<table>"):
                if table:
                    table_rows = table.split("<tr>")
                    for table_row in table_rows:
                        if table_row:
                            table_cells = table_row.split("<td>")
                            for table_cell in table_cells:
                                if table_cell:
                                    table_cell_text = table_cell.split("</td>")[0]
                                    run = document.add_paragraph()
                                    run.add_run(table_cell_text)
                                    run.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.left
                                    run.paragraph_format.space_after = document.styles['Heading 1'].paragraph_format.space_after
    
    document.save("output.docx")
    
    # Sanitize filename
    safe_client_name = "".join(c for c in client_name if c.isalnum() or c in (" ", "_", "-")).strip()
    filename = f"Proposal for {safe_client_name}.docx"
    
    headers = {
        'Content-Disposition': f'attachment; filename="{filename}"'
    }
    
    with open("output.docx", "rb") as file:
        return StreamingResponse(
            file,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            headers=headers
        )
