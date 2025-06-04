from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from io import BytesIO
import markdown2
from html2docx import html2docx

app = FastAPI()

@app.post("/convert-md-to-docx")
async def convert_md_to_docx(request: Request):
    data = await request.json()
    md_text = data.get("markdown", "")
    
    if not md_text:
        return {"error": "No markdown text provided"}
    
    # Convert Markdown to HTML
    html = markdown2.markdown(md_text)
    
    # Convert HTML to DOCX - html2docx() returns a BytesIO object directly
    docx_io = html2docx(html, title="Converted Document")
    
    docx_io.seek(0)
    
    headers = {
        'Content-Disposition': 'attachment; filename="converted.docx"'
    }
    
    return StreamingResponse(
        docx_io,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers=headers
    )
