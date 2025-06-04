from fastapi import FastAPI, Form
from fastapi.responses import StreamingResponse
from html2docx import html2docx
import io

app = FastAPI()

@app.post("/convert/")
async def convert_html_to_docx(html: str = Form(...)):
    # Create in-memory buffer
    docx_io = io.BytesIO()

    # Convert HTML string to DOCX
    html2docx(html, docx_io)

    # Reset buffer position
    docx_io.seek(0)

    # Return DOCX file as streaming response
    return StreamingResponse(
        docx_io,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=converted.docx"}
    )
