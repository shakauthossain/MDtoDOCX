import markdown2
from io import BytesIO
from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from html2docx import html2docx

app = FastAPI()

@app.post("/convert-md-to-docx")
async def convert_md_to_docx(request: Request):
    data = await request.json()
    md_text = data.get("markdown", "")
    if not md_text:
        return {"error": "No markdown text provided"}

    html = markdown2.markdown(md_text)
    docx_io = BytesIO()
    html2docx(html, docx_io)
    docx_io.seek(0)

    headers = {
        'Content-Disposition': 'attachment; filename="converted.docx"'
    }
    return StreamingResponse(docx_io, media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document', headers=headers)
