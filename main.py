from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from io import BytesIO
import markdown2
from html2docx import HtmlToDocx

app = FastAPI()

@app.post("/convert-md-to-docx")
async def convert_md_to_docx(request: Request):
    data = await request.json()
    md_text = data.get("markdown", "")

    if not md_text:
        return {"error": "No markdown text provided"}

    html = markdown2.markdown(md_text)
    print("Generated HTML:", html)  # Debug

    converter = HtmlToDocx()
    docx_io = BytesIO()
    docx = converter.parse_html(html)
    docx.save(docx_io)
    docx_io.seek(0)

    headers = {
        'Content-Disposition': 'attachment; filename="converted.docx"'
    }

    return StreamingResponse(
        docx_io,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers=headers
    )
