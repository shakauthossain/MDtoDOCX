from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse
from io import BytesIO
import markdown2
from html2docx import html2docx
from bs4 import BeautifulSoup
import tempfile
import subprocess
import os

app = FastAPI()

def remove_empty_paragraphs_around(soup, tag_names):
    for tag_name in tag_names:
        for tag in soup.find_all(tag_name):
            for prev in tag.find_all_previous():
                if prev.name == "p" and not prev.text.strip():
                    prev.decompose()
                    break
                elif prev.name not in ["p", "br", None]:
                    break
            for next_ in tag.find_all_next():
                if next_.name == "p" and not next_.text.strip():
                    next_.decompose()
                    break
                elif next_.name not in ["p", "br", None]:
                    break

def clean_extra_spacing_around_tables(soup):
    # Remove empty or whitespace-only <p> tags
    for p in soup.find_all("p"):
        if not p.text.strip():
            p.decompose()

    # Remove <br> or empty <p> directly after <table>
    for table in soup.find_all("table"):
        next_sibling = table.find_next_sibling()
        while next_sibling and (next_sibling.name == "br" or (next_sibling.name == "p" and not next_sibling.text.strip())):
            temp = next_sibling.find_next_sibling()
            next_sibling.decompose()
            next_sibling = temp

@app.post("/convert-md-to-html")
async def convert_md_to_html(request: Request):
    data = await request.json()
    md_text = data.get("markdown", "")
    client_name = data.get("client_name", "Client").strip()

    if not md_text:
        return {"error": "No markdown text provided"}

    html = markdown2.markdown(md_text, extras=[
        "tables",
        "fenced-code-blocks",
        "cuddled-lists",
        "footnotes"
    ])

    soup = BeautifulSoup(html, "html.parser")
    remove_empty_paragraphs_around(soup, ["table", "img", "h1", "h2", "h3", "h4", "h5", "h6"])
    clean_extra_spacing_around_tables(soup)

    cleaned_html = str(soup)

    # Prepare HTML as a downloadable fileAdd commentMore actions
    html_bytes = cleaned_html.encode("utf-8")
    html_io = BytesIO(html_bytes)
    html_io.seek(0)

    # Safe filename
    safe_client_name = "".join(c for c in client_name if c.isalnum() or c in (" ", "_", "-")).strip()
    filename = f"Proposal for {safe_client_name}.html"

    headers = {
        'Content-Disposition': f'attachment; filename="{filename}"'
    }

    return StreamingResponse(
        html_io,
        media_type='text/html',
        headers=headers
    )

@app.post("/convert-html-to-docx")
async def convert_html_to_docx(file: UploadFile = File(...)):
    # Read uploaded HTML content
    html_content = await file.read()

    # Save HTML content to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(html_content)
        tmp_html_path = tmp_html.name

    # Create path for DOCX output
    tmp_docx_path = tmp_html_path.replace(".html", ".docx")

    try:
        # Use Pandoc to convert HTML to DOCX
        subprocess.run(["pandoc", tmp_html_path, "-o", tmp_docx_path], check=True)

        # Return the generated DOCX file
        return FileResponse(
            tmp_docx_path,
            filename="Proposal.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except subprocess.CalledProcessError as e:
        return JSONResponse(status_code=500, content={"error": "Pandoc conversion failed", "details": str(e)})
    finally:
        # Clean up the temporary HTML file
        if os.path.exists(tmp_html_path):
            os.unlink(tmp_html_path)
        # Clean up the DOCX file after response is returned by FileResponse (optional if you use FileResponse streaming)
