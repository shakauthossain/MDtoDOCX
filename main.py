from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from io import BytesIO
import markdown2
from bs4 import BeautifulSoup
import tempfile
import subprocess
import os

app = FastAPI()

REFERENCE_DOCX = "notionhive_reference.docx"  # Ensure this file exists in the same directory

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
    for p in soup.find_all("p"):
        if not p.text.strip():
            p.decompose()

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

    # Prepare downloadable HTML
    html_bytes = cleaned_html.encode("utf-8")
    html_io = BytesIO(html_bytes)
    html_io.seek(0)

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
    html_content = await file.read()

    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(html_content)
        tmp_html_path = tmp_html.name

    tmp_docx_path = tmp_html_path.replace(".html", ".docx")

    try:
        # Build Pandoc command
        pandoc_cmd = ["pandoc", tmp_html_path, "-o", tmp_docx_path]

        # Use reference DOCX if available
        if os.path.exists(REFERENCE_DOCX):
            pandoc_cmd += ["--reference-doc", REFERENCE_DOCX]

        subprocess.run(pandoc_cmd, check=True)

        return FileResponse(
            tmp_docx_path,
            filename="Proposal.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except subprocess.CalledProcessError as e:
        return JSONResponse(status_code=500, content={"error": "Pandoc conversion failed", "details": str(e)})

    finally:
        if os.path.exists(tmp_html_path):
            os.unlink(tmp_html_path)
