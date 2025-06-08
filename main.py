from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from io import BytesIO
import markdown2
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
    for p in soup.find_all("p"):
        if not p.text.strip():
            p.decompose()

    for table in soup.find_all("table"):
        next_sibling = table.find_next_sibling()
        while next_sibling and (next_sibling.name == "br" or (next_sibling.name == "p" and not next_sibling.text.strip())):
            temp = next_sibling.find_next_sibling()
            next_sibling.decompose()
            next_sibling = temp

def add_table_borders_to_html(html_content: str) -> str:
    """Enhance <table> tags in HTML with inline border styles so Pandoc shows them in DOCX."""
    soup = BeautifulSoup(html_content, "html.parser")

    for table in soup.find_all("table"):
        table['border'] = "1"
        table['style'] = "border: 1px solid black; border-collapse: collapse; width: 100%;"

        for row in table.find_all("tr"):
            for cell in row.find_all(["th", "td"]):
                existing_style = cell.get('style', '')
                new_style = "border: 1px solid black; padding: 6px;"
                cell['style'] = f"{existing_style} {new_style}".strip()

    return str(soup)

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

    cleaned_html = add_table_borders_to_html(str(soup))

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

    cleaned_html = add_table_borders_to_html(html_content.decode("utf-8"))

    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_html:
        tmp_html.write(cleaned_html.encode("utf-8"))
        tmp_html_path = tmp_html.name

    tmp_docx_path = tmp_html_path.replace(".html", ".docx")

    try:
        subprocess.run(["pandoc", tmp_html_path, "-o", tmp_docx_path], check=True)

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
