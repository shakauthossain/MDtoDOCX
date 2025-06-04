from fastapi import FastAPI, Request
from fastapi.responses import StreamingResponse
from io import BytesIO
from html2docx import html2docx
from bs4 import BeautifulSoup

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


def clean_extra_spacing(soup):
    for p in soup.find_all("p"):
        if not p.text.strip():
            p.decompose()
    for table in soup.find_all("table"):
        next_sibling = table.find_next_sibling()
        while next_sibling and (next_sibling.name == "br" or (next_sibling.name == "p" and not next_sibling.text.strip())):
            temp = next_sibling.find_next_sibling()
            next_sibling.decompose()
            next_sibling = temp
    return str(soup)  # <-- Add this line


@app.post("/convert-html-to-docx")
async def convert_html_to_docx(request: Request):
    data = await request.json()
    html = data.get("html", "")
    client_name = data.get("client_name", "Client").strip()

    if not html:
        return {"error": "No HTML content provided"}

    # Clean and style HTML
    soup = BeautifulSoup(html, "html.parser")
    remove_empty_paragraphs_around(soup, ["table", "img", "h1", "h2", "h3", "h4", "h5", "h6"])
    cleaned_html = clean_extra_spacing(soup)

    docx_io: BytesIO = html2docx(cleaned_html, title=f"Proposal for {client_name}")
    docx_io.seek(0)

    # Safe filename
    safe_client_name = "".join(c for c in client_name if c.isalnum() or c in (" ", "_", "-")).strip()
    filename = f"Proposal for {safe_client_name}.docx"

    headers = {
        'Content-Disposition': f'attachment; filename="{filename}"'
    }

    return StreamingResponse(
        docx_io,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        headers=headers
    )
