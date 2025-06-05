from fastapi import FastAPI, Request, UploadFile, File
from fastapi.responses import StreamingResponse, JSONResponse, FileResponse
from io import BytesIO
import markdown2
from html2docx import html2docx
from bs4 import BeautifulSoup
import tempfile
import subprocess
import os
import json
from pathlib import Path
from datetime import datetime

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

def create_professional_markdown_template(md_content, client_name, project_title="Project Proposal"):
    """Create a professional markdown template matching the reference document structure"""
    
    # Extract headings for TOC - but exclude any existing TOC headings
    html_content = markdown2.markdown(md_content, extras=['tables', 'fenced-code-blocks'])
    soup = BeautifulSoup(html_content, 'html.parser')
    
    # Filter out any headings that contain "table of contents" or similar
    headings = []
    for heading in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6']):
        heading_text = heading.get_text().strip().lower()
        if 'table of contents' not in heading_text and 'toc' not in heading_text:
            headings.append(heading)
    
    # Build TOC
    toc_lines = []
    page_counter = 3  # Starting from page 3 (after cover and TOC)
    
    for heading in headings:
        level = int(heading.name[1])
        text = heading.get_text().strip()
        
        # Skip if it's a TOC-related heading
        if 'table of contents' in text.lower() or 'toc' in text.lower():
            continue
            
        indent = "  " * (level - 1) if level > 1 else ""
        
        # Create anchor-friendly ID
        anchor_id = text.lower().replace(' ', '-').replace(':', '').replace('&', '').replace('.', '')
        
        if level == 1:
            toc_lines.append(f"**{text}** ....................................... {page_counter}")
            page_counter += 2
        elif level == 2:
            toc_lines.append(f"{indent}{text} ....................................... {page_counter}")
            page_counter += 1
        else:
            toc_lines.append(f"{indent}{text}")
    
    # Create the enhanced markdown with YAML front matter
    current_date = datetime.now().strftime('%B %d, %Y')
    
    enhanced_md = f"""---
title: "{project_title}"
subtitle: "For: {client_name}"
date: "{current_date}"
geometry: "margin=1in"
fontsize: 11pt
linestretch: 1.15
documentclass: article
classoption: 
- onecolumn
header-includes: |
  \\usepackage{{fancyhdr}}
  \\usepackage{{graphicx}}
  \\usepackage{{xcolor}}
  \\usepackage{{sectsty}}
  \\usepackage{{titlesec}}
  \\usepackage{{tocloft}}
  
  % Header and footer
  \\pagestyle{{fancy}}
  \\fancyhf{{}}
  \\renewcommand{{\\headrulewidth}}{{0pt}}
  \\fancyfoot[C]{{\\thepage}}
  
  % Section formatting
  \\sectionfont{{\\color{{black}}\\large}}
  \\subsectionfont{{\\color{{black}}\\normalsize}}
  
  % TOC formatting
  \\renewcommand{{\\cftsecleader}}{{\\cftdotfill{{\\cftdotsep}}}}
  \\renewcommand{{\\cftsubsecleader}}{{\\cftdotfill{{\\cftdotsep}}}}
  
  % Title page
  \\makeatletter
  \\renewcommand{{\\maketitle}}{{
    \\begin{{titlepage}}
      \\centering
      \\vspace*{{2cm}}
      {{\\Huge\\bfseries \\@title}}\\\\[1cm]
      {{\\Large \\@subtitle}}\\\\[2cm]
      {{\\large \\@date}}
      \\vfill
    \\end{{titlepage}}
  }}
  \\makeatother
toc: true
toc-depth: 3
---

\\maketitle
\\newpage

\\tableofcontents
\\newpage

{md_content}
"""
    
    return enhanced_md

def create_reference_docx_template():
    """Create a reference.docx file with professional styling if it doesn't exist"""
    reference_path = "reference.docx"
    
    if not os.path.exists(reference_path):
        # Create a basic reference document with professional styling
        temp_md = """---
title: "Reference Document"
---

# Heading 1
This is a sample heading 1.

## Heading 2  
This is a sample heading 2.

### Heading 3
This is a sample heading 3.

Regular paragraph text with proper spacing and formatting.

| Column 1 | Column 2 | Column 3 |
|----------|----------|----------|
| Data 1   | Data 2   | Data 3   |
| Data 4   | Data 5   | Data 6   |
"""
        
        with tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode='w', encoding='utf-8') as tmp:
            tmp.write(temp_md)
            temp_md_path = tmp.name
        
        try:
            subprocess.run([
                "pandoc", temp_md_path, "-o", reference_path,
                "--standalone"
            ], check=True)
        except subprocess.CalledProcessError:
            pass  # If pandoc fails, we'll proceed without reference doc
        finally:
            if os.path.exists(temp_md_path):
                os.unlink(temp_md_path)

@app.post("/convert-md-to-html")
async def convert_md_to_html(request: Request):
    data = await request.json()
    md_text = data.get("markdown", "")
    client_name = data.get("client_name", "Client").strip()
    project_title = data.get("project_title", "Project Proposal")
    
    if not md_text:
        return {"error": "No markdown text provided"}
    
    html = markdown2.markdown(md_text, extras=[
        "tables",
        "fenced-code-blocks",
        "cuddled-lists",
        "footnotes",
        "header-ids",
        "toc"
    ])
    
    soup = BeautifulSoup(html, "html.parser")
    remove_empty_paragraphs_around(soup, ["table", "img", "h1", "h2", "h3", "h4", "h5", "h6"])
    clean_extra_spacing_around_tables(soup)
    
    # Add professional styling
    professional_html = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8">
        <title>{project_title} - {client_name}</title>
        <style>
            body {{
                font-family: 'Calibri', 'Arial', sans-serif;
                line-height: 1.6;
                margin: 0;
                padding: 0;
                color: #333;
            }}
            
            .cover-page {{
                page-break-after: always;
                height: 100vh;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                color: white;
            }}
            
            .cover-title {{
                font-size: 3rem;
                font-weight: bold;
                margin-bottom: 2rem;
                text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
            }}
            
            .cover-subtitle {{
                font-size: 1.5rem;
                margin-bottom: 3rem;
                opacity: 0.9;
            }}
            
            .cover-date {{
                font-size: 1.1rem;
                opacity: 0.8;
            }}
            
            .toc-page {{
                page-break-before: always;
                page-break-after: always;
                padding: 2rem;
                min-height: 80vh;
            }}
            
            .toc-title {{
                font-size: 2rem;
                font-weight: bold;
                text-align: center;
                margin-bottom: 2rem;
                border-bottom: 3px solid #667eea;
                padding-bottom: 1rem;
            }}
            
            .toc-content {{
                max-width: 800px;
                margin: 0 auto;
            }}
            
            .toc-item {{
                margin: 0.5rem 0;
                padding: 0.5rem 0;
                border-bottom: 1px dotted #ccc;
                display: flex;
                justify-content: space-between;
            }}
            
            .toc-item a {{
                text-decoration: none;
                color: #333;
                font-weight: 500;
            }}
            
            .toc-item a:hover {{
                color: #667eea;
            }}
            
            .content {{
                padding: 2rem;
                max-width: 1000px;
                margin: 0 auto;
            }}
            
            h1 {{
                color: #2c3e50;
                font-size: 2rem;
                margin-top: 2rem;
                margin-bottom: 1rem;
                padding-bottom: 0.5rem;
                border-bottom: 3px solid #667eea;
                page-break-before: always;
            }}
            
            h2 {{
                color: #34495e;
                font-size: 1.5rem;
                margin-top: 1.5rem;
                margin-bottom: 0.8rem;
            }}
            
            h3 {{
                color: #34495e;
                font-size: 1.2rem;
                margin-top: 1.2rem;
                margin-bottom: 0.6rem;
            }}
            
            p {{
                margin-bottom: 1rem;
                text-align: justify;
            }}
            
            ul, ol {{
                margin-bottom: 1rem;
                padding-left: 1.5rem;
            }}
            
            li {{
                margin-bottom: 0.3rem;
            }}
            
            table {{
                width: 100%;
                border-collapse: collapse;
                margin: 1.5rem 0;
                box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            }}
            
            th, td {{
                border: 1px solid #ddd;
                padding: 0.8rem;
                text-align: left;
                vertical-align: top;
            }}
            
            th {{
                background: linear-gradient(135deg, #667eea, #764ba2);
                color: white;
                font-weight: bold;
            }}
            
            tr:nth-child(even) {{
                background-color: #f8f9fa;
            }}
            
            blockquote {{
                border-left: 4px solid #667eea;
                padding-left: 1rem;
                margin: 1rem 0;
                font-style: italic;
                background-color: #f8f9fa;
                padding: 1rem;
            }}
            
            code {{
                background-color: #f4f4f4;
                padding: 0.2rem 0.4rem;
                border-radius: 3px;
                font-family: 'Courier New', monospace;
            }}
            
            pre {{
                background-color: #f4f4f4;
                padding: 1rem;
                border-radius: 5px;
                overflow-x: auto;
                margin: 1rem 0;
            }}
            
            @media print {{
                .cover-page, .toc-page {{
                    page-break-after: always;
                }}
                h1 {{
                    page-break-before: always;
                }}
                table {{
                    page-break-inside: avoid;
                }}
            }}
        </style>
    </head>
    <body>
        <div class="cover-page">
            <div class="cover-title">{project_title}</div>
            <div class="cover-subtitle">For: {client_name}</div>
            <div class="cover-date">{datetime.now().strftime('%B %d, %Y')}</div>
        </div>
        
        <div class="toc-page">
            <div class="toc-title">Table of Contents</div>
            <div class="toc-content" id="toc-content">
                <!-- TOC will be generated here -->
            </div>
        </div>
        
        <div class="content">
            {str(soup)}
        </div>
        
        <script>
            // Generate TOC
            document.addEventListener('DOMContentLoaded', function() {{
                const headings = document.querySelectorAll('.content h1, .content h2, .content h3');
                const tocContent = document.getElementById('toc-content');
                let pageCounter = 3;
                
                headings.forEach(function(heading, index) {{
                    const level = parseInt(heading.tagName.charAt(1));
                    const text = heading.textContent;
                    const id = 'heading-' + index;
                    heading.id = id;
                    
                    const tocItem = document.createElement('div');
                    tocItem.className = 'toc-item toc-level-' + level;
                    
                    const link = document.createElement('a');
                    link.href = '#' + id;
                    link.textContent = text;
                    
                    const pageNum = document.createElement('span');
                    pageNum.textContent = pageCounter;
                    
                    tocItem.appendChild(link);
                    tocItem.appendChild(pageNum);
                    tocContent.appendChild(tocItem);
                    
                    if (level === 1) pageCounter += 2;
                    else pageCounter += 1;
                }});
            }});
        </script>
    </body>
    </html>
    """
    
    html_bytes = professional_html.encode("utf-8")
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
    
    # Ensure reference document exists
    create_reference_docx_template()
    
    try:
        pandoc_cmd = [
            "pandoc", 
            tmp_html_path, 
            "-o", tmp_docx_path,
            "--standalone",
            "--toc",
            "--toc-depth=3",
        ]
        
        # Add reference document if it exists
        if os.path.exists("reference.docx"):
            pandoc_cmd.extend(["--reference-doc=reference.docx"])
        
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

@app.post("/convert-md-to-docx-professional")
async def convert_md_to_docx_professional(request: Request):
    """Professional direct conversion from markdown to DOCX with enhanced formatting"""
    data = await request.json()
    md_text = data.get("markdown", "")
    client_name = data.get("client_name", "Client").strip()
    project_title = data.get("project_title", "Project Proposal")
    
    if not md_text:
        return {"error": "No markdown text provided"}
    
    # Clean any existing TOC from the markdown
    def clean_existing_toc(md_content):
        lines = md_content.split('\n')
        cleaned_lines = []
        skip_toc = False
        
        for line in lines:
            line_lower = line.lower().strip()
            
            # Skip lines that look like TOC entries
            if ('table of contents' in line_lower or 
                line.strip().startswith('1.') or 
                line.strip().startswith('2.') or
                'page' in line_lower and '...' in line):
                skip_toc = True
                continue
                
            # Stop skipping when we hit a real heading
            if skip_toc and line.startswith('# ') and 'table of contents' not in line_lower:
                skip_toc = False
                
            if not skip_toc:
                cleaned_lines.append(line)
        
        return '\n'.join(cleaned_lines)
    
    # Clean the input markdown
    cleaned_md = clean_existing_toc(md_text)
    
    # Create enhanced markdown with YAML front matter (no manual TOC)
    current_date = datetime.now().strftime('%B %d, %Y')
    
    enhanced_md = f"""---
title: "{project_title}"
subtitle: "For: {client_name}"
date: "{current_date}"
geometry: "margin=1in"
fontsize: 11pt
linestretch: 1.15
documentclass: article
classoption: 
- onecolumn
header-includes: |
  \\usepackage{{fancyhdr}}
  \\usepackage{{graphicx}}
  \\usepackage{{xcolor}}
  \\usepackage{{sectsty}}
  \\usepackage{{titlesec}}
  \\usepackage{{tocloft}}
  
  % Header and footer
  \\pagestyle{{fancy}}
  \\fancyhf{{}}
  \\renewcommand{{\\headrulewidth}}{{0pt}}
  \\fancyfoot[C]{{\\thepage}}
  
  % Section formatting
  \\sectionfont{{\\color{{black}}\\large}}
  \\subsectionfont{{\\color{{black}}\\normalsize}}
  
  % TOC formatting
  \\renewcommand{{\\cftsecleader}}{{\\cftdotfill{{\\cftdotsep}}}}
  \\renewcommand{{\\cftsubsecleader}}{{\\cftdotfill{{\\cftdotsep}}}}
  
  % Title page
  \\makeatletter
  \\renewcommand{{\\maketitle}}{{
    \\begin{{titlepage}}
      \\centering
      \\vspace*{{2cm}}
      {{\\Huge\\bfseries \\@title}}\\\\[1cm]
      {{\\Large \\@subtitle}}\\\\[2cm]
      {{\\large \\@date}}
      \\vfill
    \\end{{titlepage}}
  }}
  \\makeatother
toc: true
toc-depth: 3
---

\\maketitle
\\newpage

\\tableofcontents
\\newpage

{cleaned_md}
"""
    
    with tempfile.NamedTemporaryFile(delete=False, suffix=".md", mode='w', encoding='utf-8') as tmp_md:
        tmp_md.write(enhanced_md)
        tmp_md_path = tmp_md.name
    
    safe_client_name = "".join(c for c in client_name if c.isalnum() or c in (" ", "_", "-")).strip()
    tmp_docx_path = tmp_md_path.replace(".md", ".docx")
    
    # Ensure reference document exists
    create_reference_docx_template()
    
    try:
        # Fixed pandoc command - removed --number-sections and simplified
        pandoc_cmd = [
            "pandoc",
            tmp_md_path,
            "-o", tmp_docx_path,
            "--standalone",
            "--table-of-contents",
            "--toc-depth=3",
            # Removed --number-sections to prevent 0.1, 0.2 numbering
        ]
        
        # Add reference document if it exists
        if os.path.exists("reference.docx"):
            pandoc_cmd.extend(["--reference-doc=reference.docx"])
        
        subprocess.run(pandoc_cmd, check=True)
        
        filename = f"Proposal for {safe_client_name}.docx"
        
        return FileResponse(
            tmp_docx_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except subprocess.CalledProcessError as e:
        return JSONResponse(status_code=500, content={"error": "Pandoc conversion failed", "details": str(e)})
    finally:
        if os.path.exists(tmp_md_path):
            os.unlink(tmp_md_path)
