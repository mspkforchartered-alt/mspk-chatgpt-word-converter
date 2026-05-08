from fastapi import FastAPI, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from latex2mathml.converter import convert as latex_to_mathml
import re
import os
import uuid

app = FastAPI()

# =========================================
# CORS - Production Ready
# =========================================
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class TextInput(BaseModel):
    text: str

# =========================================
# PAGE BORDER HELPER
# =========================================
def add_page_border(doc):
    """Adds a clean single-line border to every page."""
    sec = doc.sections
    for section in sec:
        sectPr = section._sectPr
        pgBorders = OxmlElement('w:pgBorders')
        pgBorders.set(qn('w:offsetFrom'), 'page')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border_el = OxmlElement(f'w:{border_name}')
            border_el.set(qn('w:val'), 'single')
            border_el.set(qn('w:sz'), '4')  
            border_el.set(qn('w:space'), '24') 
            border_el.set(qn('w:color'), 'auto')
            pgBorders.append(border_el)
        sectPr.append(pgBorders)

# =========================================
# CLEANING UTILS
# =========================================
def remove_file(path: str):
    if os.path.exists(path): os.remove(path)

def clean_text(text):
    """Removes #, ** and other AI markdown artifacts from text."""
    text = re.sub(r"^#+\s*", "", text) # Remove # headers
    text = text.replace("**", "") # Remove bold stars
    text = text.replace("---", "").replace("###", "").replace("##", "")
    return text.strip()

def clean_equation(eq):
    """Deep cleaning for equations to prevent broken boxes (like the Ke fraction)."""
    # 1. Fix double-escaped slashes
    eq = eq.replace("\\\\", "\\")
    
    # 2. CRITICAL FIX: Ensure % is escaped for LaTeX (\%)
    # This prevents the denominator from disappearing or becoming a blank box.
    eq = re.sub(r'(?<!\\)%', r'\\%', eq)
    
    # 3. Remove \ from already escaped percentages (\% -> %) ONLY for the final box display
    # but keep it as \% for the MathML converter. (Handled in add_word_equation)
    
    # 4. Strip AI wrappers and bold stars
    eq = eq.replace("\\[", "").replace("\\]", "").replace("\\(", "").replace("\\)", "")
    eq = eq.replace("$$", "").replace("$", "").replace("**", "")
    
    return eq.strip().lstrip("[").rstrip("]")

def is_equation(line):
    line = line.strip()
    if not line: return False
    patterns = [r"\\frac", r"\\sqrt", r"\\sum", r"\\int", r"\\times", r"\^", r"_", r"=", r"\\\(", r"\\\[", r"\$", r"\\text"]
    return any(re.search(p, line) for p in patterns)

# =========================================
# DOCUMENT ENGINES
# =========================================
def add_word_equation(paragraph, equation):
    """Converts LaTeX to Word OMML with safety fallbacks for broken math."""
    cleaned_eq = clean_equation(equation)
    if not cleaned_eq: return 
    
    try:
        # Convert LaTeX -> MathML
        mathml = latex_to_mathml(cleaned_eq)
        if 'xmlns=' not in mathml:
            mathml = mathml.replace("<math>", '<math xmlns="http://w3.org">')

        xslt_path = os.path.join(os.getcwd(), "MML2OMML.XSL")
        if not os.path.exists(xslt_path):
            # Fallback to italic text if XSL file is missing on server
            run = paragraph.add_run(f" {cleaned_eq.replace('\\%', '%')} ")
            run.italic = True
            run.font.name = "Times New Roman"
            return

        # Transform MathML to Word's OMML format
        xslt = etree.parse(xslt_path)
        transform = etree.XSLT(xslt)
        mathml_dom = etree.fromstring(mathml.encode("utf-8"))
        omml_dom = transform(mathml_dom)
        paragraph._element.append(omml_dom.getroot())
        
    except Exception as e:
        # If conversion fails, don't leave a blank box—show the raw math in italics
        print(f"Math Error: {e}")
        run = paragraph.add_run(f" {cleaned_eq.replace('\\%', '%')} ")
        run.italic = True
        run.font.name = "Times New Roman"

def add_table(doc, lines):
    rows = []
    for line in lines:
        if "|" not in line or re.search(r"[-:]{3,}", line): continue
        # Clean stars (**) out of table cells
        cols = [clean_text(c.strip()) for c in line.split("|") if c.strip()]
        if cols: rows.append(cols)
    
    if not rows: return
    
    table = doc.add_table(rows=len(rows), cols=max(len(r) for r in rows))
    table.style = "Table Grid"
    for i, row in enumerate(rows):
        for j, cell in enumerate(row):
            if j < len(table.columns):
                cell_obj = table.cell(i, j)
                cell_obj.text = "" # Clear default text
                p = cell_obj.paragraphs[0]
                run = p.add_run(cell)
                run.font.name = "Times New Roman"
                run.font.size = Pt(10)

# =========================================
# MAIN GENERATOR
# =========================================
@app.post("/generate-docx")
async def generate_docx(data: TextInput, background_tasks: BackgroundTasks):
    text = data.text
    doc = Document()

    # 1. Page & Margin Setup
    section = doc.sections[0]
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    add_page_border(doc)

    # 2. Global Professional Style
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(11)
    style.paragraph_format.line_spacing = 1.0 # Single spacing
    style.paragraph_format.space_after = Pt(0)

    # Title
    title = doc.add_paragraph()
    title_run = title.add_run("FINANCIAL REPORT")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    title.paragraph_format.space_after = Pt(18)

    lines = text.split("\n")
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        
        # Skip noise and standalone brackets
        if not line or line in ["[", "]", "(", ")"]:
            if not line: doc.add_paragraph()
            i += 1
            continue

        # Handle Tables
        if "|" in line:
            table_lines = []
            while i < len(lines) and "|" in lines[i]:
                table_lines.append(lines[i])
                i += 1
            add_table(doc, table_lines)
            continue

        # Handle Equations (Fixes the broken Ke boxes)
        if is_equation(line) and len(line) < 600:
            p = doc.add_paragraph()
            p.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(10)
            add_word_equation(p, line)
            i += 1
            continue

        # Handle Normal Text & Headings
        line_clean = clean_text(line)
        p = doc.add_paragraph()
        # Bold short lines or WN markers
        if line_clean.isupper() or (len(line_clean) < 50 and (line_clean.startswith("WN") or line_clean.endswith(":"))):
            run = p.add_run(line_clean)
            run.bold = True
            p.paragraph_format.space_before = Pt(8)
        else:
            p.add_run(line_clean)
        i += 1

    # Save and cleanup
    filename = f"report_{uuid.uuid4()}.docx"
    filepath = os.path.join(os.getcwd(), filename)
    doc.save(filepath)
    background_tasks.add_task(remove_file, filepath)
    
    return FileResponse(
        filepath, 
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
        filename="Financial_Report.docx"
    )

@app.get("/")
def home(): return {"status": "online", "message": "Ready"}
