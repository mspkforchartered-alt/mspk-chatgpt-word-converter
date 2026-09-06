from fastapi import FastAPI, BackgroundTasks
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from lxml import etree
from latex2mathml.converter import convert as latex_to_mathml

import re
import os
import uuid
import logging

# =========================================
# LOGGING
# =========================================

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
)

logger = logging.getLogger(__name__)

# =========================================
# FASTAPI
# =========================================

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
allow_origins=[
    "https://www.mspkchatgpttoword.com",
    "https://mspkchatgpttoword.com",
    "http://127.0.0.1:5500",
    "http://localhost:5500"
],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# =========================================
# INPUT MODEL
# =========================================

class TextInput(BaseModel):
    text: str = Field(..., min_length=1, max_length=500000)
    title: str = Field(
        default="CONVERTED INTO WORD"
    )

# =========================================
# PAGE BORDER
# =========================================

def add_page_border(doc):

    for section in doc.sections:

        sectPr = section._sectPr

        pgBorders = OxmlElement('w:pgBorders')

        pgBorders.set(
            qn('w:offsetFrom'),
            'page'
        )

        for border_name in [
            'top',
            'left',
            'bottom',
            'right'
        ]:

            border_el = OxmlElement(
                f'w:{border_name}'
            )

            border_el.set(
                qn('w:val'),
                'single'
            )

            border_el.set(
                qn('w:sz'),
                '4'
            )

            border_el.set(
                qn('w:space'),
                '24'
            )

            border_el.set(
                qn('w:color'),
                '0a2540'
            )

            pgBorders.append(border_el)

        sectPr.append(pgBorders)

# =========================================
# REMOVE AI GARBAGE
# =========================================

def remove_ai_garbage(text):

    garbage_patterns = [

        r"^✅\s+",
        r"^\{\s*$",
        r'^\}\s*$',
        r'^"text"\s*:',
        r"^Convert this into",
        r"^Create\s+",
        r"^Design\s+",
        r"if you want next",
        r"run.*pip.*install",

    ]

    cleaned_lines = []

    for line in text.splitlines():

        skip = False

        for pattern in garbage_patterns:

            if re.search(
                pattern,
                line,
                re.IGNORECASE
            ):

                skip = True
                break

        if skip:
            continue

        # IMPORTANT:
        # Preserve empty lines because they can separate
        # paragraphs, equations, tables and sections.

        cleaned_lines.append(line.rstrip())

    return "\n".join(cleaned_lines)

# =========================================
# CLEAN TEXT
# =========================================

def clean_text(text):

    text = re.sub(
        r"^#+\s*",
        "",
        text
    )

    text = text.replace(
        "**",
        ""
    )

    return text.strip()

# =========================================
# RECONSTRUCT CHATGPT PLAIN-MATH BLOCKS
# =========================================

def reconstruct_chatgpt_blocks(lines):

    result = []

    i = 0

    equation_operators = {
        "=",
        "−",
        "-",
        "+",
        "×",
        "÷",
        "±",
        "≤",
        "≥",
        "≠",
        "≈",
    }

    math_symbols = (
        "∑",
        "Σ",
        "√",
        "∞",
        "∫",
        "π",
        "α",
        "β",
        "γ",
        "δ",
        "σ",
        "μ",
        "ρ",
        "λ",
        "θ",
        "φ",
        "ω",
        "%",
    )

    while i < len(lines):

        current = lines[i].strip()

        # -----------------------------------------
        # EMPTY LINE
        # -----------------------------------------

        if not current:

            result.append("")

            i += 1

            continue

        # -----------------------------------------
        # NEVER RECONSTRUCT TABLES
        # -----------------------------------------

        if is_table_line(current):

            result.append(lines[i])

            i += 1

            continue


        # -----------------------------------------
        # CASE 1
        #
        # Profit
        # =
        # Revenue
        # −
        # Costs
        #
        # becomes:
        #
        # Profit = Revenue − Costs
        # -----------------------------------------

        if (
            i + 2 < len(lines)
            and
            lines[i + 1].strip() in equation_operators
            and
            lines[i + 2].strip()
        ):

            left = current

            operator = lines[i + 1].strip()

            right = lines[i + 2].strip()

            # -------------------------------------
            # Check that this really looks like math
            # -------------------------------------

            looks_math = (

                operator == "="

                or

                any(
                    symbol in right
                    for symbol in math_symbols
                )

                or

                re.search(
                    r"\d|[+\-*/=×÷^_()]",
                    right
                )
            )

            if looks_math:

                equation = (
                    left
                    + " "
                    + operator
                    + " "
                    + right
                )

                j = i + 3

                # ---------------------------------
                # Collect continuation lines
                # ---------------------------------

                while j < len(lines):

                    nxt = lines[j].strip()

                    if not nxt:
                        break

                    if is_table_line(nxt):
                        break

                    if is_heading(nxt):
                        break

                    # Another obvious paragraph
                    if (
                        len(nxt.split()) > 12
                        and
                        not any(
                            symbol in nxt
                            for symbol in math_symbols
                        )
                    ):

                        break

                    # Continue if mathematically related
                    if (
                        nxt in equation_operators
                        or
                        any(
                            symbol in nxt
                            for symbol in math_symbols
                        )
                        or
                        re.search(
                            r"\d|[+\-*/=×÷^_()]",
                            nxt
                        )
                    ):

                        equation += " " + nxt

                        j += 1

                        continue

                    break

                result.append(equation)

                i = j

                continue

        # -----------------------------------------
        # CASE 2
        #
        # LABEL
        # =
        # VALUE
        #
        # -----------------------------------------

        if (
            i + 2 < len(lines)
            and
            lines[i + 1].strip() == "="
            and
            lines[i + 2].strip()
        ):

            left = current
            right = lines[i + 2].strip()

            if (
                len(left) <= 100
                and
                len(right) <= 200
            ):

                result.append(
                    f"{left} = {right}"
                )

                i += 3

                continue

        # -----------------------------------------
        # NORMAL LINE
        # -----------------------------------------

        result.append(lines[i])

        i += 1

    return result

# =========================================
# FIX TEXT FRACTIONS
# =========================================

def fix_text_fractions(eq):

    pattern = r'\\frac\s*\{([^{}]+)\}\s*\{([^{}]+)\}'

    matches = re.findall(pattern, eq)

    for num, den in matches:

        new_num = num
        new_den = den

        if (
            " " in num
            and
            r"\text" not in num
            and
            "\\" not in num
        ):

            new_num = rf"\text{{{num}}}"

        if (
            " " in den
            and
            r"\text" not in den
            and
            "\\" not in den
        ):

            new_den = rf"\text{{{den}}}"

        old = rf"\frac{{{num}}}{{{den}}}"

        new = rf"\frac{{{new_num}}}{{{new_den}}}"

        eq = eq.replace(old, new)

    return eq

# =========================================
# CLEAN EQUATION
# =========================================

def simplify_nested_fractions(eq):

    # DOUBLE NESTED FRACTION

    pattern = (
        r'\\frac\s*\{'
        r'\s*\\frac\{([^{}]+)\}\{([^{}]+)\}'
        r'\s*\}'
        r'\s*\{'
        r'\s*\\frac\{([^{}]+)\}\{([^{}]+)\}'
        r'\s*\}'
    )

    matches = re.findall(pattern, eq)

    for a, b, c, d in matches:

        # CONVERT:
        # (a/b)/(c/d)
        # INTO:
        # (a*d)/(b*c)

        replacement = (
            rf'\\frac{{({a})({d})}}{{({b})({c})}}'
        )

        old = (
            rf'\frac{{\frac{{{a}}}{{{b}}}}}'
            rf'{{\frac{{{c}}}{{{d}}}}}'
        )

        eq = eq.replace(old, replacement)

    return eq

def clean_equation(eq):

    eq = eq.strip()

    # NORMALIZE MULTIPLE EQUAL SIGNS
    # Converts ====, ===, == into =
    eq = re.sub(r'={2,}', '=', eq)

    # REMOVE WRAPPERS

    eq = eq.replace("\\[", "")
    eq = eq.replace("\\]", "")

    eq = eq.replace("\\(", "")
    eq = eq.replace("\\)", "")

    eq = eq.replace("$$", "")

    # ALIGN SUPPORT

    if (
        r"\begin{aligned}" in eq
        or
        r"\begin{align}" in eq
    ):

        eq = eq.replace(
            r"\begin{aligned}",
            ""
        )

        eq = eq.replace(
            r"\end{aligned}",
            ""
        )

        eq = eq.replace(
            r"\begin{align}",
            ""
        )

        eq = eq.replace(
            r"\end{align}",
            ""
        )

        eq = eq.replace("&", "")

        eq = eq.replace(
            r"\\",
            " "
        )

    # KEEP LATEX SAFE

    eq = eq.replace("\\\\", "\\")

    # FIX TEXT FRACTIONS

    eq = fix_text_fractions(eq)

    # SYMBOL FIXES

    eq = eq.replace(
        "×",
        r"\times "
    )

    # REMOVE UNSUPPORTED

    eq = eq.replace("{text}", "")

    eq = eq.replace(r"\boxed", "")

    eq = eq.replace(r"\left", "")
    eq = eq.replace(r"\right", "")

    eq = eq.replace(r"\Big", "")
    eq = eq.replace(r"\big", "")

    eq = eq.replace(r"\displaystyle", "")

    # NORMALIZE

    eq = re.sub(
        r'\n+',
        ' ',
        eq
    )

    eq = re.sub(
        r'\s+',
        ' ',
        eq
    )
    # SIMPLIFY DEEP NESTED FRACTIONS

    eq = simplify_nested_fractions(eq)

    return eq.strip()

# =========================================
# EQUATION DETECTION
# =========================================

def is_equation_line(line):

    line = line.strip()

    if not line:
        return False

    # -----------------------------------------
    # EXPLICIT LATEX EQUATION
    # -----------------------------------------

    if (
        r"\text{" in line
        or
        r"\frac{" in line
        or
        r"\sqrt{" in line
        or
        r"\sum" in line
        or
        r"\int" in line
        or
        r"\prod" in line
        or
        r"\lim" in line
        or
        r"\partial" in line
        or
        r"\ln" in line
        or
        r"\log" in line
        or
        r"\sin" in line
        or
        r"\cos" in line
        or
        r"\tan" in line
        or
        r"\infty" in line
        or
        r"\alpha" in line
        or
        r"\beta" in line
        or
        r"\gamma" in line
        or
        r"\delta" in line
        or
        r"\sigma" in line
        or
        r"\mu" in line
        or
        r"\rho" in line
        or
        r"\lambda" in line
        or
        r"\theta" in line
    ):

        if (
            "=" in line
            or
            any(
                operator in line
                for operator in (
                    "+",
                    "-",
                    "−",
                    "×",
                    "÷",
                    "/",
                    "^",
                    "_"
                )
            )
        ):

            return True

    # -----------------------------------------
    # FINANCIAL EQUATION INDICATORS
    # -----------------------------------------

    financial_patterns = [

        r"\bWACC\s*=",
        r"\bNPV\s*=",
        r"\bIRR\s*=",
        r"\bROE\s*=",
        r"\bROA\s*=",
        r"\bROI\s*=",
        r"\bEPS\s*=",
        r"\bCAPM\s*=",
        r"\bEOQ\s*=",
        r"\bEAR\s*=",
        r"\bCCC\s*=",
        r"\bDCL\s*=",
        r"\bDOL\s*=",
        r"\bDFL\s*=",
        r"\bVaR\s*=",

    ]

    for pattern in financial_patterns:

        if re.search(
            pattern,
            line,
            re.IGNORECASE
        ):

            return True

    # -----------------------------------------
    # EQUATION WITH EQUAL SIGN
    # -----------------------------------------

    if "=" in line:

        if re.search(
            r"[A-Za-z0-9)\]}₹%]\s*=\s*[A-Za-z0-9(\\{₹%]",
            line
        ):

            return True

        # Handles:
        # Cost of Machine = ₹1,00,000 + ₹50,000
        # Assets = Liabilities + Equity
        # Total Cost = 7,00,000 + 3,00,000

        if re.search(
            r"[A-Za-z₹%0-9}\]]\s*=\s*",
            line
        ):

            if re.search(
                r"\d|₹|\\[A-Za-z]+|\+|-|−|×|÷|/|\*",
                line
            ):

                return True

    # -----------------------------------------
    # STRONG MATHEMATICAL SYMBOLS
    # -----------------------------------------

    if any(
        symbol in line
        for symbol in (
            "∑",
            "Σ",
            "√",
            "∞",
            "∫",
            "≤",
            "≥",
            "≠",
            "≈",
            "±",
        )
    ):

        return True

    return False

# =========================================
# HEADING DETECTION
# =========================================

def is_heading(line):

    return (

        (
            line.strip().isupper()
            and
            len(line.strip()) > 3
        )

        or

        (
            line.strip().endswith(":")
            and
            len(line.strip()) < 70
        )

        or

        re.match(
            r"^(WN|worked note)\s+\d+",
            line,
            re.IGNORECASE
        )

        or

        re.match(
            r"^#+\s+",
            line
        )
    )

# =========================================
# TABLE DETECTION
# =========================================

def is_table_line(line):

    return line.count("|") >= 2

# =========================================
# WORD EQUATION ENGINE
# =========================================

def add_word_equation(
    paragraph,
    equation
):

    cleaned_eq = clean_equation(
        equation
    )

    if not cleaned_eq:
        return

    # =====================================
    # ESCAPE % ONLY FOR REAL EQUATIONS
    # =====================================

    cleaned_eq = re.sub(
        r'(?<!\\)%',
        r'\%',
        cleaned_eq
    )

    # =====================================
    # REMOVE ONLY EXCESSIVE FRACTION FALLBACK
    # =====================================

    frac_count = cleaned_eq.count(
        r"\frac"
    )

    # Do NOT fall back for 3, 4, 5, etc.
    # fractions. They can still be valid
    # Word equations.
    #
    # Only use fallback for extremely
    # complicated equations.

    if frac_count >= 12:

        fallback_eq = cleaned_eq

        fallback_eq = fallback_eq.replace(
            r"\frac",
            "/"
        )

        fallback_eq = fallback_eq.replace(
            "{",
            "("
        )

        fallback_eq = fallback_eq.replace(
            "}",
            ")"
        )

        fallback_eq = fallback_eq.replace(
            r"\times",
            "×"
        )

        fallback_eq = fallback_eq.replace(
            r"\sqrt",
            "√"
        )

        fallback_eq = re.sub(
            r'\\[a-zA-Z]+',
            '',
            fallback_eq
        )

        run = paragraph.add_run(
            fallback_eq
        )

        run.font.name = "Times New Roman"
        run.italic = True
        run.font.size = Pt(11)

        return

    try:

        # =====================================
        # LATEX → MATHML
        # =====================================

        mathml = latex_to_mathml(
            cleaned_eq
        )

        # =====================================
        # ENSURE MATHML NAMESPACE
        # =====================================

        if "xmlns=" not in mathml:

            mathml = mathml.replace(
                "<math>",
                '<math xmlns="http://www.w3.org/1998/Math/MathML">'
            )

        # =====================================
        # NORMALIZE MATHML
        # =====================================

        mathml_dom = etree.fromstring(
            mathml.encode("utf-8")
        )

        # -------------------------------------
        # FIX LONG / STRETCHED MINUS
        # -------------------------------------

        for mo in mathml_dom.iter():

            if etree.QName(mo).localname != "mo":
                continue

            operator_text = (
                "".join(mo.itertext())
                .strip()
            )

            # Unicode minus → normal ASCII minus
            if operator_text in (
                "−",
                "–",
                "—"
            ):

                mo.text = "-"

                mo.set(
                    "stretchy",
                    "false"
                )

                mo.set(
                    "fence",
                    "false"
                )

                mo.set(
                    "form",
                    "infix"
                )

            # Explicitly prevent ASCII minus
            # from becoming a stretchy operator

            elif operator_text == "-":

                mo.set(
                    "stretchy",
                    "false"
                )

                mo.set(
                    "fence",
                    "false"
                )

                mo.set(
                    "form",
                    "infix"
                )

        # =====================================
        # XSLT FILE
        # =====================================

        xslt_path = os.path.join(
            os.getcwd(),
            "MML2OMML.XSL"
        )

        if not os.path.exists(
            xslt_path
        ):

            run = paragraph.add_run(
                cleaned_eq
            )

            run.font.name = (
                "Cambria Math"
            )

            run.font.size = Pt(11)

            return

        # =====================================
        # MATHML → OMML
        # =====================================

        xslt = etree.parse(
            xslt_path
        )

        transform = etree.XSLT(
            xslt
        )

        omml_dom = transform(
            mathml_dom
        )

        omml_root = (
            omml_dom.getroot()
        )

        # =====================================
        # EXTRA OMML MINUS PROTECTION
        # =====================================

        for node in omml_root.iter():

            if etree.QName(node).localname == "t":

                if node.text:

                    node.text = (
                        node.text
                        .replace("−", "-")
                        .replace("–", "-")
                        .replace("—", "-")
                    )

        # =====================================
        # APPEND EDITABLE WORD EQUATION
        # =====================================

        paragraph._element.append(
            omml_root
        )

    except Exception as e:

        logger.error(
            f"Equation Error: {e}"
        )

        # =====================================
        # SAFE FALLBACK
        # =====================================

        run = paragraph.add_run(
            cleaned_eq
        )

        run.font.name = (
            "Cambria Math"
        )

        run.font.size = Pt(11)

# =========================================
# PROFESSIONAL TABLE ENGINE
# =========================================

def add_table(doc, lines):

    from docx.shared import Inches
    from docx.enum.section import WD_ORIENT

    # =====================================
    # PARSE TABLE ROWS
    # =====================================

    rows = []

    for line in lines:

        if re.search(
            r"[-:]{3,}",
            line
        ):
            continue

        raw_cols = line.split("|")

        # REMOVE OUTER EMPTY CELLS

        if (
            raw_cols
            and
            raw_cols[0].strip() == ""
        ):
            raw_cols = raw_cols[1:]

        if (
            raw_cols
            and
            raw_cols[-1].strip() == ""
        ):
            raw_cols = raw_cols[:-1]

        cols = [

            clean_text(c.strip())

            for c in raw_cols

        ]

        if cols:
            rows.append(cols)

    if not rows:
        return

    # =====================================
    # MAX COLUMN COUNT
    # =====================================

    max_cols = max(
        len(r)
        for r in rows
    )

    # =====================================
    # CREATE TABLE
    # =====================================

    table = doc.add_table(

        rows=len(rows),

        cols=max_cols

    )

    table.style = "Table Grid"

    # VERY IMPORTANT
    table.autofit = False

    # =====================================
    # COLUMN WIDTHS
    # =====================================

    if max_cols >= 2:

        table.columns[0].width = Inches(3.0)
        table.columns[1].width = Inches(4.8)

    # =====================================
    # PROCESS ROWS
    # =====================================

    for i, row in enumerate(rows):

        while len(row) < max_cols:
            row.append("")

        # =================================
        # FIXED ROW HEIGHT
        # =================================

        tr = table.rows[i]._tr

        trPr = tr.get_or_add_trPr()

        trHeight = OxmlElement('w:trHeight')

        trHeight.set(
            qn('w:val'),
            "700"
        )

        trHeight.set(
            qn('w:hRule'),
            "atLeast"
        )

        trPr.append(trHeight)

        # =================================
        # PROCESS CELLS
        # =================================

        for j, cell_text in enumerate(row):

            cell_obj = table.cell(i, j)

            para = cell_obj.paragraphs[0]

            para.paragraph_format.space_before = Pt(2)

            para.paragraph_format.space_after = Pt(2)

            para.paragraph_format.line_spacing = 1.1

            cleaned_cell = cell_text.strip()

            # =============================
            # REMOVE OUTER [ ]
            # =============================

            if (
                cleaned_cell.startswith("[")
                and
                cleaned_cell.endswith("]")
            ):

                cleaned_cell = cleaned_cell[
                    1:-1
                ].strip()

            # =============================
            # REMOVE ESCAPED SPACES
            # =============================

            if (
                r"\frac" not in cleaned_cell
            ):

                cleaned_cell = cleaned_cell.replace(
                    r"\ ",
                    " "
                )

            # =============================
            # EQUATION DETECTION
            # =============================

            latex_equation = (

                r"\frac" in cleaned_cell

                or

                r"\sqrt" in cleaned_cell

                or

                r"\sum" in cleaned_cell

                or

                r"\int" in cleaned_cell

                or

                r"\beta" in cleaned_cell

                or

                r"\alpha" in cleaned_cell

                or

                r"\displaystyle" in cleaned_cell

                or

                r"\times" in cleaned_cell

                or

                r"\left" in cleaned_cell

                or

                r"\right" in cleaned_cell

                or

                "^" in cleaned_cell

                or

                "_" in cleaned_cell

            )
            # =============================
            # EQUATION CELL
            # =============================

            if latex_equation:

                para.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.CENTER
                )

                try:

                    cleaned_eq = clean_equation(
                        cleaned_cell
                    )

                    # =====================================
                    # TABLE-SPECIFIC CLEANUP ONLY
                    # =====================================

                    # REMOVE OUTER WRAPPER BRACKETS

                    if (
                        cleaned_eq.startswith("(")
                        and
                        cleaned_eq.endswith(")")
                    ):

                        cleaned_eq = cleaned_eq[
                            1:-1
                        ].strip()

                    # REMOVE DISPLAYSTYLE

                    cleaned_eq = cleaned_eq.replace(
                        r"\displaystyle",
                        ""
                    )

                    # REMOVE \left and \right

                    cleaned_eq = cleaned_eq.replace(
                        r"\left",
                        ""
                    )

                    cleaned_eq = cleaned_eq.replace(
                        r"\right",
                        ""
                    )

                    # REMOVE ESCAPED SPACES

                    cleaned_eq = cleaned_eq.replace(
                        r"\ ",
                        " "
                    )

                    # REMOVE DOUBLE SPACES

                    cleaned_eq = re.sub(
                        r"\s+",
                        " ",
                        cleaned_eq
                    ).strip()

                    # REMOVE DOUBLE BRACKETS

                    cleaned_eq = re.sub(
                        r'\(\s*\(',
                        '(',
                        cleaned_eq
                    )

                    cleaned_eq = re.sub(
                        r'\)\s*\)',
                        ')',
                        cleaned_eq
                    )

                    # -------------------------
                    # NESTED FRACTION CHECK
                    # -------------------------

                    frac_depth = cleaned_eq.count(
                        r"\frac"
                    )

                    # -------------------------
                    # SIMPLE EQUATIONS
                    # -------------------------

                    if frac_depth <= 8:

                        add_word_equation(
                            para,
                            cleaned_eq
                        )

                    # -------------------------
                    # COMPLEX EQUATION FALLBACK
                    # -------------------------

                    else:

                        fallback_eq = cleaned_eq

                        fallback_eq = fallback_eq.replace(
                            r"\frac",
                            "/"
                        )

                        fallback_eq = fallback_eq.replace(
                            "{",
                            "("
                        )

                        fallback_eq = fallback_eq.replace(
                            "}",
                            ")"
                        )

                        run = para.add_run(
                            fallback_eq
                        )

                        run.font.name = (
                            "Cambria Math"
                        )

                        run.font.size = Pt(11)

                        run.italic = True

                except Exception as e:

                    logger.error(
                        f"FAILED EQUATION: {cleaned_cell}"
                    )

                    logger.error(
                        f"Equation Error: {e}"
                    )

                    run = para.add_run(
                        cleaned_cell
                    )

                    run.font.name = (
                        "Courier New"
                    )

                    run.font.size = Pt(10)

            # =============================
            # NORMAL TEXT CELL
            # =============================

            else:

                normal_text = cleaned_cell

                normal_text = normal_text.replace(
                    r"\ ",
                    " "
                )

                if re.search(
                    r'[\d%₹,.]+',
                    normal_text
                ):

                    para.alignment = (
                        WD_PARAGRAPH_ALIGNMENT.CENTER
                    )

                else:

                    para.alignment = (
                        WD_PARAGRAPH_ALIGNMENT.LEFT
                    )

                match = re.search(
                    r'([A-Za-z]+)\(?_\{([^{}]+)\}\)?',
                    normal_text
                )

                if match:

                    main_text = match.group(1)

                    sub_text = match.group(2)

                    run_main = para.add_run(
                        main_text
                    )

                    run_main.font.name = (
                        "Times New Roman"
                    )

                    run_main.font.size = Pt(12)

                    run_sub = para.add_run(
                        sub_text
                    )

                    run_sub.font.subscript = True

                    run_sub.font.name = (
                        "Times New Roman"
                    )

                    run_sub.font.size = Pt(10)

                else:

                    run = para.add_run(
                        normal_text
                    )

                    run.font.name = (
                        "Times New Roman"
                    )

                    run.font.size = Pt(12)

                # =========================
                # HEADER ROW STYLING
                # =========================

                if i == 0:

                    for run in para.runs:

                        run.bold = True

                        run.font.color.rgb = (
                            RGBColor(
                                255,
                                255,
                                255
                            )
                        )

                    tcPr = (
                        cell_obj._tc
                        .get_or_add_tcPr()
                    )

                    shd = OxmlElement('w:shd')

                    shd.set(
                        qn('w:fill'),
                        '0A2540'
                    )

                    tcPr.append(shd)

    # =====================================
    # FINAL SPACING
    # =====================================

    doc.add_paragraph()

# =========================================
# MAIN ROUTE
# =========================================

@app.post("/generate-docx")

async def generate_docx(
    data: TextInput,
    background_tasks: BackgroundTasks
):

    text = remove_ai_garbage(
        data.text
    )

    # =========================================
    # NORMALIZE CHATGPT CONTENT
    # =========================================

    # IMPORTANT:
    # Protect explicit math blocks BEFORE
    # any line reconstruction takes place.

    protected_math_blocks = []

    def protect_math_block(match):

        block_id = f"__MSPK_MATH_BLOCK_{len(protected_math_blocks)}__"

        protected_math_blocks.append(
            match.group(0)
        )

        return block_id

    # Protect ```math ... ```
    text = re.sub(
        r"```(?:math|latex)\s*\n.*?\n```",
        protect_math_block,
        text,
        flags=re.IGNORECASE | re.DOTALL
    )

    # Protect $$ ... $$
    text = re.sub(
        r"\$\$\s*.*?\s*\$\$",
        protect_math_block,
        text,
        flags=re.DOTALL
    )

    # Protect \[ ... \]
    text = re.sub(
        r"\\\[\s*.*?\s*\\\]",
        protect_math_block,
        text,
        flags=re.DOTALL
    )

    lines = text.splitlines()

    lines = reconstruct_chatgpt_blocks(
        lines
    )

    # =========================================
    # RESTORE PROTECTED MATH BLOCKS
    # =========================================

    restored_lines = []

    for line in lines:

        restored = line

        for index, block in enumerate(
            protected_math_blocks
        ):

            placeholder = (
                f"__MSPK_MATH_BLOCK_{index}__"
            )

            if placeholder in restored:

                restored = restored.replace(
                    placeholder,
                    block
                )

        restored_lines.append(
            restored
        )

    lines = restored_lines

    doc = Document()

    section = doc.sections[0]

    section.top_margin = Inches(0.8)
    section.bottom_margin = Inches(0.8)

    section.left_margin = Inches(0.7)
    section.right_margin = Inches(0.7)

    add_page_border(doc)

    style = doc.styles["Normal"]

    style.font.name = "Times New Roman"
    style.font.size = Pt(11)

    # TITLE

    title = doc.add_paragraph()

    title.alignment = (
        WD_PARAGRAPH_ALIGNMENT.CENTER
    )

    title_run = title.add_run(
        data.title
    )

    title_run.bold = True

    title_run.font.size = Pt(16)

    title_run.font.color.rgb = (
        RGBColor(
            10,
            37,
            64
        )
    )

    title.paragraph_format.space_after = Pt(12)

    # PROCESS CONTENT

    i = 0

    while i < len(lines):

        line = lines[i].strip()

        # EMPTY

        if not line:

            i += 1

            continue

        # TABLE

        if is_table_line(line):

            table_lines = []

            while (

                i < len(lines)

                and

                is_table_line(lines[i])

            ):

                table_lines.append(
                    lines[i]
                )

                i += 1

            add_table(
                doc,
                table_lines
            )

            continue

        # =====================================
        # DISPLAY EQUATION BLOCK
        # =====================================

        clean_line = line.strip()

        # -------------------------------------
        # PROTECTED MARKDOWN MATH BLOCK
        # -------------------------------------

        if (
            "```math" in clean_line.lower()
            or
            "```latex" in clean_line.lower()
        ):

            equation = clean_line

            # Remove markdown fence
            equation = re.sub(
                r"```(?:math|latex)",
                "",
                equation,
                flags=re.IGNORECASE
            )

            equation = equation.replace(
                "```",
                ""
            )

            equation = equation.strip()

            if equation:

                p = doc.add_paragraph()

                p.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.CENTER
                )

                p.paragraph_format.space_before = Pt(3)

                p.paragraph_format.space_after = Pt(3)

                add_word_equation(
                    p,
                    equation
                )

            i += 1

            continue

        # -------------------------------------
        # $$ DISPLAY MATH BLOCK
        # -------------------------------------

        if clean_line.startswith("$$"):

            equation_lines = []

            first_line = clean_line[2:].strip()

            if first_line:

                if first_line.endswith("$$"):

                    equation_lines.append(
                        first_line[:-2].strip()
                    )

                else:

                    equation_lines.append(
                        first_line
                    )

            if not clean_line.endswith("$$"):

                i += 1

                while i < len(lines):

                    current = lines[i].strip()

                    if current.endswith("$$"):

                        last_line = (
                            current[:-2].strip()
                        )

                        if last_line:

                            equation_lines.append(
                                last_line
                            )

                        break

                    if current:

                        equation_lines.append(
                            current
                        )

                    i += 1

            equation = " ".join(
                equation_lines
            )

            if equation:

                p = doc.add_paragraph()

                p.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.CENTER
                )

                p.paragraph_format.space_before = Pt(3)

                p.paragraph_format.space_after = Pt(3)

                add_word_equation(
                    p,
                    equation
                )

            i += 1

            continue

        # -------------------------------------
        # \[ ... \] DISPLAY MATH BLOCK
        # -------------------------------------

        if (
            clean_line == r"\["
            or
            clean_line.startswith(r"\[")
        ):

            equation_lines = []

            first_line = clean_line[2:].strip()

            if first_line:

                if first_line.endswith(r"\]"):

                    equation_lines.append(
                        first_line[:-2].strip()
                    )

                else:

                    equation_lines.append(
                        first_line
                    )

            if not clean_line.endswith(r"\]"):

                i += 1

                while i < len(lines):

                    current = lines[i].strip()

                    if current.endswith(r"\]"):

                        last_line = (
                            current[:-2].strip()
                        )

                        if last_line:

                            equation_lines.append(
                                last_line
                            )

                        break

                    if current:

                        equation_lines.append(
                            current
                        )

                    i += 1

            equation = " ".join(
                equation_lines
            )

            if equation:

                p = doc.add_paragraph()

                p.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.CENTER
                )

                p.paragraph_format.space_before = Pt(3)

                p.paragraph_format.space_after = Pt(3)

                add_word_equation(
                    p,
                    equation
                )

            i += 1

            continue

        # -------------------------------------
        # BRACKET DISPLAY EQUATION
        # -------------------------------------

        if (
            clean_line == "["
            or
            clean_line.endswith("[")
        ):

            equation_lines = []

            first_line = clean_line[:-1].strip()

            if first_line:

                equation_lines.append(
                    first_line
                )

            i += 1

            while i < len(lines):

                current = lines[i].strip()

                if current == "]":

                    break

                if current:

                    equation_lines.append(
                        current
                    )

                i += 1

            equation = " ".join(
                equation_lines
            )

            if equation:

                p = doc.add_paragraph()

                p.alignment = (
                    WD_PARAGRAPH_ALIGNMENT.CENTER
                )

                p.paragraph_format.space_before = Pt(3)

                p.paragraph_format.space_after = Pt(3)

                add_word_equation(
                    p,
                    equation
                )

            i += 1

            continue

        # -------------------------------------
        # PURE EQUATION LINE ONLY
        # -------------------------------------

        if (

            is_equation_line(line)

            and

            len(line.split()) <= 20

            and

            "$" not in line

        ):

            p = doc.add_paragraph()

            p.alignment = (
                WD_PARAGRAPH_ALIGNMENT.CENTER
            )

            p.paragraph_format.space_before = Pt(2)

            p.paragraph_format.space_after = Pt(2)

            add_word_equation(
                p,
                line
            )

            i += 1

            continue

        # =====================================
        # PROFESSIONAL DIVIDER
        # =====================================

        if line.strip() == "---":

            p = doc.add_paragraph()

            run = p.add_run(
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
            )

            run.bold = True

            run.font.color.rgb = RGBColor(
                120,
                120,
                120
            )

            run.font.size = Pt(14)

            p.alignment = (
                WD_PARAGRAPH_ALIGNMENT.CENTER
            )

            p.paragraph_format.space_before = Pt(5)

            p.paragraph_format.space_after = Pt(5)

            i += 1

            continue

        # =====================================
        # HEADING
        # =====================================

        line_clean = clean_text(line)

        if is_heading(line_clean):

            p = doc.add_paragraph()

            run = p.add_run(
                line_clean
            )

            run.bold = True

            run.font.size = Pt(13)

            run.font.name = (
                "Times New Roman"
            )

            run.font.color.rgb = (
                RGBColor(
                    10,
                    37,
                    64
                )
            )

            p.paragraph_format.space_before = Pt(8)

            p.paragraph_format.space_after = Pt(3)

            p.paragraph_format.keep_with_next = True

            i += 1

            continue

        # =====================================
        # NORMAL TEXT
        # =====================================

        p = doc.add_paragraph()

        p.paragraph_format.line_spacing = 1.15

        p.paragraph_format.space_before = Pt(0)

        p.paragraph_format.space_after = Pt(3)

        clean_inline = line_clean

        # -------------------------------------
        # REMOVE MARKDOWN BOLD / ITALIC
        # -------------------------------------

        clean_inline = re.sub(
            r"\*\*(.*?)\*\*",
            r"\1",
            clean_inline
        )

        clean_inline = re.sub(
            r"(?<!\*)\*(?!\*)(.*?)\*(?<!\*)",
            r"\1",
            clean_inline
        )

        # -------------------------------------
        # REMOVE INLINE CODE MARKERS
        # -------------------------------------

        clean_inline = clean_inline.replace(
            "`",
            ""
        )

        # -------------------------------------
        # REMOVE DOLLAR WRAPPERS
        # -------------------------------------

        clean_inline = clean_inline.replace(
            "$",
            ""
        )

        # =====================================
        # INLINE LATEX SYMBOLS
        # =====================================

        clean_inline = clean_inline.replace(
            r"\beta",
            "β"
        )

        clean_inline = clean_inline.replace(
            r"\alpha",
            "α"
        )

        clean_inline = clean_inline.replace(
            r"\sigma",
            "σ"
        )

        clean_inline = clean_inline.replace(
            r"\theta",
            "θ"
        )

        clean_inline = clean_inline.replace(
            r"\lambda",
            "λ"
        )

        clean_inline = clean_inline.replace(
            r"\omega",
            "ω"
        )

        clean_inline = clean_inline.replace(
            r"\gamma",
            "γ"
        )

        clean_inline = clean_inline.replace(
            r"\delta",
            "δ"
        )

        clean_inline = clean_inline.replace(
            r"\mu",
            "μ"
        )

        clean_inline = clean_inline.replace(
            r"\rho",
            "ρ"
        )

        clean_inline = clean_inline.replace(
            r"\phi",
            "φ"
        )

        clean_inline = clean_inline.replace(
            r"\pi",
            "π"
        )

        clean_inline = clean_inline.replace(
            r"\infty",
            "∞"
        )

        clean_inline = clean_inline.replace(
            r"\times",
            "×"
        )

        clean_inline = clean_inline.replace(
            r"\div",
            "÷"
        )

        clean_inline = clean_inline.replace(
            r"\pm",
            "±"
        )

        clean_inline = clean_inline.replace(
            r"\geq",
            "≥"
        )

        clean_inline = clean_inline.replace(
            r"\leq",
            "≤"
        )

        clean_inline = clean_inline.replace(
            r"\neq",
            "≠"
        )

        clean_inline = clean_inline.replace(
            r"\approx",
            "≈"
        )

        clean_inline = clean_inline.replace(
            r"\%",
            "%"
        )

        # =====================================
        # SIMPLE INLINE FRACTIONS
        # =====================================

        clean_inline = re.sub(
            r'\\frac\{([^{}]+)\}\{([^{}]+)\}',
            r'(\1/\2)',
            clean_inline
        )

        # =====================================
        # INLINE SUBSCRIPTS
        # =====================================

        subscript_map = {
            "0": "₀",
            "1": "₁",
            "2": "₂",
            "3": "₃",
            "4": "₄",
            "5": "₅",
            "6": "₆",
            "7": "₇",
            "8": "₈",
            "9": "₉",
            "e": "ₑ",
            "m": "ₘ",
            "n": "ₙ",
            "i": "ᵢ",
            "j": "ⱼ",
            "t": "ₜ",
        }

        for key, value in subscript_map.items():

            clean_inline = clean_inline.replace(
                "_" + key,
                value
            )

        # =====================================
        # INLINE SUPERSCRIPTS
        # =====================================

        superscript_map = {
            "0": "⁰",
            "1": "¹",
            "2": "²",
            "3": "³",
            "4": "⁴",
            "5": "⁵",
            "6": "⁶",
            "7": "⁷",
            "8": "⁸",
            "9": "⁹",
            "n": "ⁿ",
            "t": "ᵗ",
        }

        for key, value in superscript_map.items():

            clean_inline = clean_inline.replace(
                "^" + key,
                value
            )

        # =====================================
        # REMOVE SIMPLE LATEX TEXT COMMANDS
        # =====================================

        clean_inline = clean_inline.replace(
            r"\text{",
            ""
        )

        # =====================================
        # REMOVE REMAINING LATEX COMMAND
        # BACKSLASHES WITHOUT DESTROYING
        # MATHEMATICAL SYMBOLS
        # =====================================

        clean_inline = re.sub(
            r'\\([a-zA-Z]+)',
            r'\1',
            clean_inline
        )

        # =====================================
        # REMOVE ONLY EMPTY / SIMPLE BRACES
        # =====================================

        clean_inline = re.sub(
            r'\{([^{}]*)\}',
            r'\1',
            clean_inline
        )

        # =====================================
        # NORMALIZE EXCESSIVE SPACES
        # =====================================

        clean_inline = re.sub(
            r'[ \t]+',
            ' ',
            clean_inline
        ).strip()

        # =====================================
        # ADD NORMAL TEXT
        # =====================================

        run = p.add_run(
            clean_inline
        )

        run.font.size = Pt(11)

        run.font.name = (
            "Times New Roman"
        )

        i += 1

    # =====================================
    # SAVE FILE
    # =====================================

    filename = (
        f"report_{uuid.uuid4()}.docx"
    )

    filepath = os.path.join(
        os.getcwd(),
        filename
    )

    doc.save(filepath)

    background_tasks.add_task(
        os.remove,
        filepath
    )

    return FileResponse(

        filepath,

        media_type=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),

        filename="MSPK_Report.docx"
    )

# =========================================
# HEALTH
# =========================================

@app.get("/health")

def health():

    return {
        "status": "ok"
    }

# =========================================
# HOME
# =========================================

@app.get("/")

def home():

    return {

        "status": "online",

        "message": "MSPK BACKEND RUNNING"

    }

# =========================================
# MAIN
# =========================================

if __name__ == "__main__":

    import uvicorn

    logger.info("=" * 60)

    logger.info(
        "MSPK AI to Word Converter – Backend v5.4"
    )

    logger.info("=" * 60)

    uvicorn.run(
        app,
        host="0.0.0.0",
        port=8000
    )