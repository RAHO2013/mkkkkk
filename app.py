import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import pdfplumber
import streamlit as st

CODE_RE = re.compile(r"\b([A-Z]\d{3}[A-Z]{2})\b")
ROW_START_RE = re.compile(r"^\s*(\d{1,4})\s+([A-Z]\d{3}[A-Z]{2})\b\s*(.*)$")


def clean_text(s: str) -> str:
    if s is None:
        return ""
    return re.sub(r"\s+", " ", str(s)).strip()


def split_course_fee_college(text_after_code: str):
    """Best-effort split for KEA option PDF parsed text."""
    t = clean_text(text_after_code)
    # Fees commonly start like: 1,20,320 - One Lakh... OR 47,100 - Forty...
    fee_match = re.search(r"(\d{1,2},\d{2},\d{3}|\d{2},\d{3})\s*-\s*", t)
    if not fee_match:
        return t, "", ""

    course_name = clean_text(t[:fee_match.start()])
    rest = t[fee_match.start():]

    # Fee text ends before college name. Most college names start with known words/patterns.
    # This is intentionally conservative; user can still see full text in Raw Text column.
    college_start = re.search(
        r"\b(R\. V\.|Dayananda|M S Ramaiah|REVA|THE CHANAKYA|PES|Malnad|Jawaharlal|P E S|St\.Joseph|Alva|SDM|Mangalore|Sahyadri|S J C|The National|Vidya|ATME|Bangalore|Atria|B M S|B N M|Govt|University|Cauvery|Cambridge|Canara)",
        rest,
    )
    if college_start:
        fee = clean_text(rest[:college_start.start()])
        college = clean_text(rest[college_start.start():])
    else:
        # Fallback: keep full rest as fee/detail
        fee = clean_text(rest)
        college = ""

    return course_name, fee, college


def extract_options_from_pdf(pdf_file) -> pd.DataFrame:
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            for line in text.splitlines():
                line = clean_text(line)
                m = ROW_START_RE.match(line)
                if not m:
                    continue
                option_no = int(m.group(1))
                code = m.group(2)
                if not CODE_RE.fullmatch(code):
                    continue
                after = m.group(3)
                course_name, fee, college_name = split_course_fee_college(after)
                rows.append({
                    "Old Option No": option_no,
                    "New Option No": option_no,
                    "College Course Code": code,
                    "Course Name": course_name,
                    "Course Fee": fee,
                    "College Name": college_name,
                    "Raw Text": line,
                    "Page": page_no,
                })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df = df.sort_values("Old Option No").drop_duplicates(subset=["Old Option No", "College Course Code"], keep="first")
    df["Console Line"] = df["College Course Code"] + " " + df["New Option No"].astype(str)
    return df


def make_excel(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Options")

        ws = writer.book["Options"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        widths = {
            "A": 14, "B": 14, "C": 18, "D": 42, "E": 32, "F": 55,
            "G": 70, "H": 10, "I": 18, "J": 20
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        # Add helper formulas after dataframe columns
        max_row = len(df) + 1
        ws.cell(row=1, column=10, value="Duplicate Check")
        ws.cell(row=1, column=11, value="Console Line")
        for r in range(2, max_row + 1):
            ws.cell(row=r, column=10, value=f'=IF(B{r}=0,"DELETE",IF(COUNTIF($B:$B,B{r})>1,"DUPLICATE","OK"))')
            ws.cell(row=r, column=11, value=f'=C{r}&" "&B{r}')

        # Console script sheet
        script = """(() => {\n  const data = prompt(\"Paste Excel Console Lines here:\");\n  if (!data) { alert(\"No data pasted.\"); return; }\n\n  const updates = Object.fromEntries(\n    data.trim().split(/\\n+/).map(line => {\n      const [code, no] = line.trim().split(/\\s+/);\n      return [code, Number(no)];\n    })\n  );\n\n  const inputs = [...document.querySelectorAll('input.oe-opno-input[aria-label^=\"Option number for \"]')];\n  const result = [];\n\n  for (const [code, newNo] of Object.entries(updates)) {\n    const input = inputs.find(i => i.getAttribute(\"aria-label\") === `Option number for ${code}`);\n    if (!input) { result.push({ code, status: \"NOT FOUND\", newNo }); continue; }\n    const oldNo = input.value;\n    input.value = String(newNo);\n    input.dispatchEvent(new Event(\"input\", { bubbles: true }));\n    input.dispatchEvent(new Event(\"change\", { bubbles: true }));\n    result.push({ code, oldNo, newNo, status: \"CHANGED\" });\n  }\n\n  console.table(result);\n  alert(\"Completed. Check all option numbers, then click Update Options manually.\");\n})();"""
        pd.DataFrame({"Master Console Script": [script]}).to_excel(writer, index=False, sheet_name="Console Script")
        writer.book["Console Script"].column_dimensions["A"].width = 120

        guide = pd.DataFrame({
            "Steps": [
                "Upload KEA option PDF in this app.",
                "Open Options sheet and edit only New Option No column.",
                "Use 0 only if you want to delete that option.",
                "Check Duplicate Check column; it must show OK or DELETE only.",
                "Copy Console Line column and paste into the prompt opened by the master console script.",
                "Run on live KEA page, verify manually, then click Update Options.",
                "Download final KEA option report and verify again."
            ]
        })
        guide.to_excel(writer, index=False, sheet_name="Instructions")
        writer.book["Instructions"].column_dimensions["A"].width = 100

    return output.getvalue()


st.set_page_config(page_title="KEA PDF to Excel Reorder Tool", layout="wide")
st.title("KEA PDF to Excel Reorder Tool")
st.write("Upload KEA option report PDF. The app creates an Excel file for fast rearranging and console-line generation.")

uploaded = st.file_uploader("Upload KEA Option Report PDF", type=["pdf"])

if uploaded:
    df = extract_options_from_pdf(uploaded)
    if df.empty:
        st.error("No KEA option rows found. Try another PDF or check if the PDF is scanned/image-only.")
    else:
        st.success(f"Extracted {len(df)} options")
        edited = st.data_editor(
            df[["Old Option No", "New Option No", "College Course Code", "Course Name", "Course Fee", "College Name", "Page"]],
            num_rows="fixed",
            use_container_width=True,
            disabled=["Old Option No", "College Course Code", "Course Name", "Course Fee", "College Name", "Page"],
            column_config={"New Option No": st.column_config.NumberColumn(min_value=0, max_value=9999, step=1)},
        )

        edited["Console Line"] = edited["College Course Code"] + " " + edited["New Option No"].astype(int).astype(str)
        duplicates = edited[edited["New Option No"].ne(0) & edited.duplicated("New Option No", keep=False)]
        if not duplicates.empty:
            st.warning("Duplicate option numbers found. Fix before using console lines.")
            st.dataframe(duplicates[["New Option No", "College Course Code", "Course Name"]], use_container_width=True)

        st.subheader("Console Lines")
        console_lines = "\n".join(edited["Console Line"].tolist())
        st.text_area("Copy these lines", console_lines, height=220)

        excel_bytes = make_excel(edited)
        st.download_button(
            "Download Excel Tool",
            data=excel_bytes,
            file_name="KEA_Option_Reorder_Tool.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
