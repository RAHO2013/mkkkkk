import re
from io import BytesIO
import pandas as pd
import pdfplumber
import streamlit as st

CODE_RE = re.compile(r"\b([A-Z]\d{3}[A-Z0-9]{2})\b")
ROW_START_RE = re.compile(r"^\s*(\d{1,4})\s+([A-Z]\d{3}[A-Z0-9]{2})\b\s*(.*)$")
FOOTER_PREFIXES = (
    "Downloaded Date:", "KARNATAKA EXAMINATIONS AUTHORITY", "ADMISSION TO",
    "FIRST ROUND OPTIONS LIST", "Mere applying", "Candidature is",
    "the provisional admission", "I hereby", "Signature of"
)

MASTER_CONSOLE_SCRIPT = r'(() => {\n  const data = prompt("Paste Excel Console Lines here (example: E005EC 1):");\n  if (!data) {\n    alert("No data pasted.");\n    return;\n  }\n\n  const updates = Object.fromEntries(\n    data.trim().split(/\\n+/).map(line => {\n      const [code, no] = line.trim().split(/\\s+/);\n      return [code, Number(no)];\n    })\n  );\n\n  const inputs = [...document.querySelectorAll(\'input.oe-opno-input[aria-label^="Option number for "]\')];\n  const result = [];\n\n  for (const [code, newNo] of Object.entries(updates)) {\n    const input = inputs.find(i => i.getAttribute("aria-label") === `Option number for ${code}`);\n\n    if (!input) {\n      result.push({ code, status: "NOT FOUND", newNo });\n      continue;\n    }\n\n    const oldNo = input.value;\n    input.value = String(newNo);\n    input.dispatchEvent(new Event("input", { bubbles: true }));\n    input.dispatchEvent(new Event("change", { bubbles: true }));\n    result.push({ code, oldNo, newNo, status: "CHANGED" });\n  }\n\n  const allAfter = [...document.querySelectorAll(\'input.oe-opno-input[aria-label^="Option number for "]\')]\n    .map(input => ({\n      code: input.getAttribute("aria-label").replace("Option number for ", ""),\n      no: Number(input.value)\n    }))\n    .filter(x => x.no !== 0);\n\n  const seen = new Map();\n  const duplicates = [];\n  for (const item of allAfter) {\n    if (seen.has(item.no)) duplicates.push({ optionNo: item.no, codes: [seen.get(item.no), item.code] });\n    else seen.set(item.no, item.code);\n  }\n\n  console.table(result);\n  if (duplicates.length) {\n    console.warn("Duplicate option numbers found:", duplicates);\n    alert("Duplicate option numbers found. Check console before clicking Update Options.");\n  } else {\n    alert("Completed. Check all option numbers manually, then click Update Options.");\n  }\n})();'

def clean_text(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def parse_full_row(option_no, code, full, page_no):
    full = clean_text(full)
    fee_match = re.search(r"(\d{1,2},\d{2},\d{3}|\d{2},\d{3})\s*-\s*", full)
    if fee_match:
        course_name = clean_text(full[:fee_match.start()])
        rest = clean_text(full[fee_match.start():])
        fee_end = re.search(r"Rupees Only", rest, flags=re.I)
        if fee_end:
            course_fee = clean_text(rest[:fee_end.end()])
            college_name = clean_text(rest[fee_end.end():])
        else:
            course_fee = rest
            college_name = ""
    else:
        course_name = full
        course_fee = ""
        college_name = ""
    return {
        "Old Option No": option_no,
        "New Option No": option_no,
        "College Course Code": code,
        "Course Name": course_name,
        "Course Fee": course_fee,
        "College Name": college_name,
        "Page": page_no,
        "Raw Text": clean_text(f"{option_no} {code} {full}"),
    }

def extract_options_from_pdf(pdf_file):
    rows = []
    with pdfplumber.open(pdf_file) as pdf:
        for page_no, page in enumerate(pdf.pages, start=1):
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            current = None

            def flush_current():
                nonlocal current
                if current:
                    full = clean_text(current["after"] + " " + " ".join(current["extra"]))
                    rows.append(parse_full_row(current["old"], current["code"], full, current["page"]))
                    current = None

            for raw_line in text.splitlines():
                line = clean_text(raw_line)
                if not line:
                    continue
                m = ROW_START_RE.match(line)
                if m:
                    flush_current()
                    current = {"old": int(m.group(1)), "code": m.group(2), "after": m.group(3), "extra": [], "page": page_no}
                else:
                    if current and not any(line.startswith(p) for p in FOOTER_PREFIXES):
                        current["extra"].append(line)
            flush_current()
    if not rows:
        return pd.DataFrame()
    df = pd.DataFrame(rows)
    df = df.sort_values("Old Option No").drop_duplicates(subset=["Old Option No", "College Course Code"], keep="first")
    return df

def make_excel(df):
    output = BytesIO()
    df = df.copy()
    df["Duplicate Check"] = df["New Option No"].apply(lambda x: "DELETE" if int(x) == 0 else "OK")
    df["Console Line"] = df["College Course Code"] + " " + df["New Option No"].astype(int).astype(str)

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Options_Reorder")
        wb = writer.book
        ws = writer.sheets["Options_Reorder"]
        header_fmt = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1E3A5F", "align": "center", "valign": "vcenter", "text_wrap": True})
        wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        for col, name in enumerate(df.columns):
            ws.write(0, col, name, header_fmt)
        widths = [12, 12, 18, 38, 34, 55, 8, 70, 16, 18]
        for i, w in enumerate(widths):
            ws.set_column(i, i, w, wrap)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, len(df), len(df.columns)-1)
        # Replace Duplicate Check and Console Line columns with formulas
        dup_col = df.columns.get_loc("Duplicate Check")
        line_col = df.columns.get_loc("Console Line")
        for r in range(1, len(df)+1):
            excel_row = r + 1
            ws.write_formula(r, dup_col, f'=IF(B{excel_row}=0,"DELETE",IF(COUNTIF($B:$B,B{excel_row})>1,"DUPLICATE","OK"))')
            ws.write_formula(r, line_col, f'=C{excel_row}&" "&B{excel_row}')
        ws2 = wb.add_worksheet("Master_Console_Code")
        ws2.write(0, 0, "Paste this full code in browser console. It will ask for Console Lines.", header_fmt)
        ws2.write(1, 0, MASTER_CONSOLE_SCRIPT, wrap)
        ws2.set_column(0, 0, 120)
        ws2.set_row(1, 420)
        ws3 = wb.add_worksheet("Instructions")
        instructions = [
            "1. Edit only New Option No column in Options_Reorder.",
            "2. Use 0 only when you want to delete an option.",
            "3. Duplicate Check must show OK or DELETE only.",
            "4. Copy Console Line column.",
            "5. Paste Master Console Code in KEA page console and paste the console lines when asked.",
            "6. Verify manually and click official Update Options button.",
            "7. Download final KEA option report and check again.",
        ]
        ws3.write(0, 0, "Instructions", header_fmt)
        for i, item in enumerate(instructions, start=1):
            ws3.write(i, 0, item, wrap)
        ws3.set_column(0, 0, 100)
    return output.getvalue()

st.set_page_config(page_title="KEA PDF to Excel Full Setup", layout="wide")
st.title("KEA PDF to Excel Reorder Tool")
st.write("Upload KEA option report PDF. Edit New Option No, download Excel, and use generated console lines.")
file = st.file_uploader("Upload KEA Option Report PDF", type=["pdf"])
if file:
    df = extract_options_from_pdf(file)
    if df.empty:
        st.error("No option rows found. This may be a scanned/image-only PDF.")
    else:
        st.success(f"Extracted {len(df)} options")
        edited = st.data_editor(
            df[["Old Option No", "New Option No", "College Course Code", "Course Name", "Course Fee", "College Name", "Page"]],
            use_container_width=True,
            num_rows="fixed",
            disabled=["Old Option No", "College Course Code", "Course Name", "Course Fee", "College Name", "Page"],
            column_config={"New Option No": st.column_config.NumberColumn(min_value=0, max_value=9999, step=1)},
        )
        edited["New Option No"] = edited["New Option No"].astype(int)
        duplicates = edited[edited["New Option No"].ne(0) & edited.duplicated("New Option No", keep=False)]
        if not duplicates.empty:
            st.warning("Duplicate option numbers found. Fix before using console lines.")
            st.dataframe(duplicates[["New Option No", "College Course Code", "Course Name"]], use_container_width=True)
        console_lines = "\n".join((edited["College Course Code"] + " " + edited["New Option No"].astype(str)).tolist())
        st.subheader("Console Lines")
        st.text_area("Copy these lines", console_lines, height=260)
        st.subheader("Master Console Script")
        st.text_area("Paste this once in KEA browser console", MASTER_CONSOLE_SCRIPT, height=260)
        st.download_button(
            "Download Excel With Code",
            data=make_excel(edited),
            file_name="KEA_Option_Reorder_With_Code.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
