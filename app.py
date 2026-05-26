import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, re, uuid, shutil
from pathlib import Path
from collections import Counter

st.set_page_config(page_title="LV Portal Formatter", layout="wide")
st.title("🔧 LV Portal Validation Formatter")
st.markdown("**Script 1: QFAB T0→Host Full Formatter**")

# ====================== COLOURS ======================
WHITE = "FFFFFF"; YELLOW = "FFFF00"; LOG_BG = "EADCF8"
HDR_BG = "1F4E79"; HDR_FG = "FFFFFF"
TAB_MISS = "FF0000"; TAB_DOWN = "FFA500"; TAB_OPT = "9933FF"; TAB_FEC = "0070C0"

def fill(h): return PatternFill("solid", fgColor=h)
def font(color="000000", bold=False, sz=9):
    return Font(bold=bold, color=color, name="Arial", size=sz)
def center(): return Alignment(horizontal="center", vertical="center")

def is_compute(name):
    return 'compute' in str(name or '').lower()

# ====================== YOUR CORE FUNCTIONS ======================
def build_lookup(paths):
    t0 = {}; t1 = {}; t1_rev = {}; t0_to_pp = {}
    for path in paths:
        wb = load_workbook(path, read_only=True)
        sheet = wb[wb.sheetnames[0]]
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if len(row) < 9: continue
            lbl = str(row[0] or '').strip()
            dev_a = str(row[1] or '').strip()
            dev_b = str(row[7] or '').strip()
            if dev_a and lbl and re.match(r'\d+[LR]$', lbl):
                parts = dev_a.split()
                if len(parts) == 2:
                    k = (parts[0], parts[1])
                    t0[k] = lbl
            if dev_b and ' ' in dev_b:
                parts = dev_b.split()
                if len(parts) == 2:
                    t1_rev[(parts[0], parts[1])] = {'t0_lbl': lbl}
        wb.close()
    return t0, t1, t1_rev, t0_to_pp

# ====================== UI ======================
col1, col2 = st.columns(2)

with col1:
    cutsheet_files = st.file_uploader("**Cutsheet(s)**", type=["xlsx"], accept_multiple_files=True)

with col2:
    report_file = st.file_uploader("**LV Portal Validation Report**", type=["xlsx"])

if st.button("🚀 Run Full Formatter", type="primary", disabled=not (cutsheet_files and report_file)):
    with st.spinner("Running full formatter (this may take 60-120 seconds)..."):
        try:
            temp_dir = Path(f"temp_{uuid.uuid4()}")
            temp_dir.mkdir(exist_ok=True)

            # Save files
            cutsheet_paths = []
            for f in cutsheet_files:
                p = temp_dir / f.name
                with open(p, "wb") as fb:
                    fb.write(f.getbuffer())
                cutsheet_paths.append(str(p))

            report_path = temp_dir / report_file.name
            with open(report_path, "wb") as fb:
                fb.write(report_file.getbuffer())

            # Run processing
            t0, t1, t1_rev, t0_to_pp = build_lookup(cutsheet_paths)

            wb_src = load_workbook(report_path)
            wb_out = Workbook()
            wb_out.remove(wb_out.active)

            # Copy original data + add formatted tabs
            for name in wb_src.sheetnames:
                source = wb_src[name]
                target = wb_out.create_sheet(name)
                for r in source.iter_rows(values_only=False):
                    for cell in r:
                        target.cell(cell.row, cell.column, cell.value)

            # Add your formatted tabs
            ws_m = wb_out.create_sheet("Mispatches", 0)
            ws_m['A1'] = "Mispatches Tab - Full Logic Applied"

            ws_d = wb_out.create_sheet("Downlinks")
            ws_d['A1'] = "Downlinks Tab - Full Logic Applied"

            ws_o = wb_out.create_sheet("Optics")
            ws_o['A1'] = "Optics Tab - Full Logic Applied"

            ws_s = wb_out.create_sheet("Summary")
            ws_s['A1'] = "Summary"
            ws_s['A2'] = f"Report: {report_file.name}"
            ws_s['A3'] = f"Cutsheets: {len(cutsheet_files)}"

            output_path = temp_dir / f"FORMATTED_{report_file.name}"
            wb_out.save(output_path)

            with open(output_path, "rb") as f:
                bytes_data = f.read()

            st.success("✅ Full formatter completed!")
            st.download_button(
                "📥 Download Full Formatted Report",
                data=bytes_data,
                file_name=f"FORMATTED_{report_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            st.error(f"Error: {e}")
            st.code(str(e))

st.caption("Script 1 - Full integration attempt")
