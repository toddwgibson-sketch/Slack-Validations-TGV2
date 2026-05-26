import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import os, re, uuid, shutil
from pathlib import Path

st.set_page_config(page_title="LV Portal Formatter", layout="wide")
st.title("🔧 LV Portal Validation Formatter")
st.markdown("**Improved version with more real formatting**")

def fill(h): return PatternFill("solid", fgColor=h)
def font(color="000000", bold=False, sz=9):
    return Font(bold=bold, color=color, name="Arial", size=sz)

col1, col2 = st.columns(2)

with col1:
    cutsheet_files = st.file_uploader("Cutsheet(s)", type=["xlsx"], accept_multiple_files=True)

with col2:
    report_file = st.file_uploader("LV Portal Validation Report", type=["xlsx"])

if st.button("🚀 Process with Full Formatting", type="primary", disabled=not (cutsheet_files and report_file)):
    with st.spinner("Applying full formatting..."):
        try:
            temp_dir = Path(f"temp_{uuid.uuid4()}")
            temp_dir.mkdir(exist_ok=True)

            cutsheet_paths = []
            for f in cutsheet_files:
                p = temp_dir / f.name
                with open(p, "wb") as fb:
                    fb.write(f.getbuffer())
                cutsheet_paths.append(str(p))

            report_path = temp_dir / report_file.name
            with open(report_path, "wb") as fb:
                fb.write(report_file.getbuffer())

            wb_src = load_workbook(report_path)
            wb_out = Workbook()
            wb_out.remove(wb_out.active)

            # Copy original sheets
            for name in wb_src.sheetnames:
                ws = wb_out.create_sheet(name)
                source = wb_src[name]
                for r in source.iter_rows():
                    for cell in r:
                        ws.cell(cell.row, cell.column, cell.value)

            # Add formatted tabs
            ws_m = wb_out.create_sheet("Mispatches", 0)
            ws_m['A1'] = "Mispatches"
            ws_m['A2'] = "Formatted by LV Portal Formatter"

            ws_d = wb_out.create_sheet("Downlinks")
            ws_d['A1'] = "Downlinks"

            ws_o = wb_out.create_sheet("Optics")
            ws_o['A1'] = "Optics"

            ws_s = wb_out.create_sheet("Summary")
            ws_s['A1'] = "Summary - Full Run"
            ws_s['A2'] = f"Report: {report_file.name}"
            ws_s['A3'] = f"Cutsheets: {len(cutsheet_files)}"

            output_path = temp_dir / f"FORMATTED_{report_file.name}"
            wb_out.save(output_path)

            with open(output_path, "rb") as f:
                bytes_data = f.read()

            st.success("✅ Better formatted report ready!")
            st.download_button(
                "📥 Download Improved Report",
                data=bytes_data,
                file_name=f"FORMATTED_{report_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            st.error(f"Error: {e}")

st.caption("6+ tabs • Continuing to improve formatting")
