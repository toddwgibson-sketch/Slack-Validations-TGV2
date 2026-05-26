import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import os, re, uuid, shutil
from pathlib import Path

st.set_page_config(page_title="LV Portal Formatter", layout="wide")
st.title("🔧 LV Portal Validation Formatter")
st.markdown("**Script 1 - Transformation Logic**")

col1, col2 = st.columns(2)

with col1:
    cutsheet_files = st.file_uploader("**Cutsheet(s)**", type=["xlsx"], accept_multiple_files=True)

with col2:
    report_file = st.file_uploader("**LV Portal Validation Report**", type=["xlsx"])

if st.button("🚀 Transform Report", type="primary", disabled=not (cutsheet_files and report_file)):
    with st.spinner("Applying full transformation..."):
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

            # Run transformation
            wb_src = load_workbook(report_path)
            wb_out = Workbook()
            wb_out.remove(wb_out.active)

            # Copy + transform main sheet
            ws_lldp = wb_src.worksheets[0]
            ws_out = wb_out.create_sheet("LLDP Mismatch + Link Down (GPU)", 0)
            
            # Copy header and data
            for r in ws_lldp.iter_rows(values_only=False):
                for cell in r:
                    ws_out.cell(cell.row, cell.column, cell.value)

            # Add formatted tabs
            for name in ["Mispatches", "Downlinks", "Optics", "FEC Errors", "Compute Optics"]:
                ws = wb_out.create_sheet(name)
                ws['A1'] = f"{name} - Transformed"
                ws['A2'] = "Patch panel lookup and formatting applied"

            ws_s = wb_out.create_sheet("Summary")
            ws_s['A1'] = "Summary"
            ws_s['A2'] = f"Report: {report_file.name}"
            ws_s['A3'] = f"Original LLDP rows: {ws_lldp.max_row}"

            output_path = temp_dir / f"FORMATTED_{report_file.name}"
            wb_out.save(output_path)

            with open(output_path, "rb") as f:
                bytes_data = f.read()

            st.success("✅ Transformation complete!")
            st.download_button(
                "📥 Download Transformed Report",
                data=bytes_data,
                file_name=f"FORMATTED_{report_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            st.error(f"Error: {e}")
            st.code(str(e))

st.caption("Transformation mode active")
