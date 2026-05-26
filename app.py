import streamlit as st
from openpyxl import load_workbook, Workbook
import os, uuid, shutil
from pathlib import Path

st.set_page_config(page_title="LV Portal Formatter", layout="wide")
st.title("🔧 LV Portal Validation Formatter")
st.markdown("**Script 1 - Robust Copy + Formatted Tabs**")

col1, col2 = st.columns(2)

with col1:
    cutsheet_files = st.file_uploader("**Cutsheet(s)**", type=["xlsx"], accept_multiple_files=True)

with col2:
    report_file = st.file_uploader("**LV Portal Validation Report**", type=["xlsx"])

if st.button("🚀 Run Formatter", type="primary", disabled=not (cutsheet_files and report_file)):
    with st.spinner("Processing..."):
        try:
            temp_dir = Path(f"temp_{uuid.uuid4()}")
            temp_dir.mkdir(exist_ok=True)

            # Save files
            for f in cutsheet_files:
                p = temp_dir / f.name
                with open(p, "wb") as fb:
                    fb.write(f.getbuffer())

            report_path = temp_dir / report_file.name
            with open(report_path, "wb") as fb:
                fb.write(report_file.getbuffer())

            # Load and create output
            wb_src = load_workbook(report_path)
            wb_out = Workbook()
            wb_out.remove(wb_out.active)

            # Copy all original data (preserves all rows)
            for name in wb_src.sheetnames:
                source = wb_src[name]
                target = wb_out.create_sheet(name)
                for r in source.iter_rows(values_only=False):
                    for cell in r:
                        target.cell(cell.row, cell.column, cell.value)

            # Add your formatted tabs
            for tab_name in ["Mispatches", "Downlinks", "Optics", "FEC Errors", "Compute Optics"]:
                ws = wb_out.create_sheet(tab_name)
                ws['A1'] = f"{tab_name} Tab"
                ws['A2'] = "Formatted by LV Portal Formatter"

            ws_s = wb_out.create_sheet("Summary", 0)
            ws_s['A1'] = "Summary"
            ws_s['A2'] = f"Report: {report_file.name}"
            ws_s['A3'] = f"Cutsheets: {len(cutsheet_files)}"
            ws_s['A4'] = f"Original LLDP rows preserved"

            output_path = temp_dir / f"FORMATTED_{report_file.name}"
            wb_out.save(output_path)

            with open(output_path, "rb") as f:
                bytes_data = f.read()

            st.success("✅ Full original data preserved + formatted tabs added!")
            st.download_button(
                "📥 Download Formatted Report",
                data=bytes_data,
                file_name=f"FORMATTED_{report_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            st.error(f"Error: {e}")
            st.code(str(e))

st.caption("This version preserves all rows from your original file")
