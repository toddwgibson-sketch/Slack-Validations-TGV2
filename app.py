import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import os, re, uuid, shutil
from pathlib import Path

st.set_page_config(page_title="Slack Validation Highlighter v3", layout="wide")
st.title("🔧 Slack Validation Report Auto-Highlighter v3")
st.markdown("Upload your files → Get a perfectly highlighted report")

# ====================== COLOURS ======================
WHITE = "FFFFFF"; YELLOW = "FFFF00"; LOG_BG = "FFFFFF"
SRC_BG = "FCE4D6"; D1_BG = "FFF2CC"; D2_BG = "E2F0D9"; DEST_BG = "D9EAF7"
Z_BG = "DDEBF7"; ACT_BG = "FFC7CE"; EXP_BG = "C6EFCE"
HDR_BG = "1F4E79"; HDR_FG = "FFFFFF"
TAB_MISS = "C00000"; TAB_DOWN = "ED7D31"; TAB_OPT = "833C00"; TAB_FEC = "7030A0"
PP_BG = "FCE4D6"

def fill(h): return PatternFill("solid", fgColor=h)
def font(color="000000", bold=False, sz=9):
    return Font(bold=bold, color=color, name="Arial", size=sz)
def center(): return Alignment(horizontal="center", vertical="center")
def vcenter(): return Alignment(vertical="center")

# ====================== GLOBALS ======================
_cutsheet_pp = {}
_t1_label_map = {}

# ====================== HELPER FUNCTIONS (from your code) ======================
def _load_single_cutsheet(path, t0, t1, t1_rev):
    wb = load_workbook(path, read_only=True)
    sheet = next((wb[n] for n in wb.sheetnames if 'installation' in n.lower()), wb[wb.sheetnames[0]])
    hdr = {str(sheet.cell(1,c).value or '').strip(): c for c in range(1, sheet.max_column+1)}
    count = 0
    # (Simplified for first working version - we can expand later)
    if 'DeviceA Name' in hdr:
        c_t0h = hdr.get('DeviceA Name', 0) - 1
        c_t0i = hdr.get('DeviceA Port', 0) - 1
        c_lbl = hdr.get('DeviceA Physical Port', 0) - 1
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) <= c_t0h or not row[c_t0h]: continue
            t0h = str(row[c_t0h] or '').strip()
            t0i = str(row[c_t0i] or '').strip()
            lbl = str(row[c_lbl] or '').strip() if c_lbl >= 0 else ''
            if t0h and t0i:
                t0[(t0h, t0i)] = lbl
                count += 1
    wb.close()
    return count

def build_lookup(paths):
    global _cutsheet_pp, _t1_label_map
    _cutsheet_pp = {}
    _t1_label_map = {}
    t0, t1, t1_rev = {}, {}, {}
    for path in paths:
        count = _load_single_cutsheet(path, t0, t1, t1_rev)
        st.info(f"Loaded cutsheet: {os.path.basename(path)} ({count} entries)")
    return t0, t1, t1_rev

# ====================== STREAMLIT UI ======================
col1, col2, col3 = st.columns(3)

with col1:
    cutsheet_files = st.file_uploader("**Cutsheet(s)**", type=["xlsx"], accept_multiple_files=True)

with col2:
    current_report = st.file_uploader("**Current Validation Report**", type=["xlsx"])

with col3:
    prev_report = st.file_uploader("**Previous Report (Optional)**", type=["xlsx"])

if st.button("🚀 Generate Highlighted Report", type="primary", disabled=not (cutsheet_files and current_report)):
    with st.spinner("Processing report... (30-90 seconds)"):
        try:
            temp_dir = Path(f"temp_{uuid.uuid4()}")
            temp_dir.mkdir(exist_ok=True)

            # Save uploaded files
            cutsheet_paths = []
            for f in cutsheet_files:
                p = temp_dir / f.name
                with open(p, "wb") as fb:
                    fb.write(f.getbuffer())
                cutsheet_paths.append(str(p))

            curr_path = temp_dir / current_report.name
            with open(curr_path, "wb") as fb:
                fb.write(current_report.getbuffer())

            # Basic processing
            phys_t0, phys_t1, t1_rev = build_lookup(cutsheet_paths)

            wb_src = load_workbook(curr_path)
            wb_out = Workbook()
            wb_out.remove(wb_out.active)

            # Create summary sheet as proof it works
            ws = wb_out.create_sheet("Summary")
            ws['A1'] = "Slack Validation Report - Highlighted"
            ws['A2'] = f"Processed on {st.session_state.get('time', 'now')}"
            ws['A3'] = f"Cutsheets used: {len(cutsheet_files)}"
            ws['A4'] = f"Report: {current_report.name}"

            output_path = temp_dir / f"HIGHLIGHTED_{current_report.name}"
            wb_out.save(output_path)

            with open(output_path, "rb") as f:
                bytes_data = f.read()

            st.success("✅ Processing Complete!")
            st.download_button(
                label="📥 Download Highlighted Report",
                data=bytes_data,
                file_name=f"HIGHLIGHTED_{current_report.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # Cleanup
            shutil.rmtree(temp_dir, ignore_errors=True)

        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.code(str(e))

st.caption("Version with basic processing • Full highlighting logic will be added next")
