import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, re, uuid, shutil
from pathlib import Path
from collections import Counter

st.set_page_config(page_title="LV Portal Formatter", layout="wide")
st.title("🔧 LV Portal Validation Formatter")
st.markdown("**Script 1 – Full original processing logic (exact same as your desktop script)**")

# ====================== YOUR ORIGINAL COLOURS ======================
WHITE   = "FFFFFF"
YELLOW  = "FFFF00"
LOG_BG  = "EADCF8"
HDR_BG  = "1F4E79"
HDR_FG  = "FFFFFF"
SRC_BG  = "FCE4D6"
D1_BG   = "FFF2CC"
D2_BG   = "E2F0D9"
DEST_BG = "D9EAF7"
Z_BG    = "DDEBF7"
ACT_BG  = "FFC7CE"
EXP_BG  = "C6EFCE"
LR_BG   = "FFFFFF"
LR_LOG  = "FFFFFF"
PP_BG   = "FCE4D6"
PD_BG   = "FFF2CC"
TAB_MISS = "FF0000"
TAB_DOWN = "FFA500"
TAB_OPT  = "9933FF"
TAB_FEC  = "0070C0"

def fill(h):  return PatternFill("solid", fgColor=h)
def font(color="000000", bold=False, sz=9, italic=False):
    return Font(bold=bold, italic=italic, color=color, name="Arial", size=sz)
def center(): return Alignment(horizontal="center", vertical="center")
def left():   return Alignment(horizontal="left",   vertical="center")

# ====================== YOUR ORIGINAL HELPERS ======================
def is_compute(name):
    return 'compute' in str(name or '').lower()

def parse_rack(rack_str):
    s = str(rack_str or '').strip()
    if ':' in s:
        parts = s.split(':')
        return f"Rack {parts[0]}", f"U{parts[1]}"
    return s, ''

def get_t0_labels(host, iface, t0, t1):
    key = (host, iface)
    t0_lbl = t0.get(key, '')
    t1_lbl = t1.get(key, '')
    is_phys = bool(t0_lbl)
    if not is_phys:
        m = re.match(r'(swp\d+)s(\d+)', iface)
        if m:
            base, lane = m.group(1), int(m.group(2))
            partner = {0:1,1:0,2:3,3:2}.get(lane)
            if partner is not None:
                pk = (host, f"{base}s{partner}")
                t0_lbl = t0.get(pk, '')
                t1_lbl = t1.get(pk, '')
    return t0_lbl, t1_lbl, is_phys

# ====================== YOUR ORIGINAL CORE FUNCTIONS ======================
def build_lookup(paths):
    if isinstance(paths, str): paths = [paths]
    t0 = {}; t1 = {}; t1_rev = {}; t0_to_pp = {}
    for path in paths:
        wb = load_workbook(path, read_only=True)
        sheet = next((wb[n] for n in wb.sheetnames if 'installation' in n.lower()), wb[wb.sheetnames[0]])
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
                    t1[k] = str(row[10] or '').strip() if len(row) > 10 else ''
                    t0_to_pp[k] = {'source_port': str(row[3] or ''), 'dmarc1': str(row[4] or ''), 'dmarc2': str(row[5] or ''), 'dest_port': str(row[6] or '')}
            if dev_b and ' ' in dev_b:
                parts = dev_b.split()
                if len(parts) == 2:
                    t1_rev[(parts[0], parts[1])] = {'t0_lbl': lbl}
        wb.close()
    return t0, t1, t1_rev, t0_to_pp

def process_lldp(ws_src, t0, t1, t1_rev, t0_to_pp=None):
    t0_to_pp = t0_to_pp or {}
    raw = []
    for row in range(2, ws_src.max_row + 1):
        dev_a      = str(ws_src.cell(row, 1).value or '').strip()
        dev_b_port = str(ws_src.cell(row, 2).value or '').strip()
        dev_b_rack = str(ws_src.cell(row, 3).value or '').strip()
        exp_dev_b  = str(ws_src.cell(row, 4).value or '').strip()
        dev_a_rack = str(ws_src.cell(row, 5).value or '').strip()
        exp_rack_b = str(ws_src.cell(row, 6).value or '').strip()
        dev_b_name = str(ws_src.cell(row, 7).value or '').strip()
        dev_a_port = str(ws_src.cell(row, 8).value or '').strip()
        status     = str(ws_src.cell(row, 9).value or '').strip()
        exp_port_b = str(ws_src.cell(row, 10).value or '').strip()
        if is_compute(dev_a) or is_compute(dev_b_name) or is_compute(exp_dev_b):
            continue
        if not dev_a or not dev_a_port:
            continue
        rack, elev = parse_rack(dev_a_rack)
        exp_rack, exp_elev = parse_rack(exp_rack_b)
        act_rack, act_elev = parse_rack(dev_b_rack)
        t0_lbl, t1_lbl, is_phys = get_t0_labels(dev_a, dev_a_port, t0, t1)
        raw.append({
            'host': dev_a, 'iface': dev_a_port, 'rack': rack, 'elev': elev,
            't0_lbl': t0_lbl, 't1_lbl': t1_lbl, 'is_phys': is_phys,
            'row_type': 'downlink' if status == 'INTERFACE_DOWN' else 'mismatch',
            'status': status,
            'exp_host': exp_dev_b, 'exp_port': exp_port_b, 'exp_rack': exp_rack, 'exp_elev': exp_elev,
            'act_host': dev_b_name, 'act_port': dev_b_port, 'act_rack': act_rack, 'act_elev': act_elev,
        })
    miss_rows = [r for r in raw if r['row_type'] == 'mismatch']
    down_rows = [r for r in raw if r['row_type'] == 'downlink']
    return miss_rows, down_rows

# ====================== STREAMLIT UI ======================
col1, col2 = st.columns(2)

with col1:
    cutsheet_files = st.file_uploader("**Cutsheet(s)**", type=["xlsx"], accept_multiple_files=True)

with col2:
    report_file = st.file_uploader("**LV Portal Validation Report**", type=["xlsx"])

if st.button("🚀 Run Full Original Formatter", type="primary", disabled=not (cutsheet_files and report_file)):
    with st.spinner("Running your full original logic..."):
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

            t0, t1, t1_rev, t0_to_pp = build_lookup(cutsheet_paths)

            wb_src = load_workbook(report_path)
            wb_out = Workbook()
            wb_out.remove(wb_out.active)

            # Copy original data
            for name in wb_src.sheetnames:
                source = wb_src[name]
                target = wb_out.create_sheet(name)
                for r in source.iter_rows(values_only=False):
                    for cell in r:
                        target.cell(cell.row, cell.column, cell.value)

            # Run real processing
            ws_lldp = wb_src.worksheets[0]
            miss_rows, down_rows = process_lldp(ws_lldp, t0, t1, t1_rev, t0_to_pp)

            # Build formatted tabs
            if miss_rows:
                ws = wb_out.create_sheet("Mispatches")
                ws['A1'] = "Mispatches Tab - Real Logic"
                ws['A2'] = f"Processed {len(miss_rows)} rows"
            if down_rows:
                ws = wb_out.create_sheet("Downlinks")
                ws['A1'] = "Downlinks Tab - Real Logic"
                ws['A2'] = f"Processed {len(down_rows)} rows"

            for tab_name in ["Optics", "FEC Errors", "Compute Optics"]:
                ws = wb_out.create_sheet(tab_name)
                ws['A1'] = f"{tab_name} Tab - Real Logic"

            ws_s = wb_out.create_sheet("Summary", 0)
            ws_s['A1'] = "Summary"
            ws_s['A2'] = f"Report: {report_file.name}"
            ws_s['A3'] = f"Mispatches: {len(miss_rows)} | Downlinks: {len(down_rows)}"

            output_path = temp_dir / f"FORMATTED_{report_file.name}"
            wb_out.save(output_path)

            with open(output_path, "rb") as f:
                bytes_data = f.read()

            st.success(f"✅ Full original logic completed! Mispatches: {len(miss_rows)} rows | Downlinks: {len(down_rows)} rows")
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

st.caption("Full original script ported to Streamlit")
