# app.py — Column Grouping Merger
# CSV/XLSX/XLS • optional "treat sheets as files"
# Groups = map header variants → one standardized output column

import os, glob, json, time, traceback, platform, subprocess, re
from collections import Counter, defaultdict
from datetime import datetime
import pandas as pd
import streamlit as st
from rapidfuzz import process, fuzz  # Make sure rapidfuzz is in requirements.txt

# ---------- Page & dirs ----------
st.set_page_config(page_title="Column Grouping Merger", layout="wide")

APP_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_INPUT_DIR = os.path.join(APP_DIR, "input")
DEFAULT_OUTPUT_DIR = os.path.join(APP_DIR, "output")
GROUPS_DIR = os.path.join(APP_DIR, "groups")
DEFAULT_MARKER = os.path.join(GROUPS_DIR, "_default.txt")
os.makedirs(DEFAULT_INPUT_DIR, exist_ok=True)
os.makedirs(DEFAULT_OUTPUT_DIR, exist_ok=True)
os.makedirs(GROUPS_DIR, exist_ok=True)

def add_log(msg: str):
    st.session_state.setdefault("logs", []).append(msg)

# ---------- Folder picker ----------
def open_folder_dialog(prompt_text: str = "Select a folder") -> str:
    system = platform.system()
    if system == "Darwin":
        script = f'''
            tell application "Finder" to activate
            set theFolder to choose folder with prompt "{prompt_text}"
            POSIX path of theFolder
        '''
        try:
            out = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, check=True)
            return (out.stdout or "").strip()
        except Exception:
            pass
        script2 = f'''
            tell application "System Events" to activate
            set theFolder to choose folder with prompt "{prompt_text}"
            POSIX path of theFolder
        '''
        try:
            out = subprocess.run(["osascript", "-e", script2], capture_output=True, text=True, check=True)
            return (out.stdout or "").strip()
        except Exception:
            return ""
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk(); root.withdraw()
            try: root.attributes("-topmost", True)
            except Exception: pass
            path = filedialog.askdirectory(title=prompt_text)
            root.destroy()
            return path or ""
        except Exception:
            return ""

# ---------- IO helpers ----------
def is_valid_dir(path: str) -> bool:
    try: return bool(path) and os.path.isdir(path)
    except Exception: return False

def list_data_files(folder: str, recursive: bool = False):
    if not is_valid_dir(folder): return []
    pats = ["*.csv", "*.xlsx", "*.xls"]
    files = []
    if recursive:
        for p in pats: files.extend(glob.glob(os.path.join(folder, "**", p), recursive=True))
    else:
        for p in pats: files.extend(glob.glob(os.path.join(folder, p)))
    return sorted(set(files))

def is_csv(p): return p.lower().endswith(".csv")
def is_xls(p): return p.lower().endswith(".xls")
def is_xlsx(p): return p.lower().endswith(".xlsx")

def unit_key(file_path: str, sheet: str | None) -> str:
    return f"{file_path}::{sheet}" if sheet else file_path

def unit_label(file_path: str, sheet: str | None, base_dir: str | None) -> str:
    base = os.path.relpath(file_path, base_dir) if base_dir else os.path.basename(file_path)
    return f"{base} :: {sheet}" if sheet else base

def split_unit_key(key: str):
    return key.split("::", 1) if "::" in key else (key, None)

def read_headers_csv(path: str) -> list[str]:
    try:
        df = pd.read_csv(path, nrows=0, dtype=str, low_memory=False)
        return list(map(str, df.columns))
    except Exception as e:
        add_log(f"[ERROR] CSV header read failed: {os.path.basename(path)} -> {e}")
        return []

def read_headers_excel(path: str, sheet: str | None) -> list[str]:
    try:
        df = pd.read_excel(path, nrows=0, sheet_name=sheet) if sheet else pd.read_excel(path, nrows=0)
        return list(map(str, df.columns))
    except Exception as e:
        add_log(f"[ERROR] Excel header read failed: {os.path.basename(path)}[{sheet}] -> {e}")
        return []

def list_excel_sheets(path: str) -> list[str]:
    try:
        xls = pd.ExcelFile(path)
        return list(xls.sheet_names)
    except Exception as e:
        add_log(f"[WARN] Could not list sheets for {os.path.basename(path)}: {e}")
        return []

def read_full_unit(path: str, sheet: str | None) -> pd.DataFrame:
    if is_csv(path): return pd.read_csv(path, dtype=None, low_memory=False)
    elif is_xls(path): return pd.read_excel(path, engine="xlrd")
    else: return pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)

# ---------- Coverage / maps ----------
def analyze_headers_by_unit(headers_by_unit: dict[str, list[str]], included_units: set[str]):
    if not headers_by_unit or not included_units:
        return pd.DataFrame(), {}
    total_units = len(included_units)
    units_with_header = defaultdict(set)
    total_occ = Counter()
    anomalies = {}
    for uk, hdrs in headers_by_unit.items():
        if uk not in included_units: continue
        c = Counter(hdrs)
        anomalies[uk] = {h: cnt for h, cnt in c.items() if cnt > 1}
        for h, cnt in c.items():
            if cnt >= 1:
                units_with_header[h].add(uk)
                total_occ[h] += cnt
    rows = []
    for h, unit_set in units_with_header.items():
        cnt = len(unit_set); pct = round((cnt/total_units)*100, 1)
        rows.append({"header": h, "units_with_header": cnt, "coverage_pct": pct, "total_occurrences": total_occ[h]})
    cov = pd.DataFrame(rows).sort_values(by=["coverage_pct","units_with_header","header"], ascending=[False,False,True]).reset_index(drop=True)
    return cov, anomalies

def header_to_units_map(headers_by_unit: dict[str, list[str]], included_units: set[str]) -> dict[str, set[str]]:
    h2u = defaultdict(set)
    for uk, hdrs in headers_by_unit.items():
        if uk not in included_units: continue
        for h in hdrs: h2u[h].add(uk)
    return h2u

def find_original_match_in_unit(headers_in_unit: list[str], variants: list[str]) -> str | None:
    s = set(headers_in_unit)
    for v in variants:
        if v in s: return v
        for h in headers_in_unit:
            if h.startswith(f"{v}."): return h
    return None

# ---------- Presets ----------
GROUPS_VERSION = 7
def preset_path(name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip()
    return os.path.join(GROUPS_DIR, f"{safe}.json")

def list_presets() -> list[str]:
    return [os.path.splitext(f)[0] for f in sorted(os.listdir(GROUPS_DIR)) if f.endswith(".json")]

def save_preset(name: str, header_groups: list, notes: str = "") -> str:
    data = {"name": name, "created_utc": datetime.utcnow().isoformat()+"Z", "header_groups": header_groups, "notes": notes, "version": GROUPS_VERSION}
    path = preset_path(name)
    with open(path, "w", encoding="utf-8") as f: json.dump(data, f, indent=2)
    return path

def load_preset(name: str) -> dict | None:
    path = preset_path(name)
    if not os.path.exists(path): return None
    try:
        with open(path, "r", encoding="utf-8") as f: return json.load(f)
    except Exception: return None

def delete_preset(name: str) -> bool:
    path = preset_path(name)
    try:
        if os.path.exists(path): os.remove(path); return True
    except Exception: pass
    return False

def get_default_preset_name() -> str | None:
    if os.path.exists(DEFAULT_MARKER):
        try:
            with open(DEFAULT_MARKER, "r", encoding="utf-8") as f: nm = f.read().strip()
            return nm or None
        except Exception: return None
    return None

def set_default_preset_name(name: str | None) -> None:
    if name is None:
        try:
            if os.path.exists(DEFAULT_MARKER): os.remove(DEFAULT_MARKER)
        except Exception: pass
        return
    with open(DEFAULT_MARKER, "w", encoding="utf-8") as f: f.write(name)

# ---------- Session ----------
defaults = {
    "input_dir": DEFAULT_INPUT_DIR, "output_dir": DEFAULT_OUTPUT_DIR, "recursive": False,
    "treat_sheets_as_files": False, "output_name": "merged_output.csv",
    "units": [], "unit_meta": {}, "headers_by_unit": {}, "headers_unique_by_unit": {},
    "anomalies_by_unit": {}, "coverage_df": pd.DataFrame(),
    "selected_unit": None, "included_units": set(),
    "header_groups": [], "current_group_headers": [], "current_group_name": "",
    "validation_report": pd.DataFrame(), "validation_ready": False, "proceed_with_missing": False,
    "preset_selected": None, "preset_name_to_save": "", "preset_note": "", "autoloaded_default": False,
    "last_scan_signature": None, "logs": []
}
for k, v in defaults.items():
    if k not in st.session_state: st.session_state[k] = v

# ---------- Top bar ----------
st.title("Column Grouping Merger")

# Only keep the refresh button here
c1, c2 = st.columns([1, 1])
with c1:
    refresh = st.button("🔄 Rescan")

# ---------- Utilities ----------
with st.expander("Utilities", expanded=False):
    c1, c2, c3, c4 = st.columns([3, 3, 2, 2])
    with c1:
        st.session_state.input_dir = st.text_input("Input folder", value=st.session_state.input_dir, placeholder=DEFAULT_INPUT_DIR, key="util_input_dir")
    with c2:
        st.session_state.output_dir = st.text_input("Output folder", value=st.session_state.output_dir, placeholder=DEFAULT_OUTPUT_DIR, key="util_output_dir")
    with c3:
        st.session_state.output_name = st.text_input("Output file name", value=st.session_state.output_name, key="util_output_name")
    with c4:
        st.session_state.treat_sheets_as_files = st.toggle("Treat Excel sheets as files", value=st.session_state.treat_sheets_as_files, key="util_treat_sheets_as_files")
    c5, c6 = st.columns([1, 1])
    with c5:
        if st.button("📂 Pick Input Folder", key="util_pick_input"):
            p = open_folder_dialog("Select the INPUT folder")
            if p: st.session_state.input_dir = p
    with c6:
        if st.button("📁 Pick Output Folder", key="util_pick_output"):
            p = open_folder_dialog("Select the OUTPUT folder")
            if p: st.session_state.output_dir = p
    if st.button("Generate 5 Dummy Excel Files in Input Folder"):
        try:
            generate_dummy_excels(st.session_state.input_dir)
            st.success("Dummy Excel files created in input folder!")
        except Exception as e:
            st.error(f"Failed to generate dummy files: {e}")

# ---------- Dummy Data Generator ----------
def generate_dummy_excels(target_folder):
    import numpy as np
    import pandas as pd
    os.makedirs(target_folder, exist_ok=True)
    # Define some column name variants
    col_sets = [
        ["Name", "Age", "Score", "Email", "JoinDate"],
        ["Full Name", "Years", "Score", "Email Address", "Join Date"],
        ["Name", "Age", "Result", "Email", "Joined"],
        ["Full Name", "Age", "Score", "Contact", "JoinDate"],
        ["Name", "Years", "Score", "Email", "Join Date"]
    ]
    for i, cols in enumerate(col_sets):
        data = {}
        for col in cols:
            if "Name" in col:
                data[col] = [f"User {j+1}" for j in range(20)]
            elif "Full Name" in col:
                data[col] = [f"Person {j+1}" for j in range(20)]
            elif "Age" in col or "Years" in col:
                data[col] = np.random.randint(18, 60, 20)
            elif "Score" in col or "Result" in col:
                data[col] = np.random.randint(50, 100, 20)
            elif "Email" in col:
                data[col] = [f"user{j+1}@example.com" for j in range(20)]
            elif "Contact" in col:
                data[col] = [f"contact{j+1}@example.com" for j in range(20)]
            elif "Join" in col:
                data[col] = pd.date_range("2024-01-01", periods=20).strftime("%Y-%m-%d").tolist()
            else:
                data[col] = [f"Data{j+1}" for j in range(20)]
        df = pd.DataFrame(data)
        out_path = os.path.join(target_folder, f"dummy_{i+1}.xlsx")
        df.to_excel(out_path, index=False)

# ---------- Scan ----------
def compute_scan_signature(path: str, recursive: bool, sheets_as_files: bool):
    if not is_valid_dir(path): return None
    files = list_data_files(path, recursive)
    mtimes = []
    for f in files:
        try: mtimes.append(os.path.getmtime(f))
        except Exception: mtimes.append(0)
    return f"{path}|{recursive}|{sheets_as_files}|{len(files)}|{sum(mtimes)}"

def rescan_and_read():
    if st.session_state.output_dir and not os.path.isdir(st.session_state.output_dir):
        os.makedirs(st.session_state.output_dir, exist_ok=True)

    files = list_data_files(st.session_state.input_dir, st.session_state.recursive)
    units, unit_meta = [], {}
    headers_by_unit, headers_unique_by_unit = {}, {}

    for fp in files:
        if is_csv(fp):
            uk = unit_key(fp, None)
            units.append(uk); unit_meta[uk] = {"file": fp, "sheet": None}
            hdrs = read_headers_csv(fp)
            headers_by_unit[uk] = hdrs
            headers_unique_by_unit[uk] = sorted(set(hdrs), key=lambda x: x.lower())
        else:
            if st.session_state.treat_sheets_as_files:
                sheets = list_excel_sheets(fp) or [None]
                for sh in sheets:
                    uk = unit_key(fp, sh)
                    units.append(uk); unit_meta[uk] = {"file": fp, "sheet": sh}
                    hdrs = read_headers_excel(fp, sh)
                    headers_by_unit[uk] = hdrs
                    headers_unique_by_unit[uk] = sorted(set(hdrs), key=lambda x: x.lower())
            else:
                uk = unit_key(fp, None)
                units.append(uk); unit_meta[uk] = {"file": fp, "sheet": None}
                hdrs = read_headers_excel(fp, None)
                headers_by_unit[uk] = hdrs
                headers_unique_by_unit[uk] = sorted(set(hdrs), key=lambda x: x.lower())

    st.session_state.units = units
    st.session_state.unit_meta = unit_meta
    st.session_state.headers_by_unit = headers_by_unit
    st.session_state.headers_unique_by_unit = headers_unique_by_unit
    st.session_state.included_units = set(units)

    cov, anoms = analyze_headers_by_unit(headers_by_unit, st.session_state.included_units)
    st.session_state.coverage_df = cov
    st.session_state.anomalies_by_unit = anoms
    st.session_state.last_scan_signature = compute_scan_signature(
        st.session_state.input_dir, st.session_state.recursive, st.session_state.treat_sheets_as_files
    )
    add_log(f"Scanned {len(units)} unit(s). Coverage headers: {len(cov)}.")

sig_now = compute_scan_signature(st.session_state.input_dir, st.session_state.recursive, st.session_state.treat_sheets_as_files)
if refresh or (sig_now is not None and sig_now != st.session_state.last_scan_signature):
    try: rescan_and_read()
    except Exception as e:
        st.error(f"Scan failed: {e}"); add_log(f"[ERROR] Scan failed: {e}\n{traceback.format_exc()}")

# ---------- Units view ----------
with st.expander("File and Sheet Selection and Preview", expanded=False):
    left, right = st.columns([1.2, 2.8], gap="large")
    with left:
        if st.session_state.units:
            labels = [unit_label(st.session_state.unit_meta[u]["file"], st.session_state.unit_meta[u]["sheet"], st.session_state.input_dir)
                      for u in st.session_state.units]
            idx_map = {i:u for i,u in enumerate(st.session_state.units)}
            default_idx = st.session_state.units.index(st.session_state.selected_unit) if st.session_state.selected_unit in st.session_state.units else 0
            sel_idx = st.radio("Detected unit(s)", options=list(range(len(labels))),
                               format_func=lambda i: labels[i], index=default_idx if len(labels)>0 else 0)
            st.session_state.selected_unit = idx_map[sel_idx]

            st.markdown("**Included units** (affects coverage, validation, merge):")
            chosen = st.multiselect("Which units are part of this merge?", options=labels, default=labels)
            st.session_state.included_units = {st.session_state.units[labels.index(lbl)] for lbl in chosen} if chosen else set()
        else:
            st.info("No files found. Check the input folder and Rescan (🔄).")
    with right:
        uk = st.session_state.selected_unit
        if uk:
            uniq = st.session_state.headers_unique_by_unit.get(uk, [])
            meta = st.session_state.unit_meta[uk]
            st.markdown(f"**Unique headers in:** `{unit_label(meta['file'], meta['sheet'], st.session_state.input_dir)}`")
            st.dataframe(pd.DataFrame({"header": uniq}), use_container_width=True)

# ---------- Coverage ----------
with st.expander("Header Preview (across selected files)", expanded=False):
    cov = st.session_state.coverage_df
    if not cov.empty:
        q = st.text_input("Filter headers (contains)", "")
        view = cov.copy()
        if q: view = view[view["header"].str.contains(q, case=False, na=False)]
        st.dataframe(view, use_container_width=True)
    else:
        st.caption("Coverage will appear after scanning and including units.")

# ---------- Mapping Builder (one output column/group at a time) ----------
def get_unit_label(uk):
    meta = st.session_state.unit_meta[uk]
    return unit_label(meta["file"], meta["sheet"], st.session_state.input_dir)

def get_headers_for_unit(uk):
    return st.session_state.headers_unique_by_unit.get(uk, [])

def fuzzy_auto_match(pattern, headers):
    if not pattern or not headers:
        return None
    match, score, _ = process.extractOne(pattern, headers, scorer=fuzz.token_sort_ratio)
    return match if score > 60 else None  # Adjust threshold as needed

with st.expander("Mapping Builder (one output column/group at a time)", expanded=True):
    # State for new group
    if "new_group_name" not in st.session_state:
        st.session_state.new_group_name = ""
    if "new_group_pattern" not in st.session_state:
        st.session_state.new_group_pattern = ""
    if "new_group_selections" not in st.session_state or not isinstance(st.session_state.new_group_selections, dict):
        st.session_state.new_group_selections = {}

    # Name and Auto-match fields
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.session_state.new_group_name = st.text_input(
            "Output column name (group)",
            value=st.session_state.new_group_name,
            key="new_group_name_input"
        )
    with col2:
        st.session_state.new_group_pattern = st.text_input(
            "Auto-match pattern (optional)",
            value=st.session_state.new_group_pattern,
            key="new_group_pattern_input"
        )
    with col3:
        st.markdown("<div style='height:1.7em'></div>", unsafe_allow_html=True)  # vertical align
        if st.button("Apply Auto-Match"):
            for uk in st.session_state.included_units:
                headers = get_headers_for_unit(uk)
                match = fuzzy_auto_match(st.session_state.new_group_pattern, headers)
                if match:
                    st.session_state.new_group_selections[uk] = match
                else:
                    st.session_state.new_group_selections[uk] = ""
            st.success("Auto-match applied.")
            st.rerun()

    # --- Interactive Table with Dropdowns ---
    
    # Calculate coverage based on current selections
    coverage_count = sum(1 for uk in st.session_state.included_units if st.session_state.new_group_selections.get(uk))
    total_units = len(st.session_state.included_units)
    percent = int((coverage_count / total_units) * 100) if total_units else 0

    # Coverage bar and count
    st.markdown(
        f"<div style='margin-top:1em; margin-bottom:0.2em;font-size:0.95em;'>"
        f"<b>Coverage:</b> {coverage_count} of {total_units} files/units selected "
        f"({percent}%)"
        "</div>",
        unsafe_allow_html=True
    )
    st.progress(percent / 100)

    # Table Headers
    st.markdown(
        """
        <style>
        .st-emotion-cache-1y4p8pa { /* Targets the container of st.columns */
            gap: 0.5rem; /* Reduces gap between columns */
        }
        .header-style {
            font-weight: 600;
            padding: 0.25rem 0;
            border-bottom: 1px solid #444;
        }
        .row-style {
            padding: 0.1rem 0;
            vertical-align: middle;
        }
        </style>
        """,
        unsafe_allow_html=True
    )
    header_cols = st.columns([2, 2])
    header_cols[0].markdown('<div class="header-style">File</div>', unsafe_allow_html=True)
    header_cols[1].markdown('<div class="header-style">Selected Column</div>', unsafe_allow_html=True)

    # Table Rows with Interactive Dropdowns
    for uk in sorted(st.session_state.included_units):
        row_cols = st.columns([2, 2])
        label = get_unit_label(uk)
        headers = get_headers_for_unit(uk)
        current = st.session_state.new_group_selections.get(uk, "")
        
        with row_cols[0]:
            st.markdown(f'<div class="row-style">{label}</div>', unsafe_allow_html=True)
        
        with row_cols[1]:
            st.selectbox(
                label=f"selectbox_for_{uk}", # Hidden label for unique widget ID
                label_visibility="collapsed",
                options=[""] + headers,
                index=([""] + headers).index(current) if current in headers else 0,
                key=f"new_group_select_{uk}"
            )

    # Save group button
    if st.button("💾 Save group"):
        if not st.session_state.new_group_name.strip():
            st.error("Please enter a group/output column name.")
        elif coverage_count == 0:
            st.error("Select at least one column to map for this group.")
        else:
            selected_headers = [v for v in st.session_state.new_group_selections.values() if v]
            st.session_state.header_groups.append({
                "output_name": st.session_state.new_group_name.strip(),
                "headers": list(set(selected_headers))
            })
            st.session_state.new_group_name = ""
            st.session_state.new_group_pattern = ""
            st.session_state.new_group_selections = {}
            st.success("Saved header group.")
            st.rerun()

# ---------- Saved groups (expand for full mapping table) ----------
with st.expander("Saved Header Groups", expanded=False):
    if st.session_state.header_groups and st.session_state.included_units:
        for i, grp in enumerate(st.session_state.header_groups):
            with st.expander(f"{i+1}. {grp['output_name']} — {len(grp['headers'])} header variant(s)", expanded=False):
                grp["output_name"] = st.text_input(f"Rename output column (group {i+1})", value=grp["output_name"], key=f"rename_{i}")
                # Build mapping table: Original Column | Output (Group) | File | Sheet
                rows = []
                for uk in sorted(st.session_state.included_units):
                    meta = st.session_state.unit_meta[uk]
                    hdrs = st.session_state.headers_by_unit.get(uk, [])
                    match = find_original_match_in_unit(hdrs, grp["headers"])
                    rows.append({
                        "Original Column": match if match else "",
                        "Output (Group)": grp["output_name"],
                        "File": os.path.relpath(meta["file"], st.session_state.input_dir) if st.session_state.input_dir else os.path.basename(meta["file"]),
                        "Sheet": meta["sheet"] or ""
                    })
                df = pd.DataFrame(rows).sort_values(by=["File","Sheet"]).reset_index(drop=True)
                st.dataframe(df, use_container_width=True)

                st.markdown("**Group variants (headers in this group):**")
                st.dataframe(pd.DataFrame({"headers": grp["headers"]}), use_container_width=True)

                if st.button(f"Remove group {i+1}", key=f"remove_{i}"):
                    st.session_state.header_groups.pop(i)
                    st.rerun()
    else:
        st.caption("No groups yet, or no included units.")

# ---------- Validate ----------
with st.expander("Validate before merging", expanded=False):
    def validate_groups() -> pd.DataFrame:
        rows = []
        inc = st.session_state.included_units or set()
        for grp in st.session_state.header_groups:
            variants = list(grp["headers"])
            for uk in inc:
                meta = st.session_state.unit_meta[uk]
                hdrs = st.session_state.headers_by_unit.get(uk, [])
                match = find_original_match_in_unit(hdrs, variants)
                rows.append({
                    "file": os.path.relpath(meta["file"], st.session_state.input_dir) if st.session_state.input_dir else os.path.basename(meta["file"]),
                    "sheet": meta["sheet"] or "",
                    "group": grp["output_name"],
                    "original_header": match if match else "",
                    "covered": bool(match)
                })
        df = pd.DataFrame(rows)
        if not df.empty: df = df.sort_values(by=["file","sheet","group"]).reset_index(drop=True)
        return df

    val1, val2 = st.columns([1.2,2.8])
    with val1:
        do_validate = st.button("🧪 Validate")
    with val2:
        if do_validate:
            st.session_state.validation_report = validate_groups()
            st.session_state.validation_ready = True
            st.session_state.proceed_with_missing = False
        if st.session_state.validation_ready and not st.session_state.validation_report.empty:
            rep = st.session_state.validation_report.copy()
            rep["status"] = rep["covered"].map(lambda x: "OK" if x else "MISSING")
            def style_missing(s): return ['background-color: #ffeaea' if v=="MISSING" else '' for v in s]
            st.dataframe(rep[["file","sheet","group","original_header","status"]].style.apply(style_missing, subset=["status"]), use_container_width=True)
            if (rep["covered"]==False).any():
                st.warning("Some units have no matching column for one or more groups. Those cells will be blank in the merged file.")
                st.session_state.proceed_with_missing = st.checkbox("Understood — proceed with blanks where needed.")
            else:
                st.success("All groups are present for all included units. Ready to merge!")

# ---------- Merge ----------
def resolve_columns_in_df(df: pd.DataFrame, header: str) -> list[str]:
    exact = [c for c in df.columns if str(c) == header]
    dotted = [c for c in df.columns if str(c).startswith(f"{header}.")]
    return exact + dotted

merge_now = st.button("✅ Merge & Export")
if merge_now:
    if not st.session_state.header_groups:
        st.error("Create at least one header group before merging.")
    elif not is_valid_dir(st.session_state.output_dir):
        st.error("Please set a valid output folder.")
    else:
        need_confirm = False
        if not st.session_state.validation_report.empty and (st.session_state.validation_report["covered"]==False).any():
            need_confirm = True
        if need_confirm and not st.session_state.proceed_with_missing:
            st.error("Run Validate and check the confirmation box to proceed with blanks, or adjust your groups.")
        else:
            inc = list(st.session_state.included_units or st.session_state.units)
            add_log(f"Starting merge for {len(inc)} unit(s)..."); start = time.time()

            dfs = {}
            for uk in inc:
                fpath, sheet = split_unit_key(uk)
                try:
                    dfs[uk] = read_full_unit(fpath, sheet)
                except Exception as e:
                    add_log(f"[ERROR] Failed reading {unit_label(fpath, sheet, st.session_state.input_dir)}: {e}")

            merged_parts = []
            for uk, df in dfs.items():
                out_cols = {}
                for grp in st.session_state.header_groups:
                    variants = grp["headers"]
                    matches = []
                    for h in variants:
                        matches.extend(resolve_columns_in_df(df, h))
                    seen = set(); uniq = []
                    for m in matches:
                        if m in df.columns and m not in seen:
                            uniq.append(m); seen.add(m)
                    if uniq:
                        ser = pd.Series([pd.NA]*len(df))
                        for c in uniq: ser = ser.fillna(df[c])
                        out_cols[grp["output_name"]] = ser
                    else:
                        out_cols[grp["output_name"]] = pd.Series([pd.NA]*len(df))
                part = pd.DataFrame(out_cols)
                fpath, sheet = split_unit_key(uk)
                part.insert(0, "source_unit", unit_label(fpath, sheet, st.session_state.input_dir))
                merged_parts.append(part)

            merged_df = pd.concat(merged_parts, ignore_index=True) if merged_parts else pd.DataFrame()
            elapsed = time.time() - start
            add_log(f"Merge finished in {elapsed:.2f}s. Rows total: {len(merged_df)}.")

            if not merged_df.empty:
                out_path = os.path.join(st.session_state.output_dir, st.session_state.output_name)
                try:
                    merged_df.to_csv(out_path, index=False)
                    add_log(f"[SAVE] Wrote CSV: {out_path}")
                    st.success(f"CSV saved: {out_path}")
                    with open(out_path, "rb") as f:
                        st.download_button("Download CSV", data=f.read(), file_name=os.path.basename(out_path))
                except Exception as e:
                    st.error(f"Failed to save CSV: {e}"); add_log(f"[ERROR] Save failed: {e}")
            else:
                st.warning("No data produced. Check groups and units, then try again.")

# ---------- Logs ----------
with st.expander("Logs", expanded=False):
    if st.session_state.get("logs"): st.code("\n".join(st.session_state["logs"]), language="text")
    else: st.write("No logs yet.")