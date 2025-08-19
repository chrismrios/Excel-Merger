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
    "add_source_cols": False,
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

# ---------- Autoload Default Preset ----------
if not st.session_state.autoloaded_default:
    default_preset = get_default_preset_name()
    if default_preset:
        preset_data = load_preset(default_preset)
        if preset_data:
            st.session_state.header_groups = preset_data.get("header_groups", [])
            add_log(f"Autoloaded default preset: {default_preset}")
    st.session_state.autoloaded_default = True # Ensure this runs only once per session

# ---------- Top bar ----------
st.title("Column Grouping Merger")

# Only keep the refresh button here
c1, c2 = st.columns([1, 1])
with c1:
    refresh = st.button("🔄 Rescan")

# ---------- Utilities ----------
with st.expander("Utilities", expanded=False):
    c1, c2, c3 = st.columns(3)
    with c1:
        st.session_state.input_dir = st.text_input("Input folder", value=st.session_state.input_dir, placeholder=DEFAULT_INPUT_DIR, key="util_input_dir")
    with c2:
        st.session_state.output_dir = st.text_input("Output folder", value=st.session_state.output_dir, placeholder=DEFAULT_OUTPUT_DIR, key="util_output_dir")
    with c3:
        st.session_state.output_name = st.text_input("Output file name", value=st.session_state.output_name, key="util_output_name")
    
    c4, c5 = st.columns(2)
    with c4:
        if st.button("📂 Pick Input Folder", key="util_pick_input"):
            p = open_folder_dialog("Select the INPUT folder")
            if p: st.session_state.input_dir = p
    with c5:
        if st.button("📁 Pick Output Folder", key="util_pick_output"):
            p = open_folder_dialog("Select the OUTPUT folder")
            if p: st.session_state.output_dir = p

    st.markdown("---")
    c6, c7 = st.columns(2)
    with c6:
        st.session_state.treat_sheets_as_files = st.toggle("Treat Excel sheets as separate files", value=st.session_state.treat_sheets_as_files, key="util_treat_sheets_as_files")
    with c7:
        st.session_state.add_source_cols = st.toggle("Add source file column for each group", value=st.session_state.add_source_cols, key="util_add_source_cols")

    if st.button("Generate 5 Dummy Excel Files in Input Folder"):
        try:
            generate_dummy_excels(st.session_state.input_dir)
            st.toast("✅ Dummy Excel files created in input folder!", icon="🎉")
        except Exception as e:
            st.toast(f"Failed to generate dummy files: {e}", icon="🚨")

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

# ---------- Group Presets ----------
with st.expander("Group Presets", expanded=False):
    st.markdown("#### Load & Manage Presets")
    presets = list_presets()
    
    # This 'if/else' block handles the case where no presets exist
    if not presets:
        st.caption("No saved presets found. Use the form below to save the current groups as a new preset.")
    else:
        # Find index of default preset for the selectbox
        default_preset_name = get_default_preset_name()
        try:
            default_idx = presets.index(default_preset_name) if default_preset_name in presets else 0
        except ValueError:
            default_idx = 0

        st.session_state.preset_selected = st.selectbox(
            "Select a preset to load or manage",
            options=presets,
            index=default_idx,
            key="preset_selector"
        )

        p_c1, p_c2, p_c3, p_c4 = st.columns(4)
        with p_c1:
            if st.button("📂 Load Selected Preset"):
                preset_data = load_preset(st.session_state.preset_selected)
                if preset_data:
                    st.session_state.header_groups = preset_data.get("header_groups", [])
                    st.toast(f"Loaded preset: {st.session_state.preset_selected}", icon="✅")
                    st.rerun()
                else:
                    st.toast(f"Failed to load preset: {st.session_state.preset_selected}", icon="🚨")
        with p_c2:
            if st.button("📠 Duplicate Preset"):
                selected_preset = st.session_state.preset_selected
                preset_data = load_preset(selected_preset)
                if preset_data:
                    copy_num = 1
                    new_name = f"{selected_preset}_copy"
                    while os.path.exists(preset_path(new_name)):
                        copy_num += 1
                        new_name = f"{selected_preset}_copy_{copy_num}"
                    
                    save_preset(new_name, preset_data.get("header_groups", []))
                    st.toast(f"Duplicated '{selected_preset}' to '{new_name}'", icon="✅")
                    st.rerun()
                else:
                    st.toast(f"Failed to load preset for duplication: {selected_preset}", icon="🚨")
        with p_c3:
            is_default = (st.session_state.preset_selected == default_preset_name)
            if is_default:
                st.markdown("⭐ **Default**")
            else:
                if st.button("⭐ Set as Default"):
                    set_default_preset_name(st.session_state.preset_selected)
                    st.toast(f"'{st.session_state.preset_selected}' is now the default preset.", icon="⭐")
                    st.rerun()
        with p_c4:
            if st.button("🗑️ Delete Preset", type="secondary"):
                if delete_preset(st.session_state.preset_selected):
                    if st.session_state.preset_selected == default_preset_name:
                        set_default_preset_name(None)
                    st.toast(f"Deleted preset: {st.session_state.preset_selected}", icon="🗑️")
                    st.rerun()
                else:
                    st.toast(f"Failed to delete preset: {st.session_state.preset_selected}", icon="🚨")

    st.markdown("---")
    st.markdown("#### Save Current Groups as a Preset")
    st.session_state.preset_name_to_save = st.text_input(
        "Preset name (saving will overwrite if name exists)",
        key="preset_name_input"
    )
    if st.button("💾 Save Current Groups as Preset"):
        if not st.session_state.preset_name_to_save.strip():
            st.toast("Please enter a name for the preset.", icon="⚠️")
        elif not st.session_state.header_groups:
            st.toast("There are no groups to save. Please create groups first.", icon="⚠️")
        else:
            save_preset(st.session_state.preset_name_to_save, st.session_state.header_groups)
            st.toast(f"Saved preset: {st.session_state.preset_name_to_save}", icon="💾")
            st.session_state.preset_name_to_save = ""
            st.rerun()

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
        st.toast(f"Scan failed: {e}", icon="🚨"); add_log(f"[ERROR] Scan failed: {e}\n{traceback.format_exc()}")

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

def get_used_columns_map() -> dict[str, set[str]]:
    """
    Scans saved groups to find which specific columns are already mapped for each unit.
    Returns a dict: {unit_key: {used_header_1, used_header_2}}
    """
    used_map = defaultdict(set)
    if not st.session_state.header_groups or not st.session_state.included_units:
        return used_map

    for group in st.session_state.header_groups:
        # The new group structure has a 'mapping' key
        for unit_key, header in group.get("mapping", {}).items():
            if header: # Ensure there's a mapped header
                used_map[unit_key].add(header)
    return used_map

def fuzzy_auto_match(pattern, headers):
    if not pattern or not headers:
        return None
    match, score, _ = process.extractOne(pattern, headers, scorer=fuzz.token_sort_ratio)
    return match if score > 60 else None  # Adjust threshold as needed

with st.expander("Mapping Builder (one output column/group at a time)", expanded=st.session_state.get('is_editing', False)):
    # State for new group
    if "new_group_name" not in st.session_state:
        st.session_state.new_group_name = ""
    if "new_group_pattern" not in st.session_state:
        st.session_state.new_group_pattern = ""
    if "new_group_selections" not in st.session_state or not isinstance(st.session_state.new_group_selections, dict):
        st.session_state.new_group_selections = {}

    # Get a map of all columns that are already used in saved groups
    used_columns_map = get_used_columns_map()
    is_editing_mode = st.session_state.get('is_editing', False)

    # Name and Auto-match fields
    col1, col2, col3 = st.columns([2, 2, 1])
    with col1:
        st.text_input(
            "Output column name (group)",
            key="new_group_name"
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
            used_cols_for_match = get_used_columns_map() # Re-fetch in case state changed
            for uk in st.session_state.included_units:
                all_headers = get_headers_for_unit(uk)
                used_for_this_unit = used_cols_for_match.get(uk, set())
                available_headers = [h for h in all_headers if h not in used_for_this_unit]
                
                match = fuzzy_auto_match(st.session_state.new_group_pattern, available_headers)
                # Only update if there's a match
                if match:
                    st.session_state[f"new_group_select_{uk}"] = match
            st.toast("Auto-match applied.", icon="✨")
            st.rerun()

    # --- Interactive Table with Dropdowns ---
    
    # Sync selections from widgets to the dictionary on every run
    for uk in st.session_state.included_units:
        widget_key = f"new_group_select_{uk}"
        if widget_key in st.session_state:
            st.session_state.new_group_selections[uk] = st.session_state[widget_key]

    # Calculate coverage based on the synced selections
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

    # Add the filter toggle, on by default, disabled when editing
    show_uncovered_only = st.toggle(
        "Show only files without a selection",
        value=False if is_editing_mode else True,
        key="filter_uncovered_toggle",
        disabled=is_editing_mode
    )

    # Add the file name filter
    file_filter_text = st.text_input("Filter files by name", key="file_filter_text")

    # Table Headers and Styling
    st.markdown(
        """
        <style>
        .header-style {
            font-weight: 600;
            padding: 0.5rem 0.25rem;
            border-bottom: 1px solid #444;
            margin-bottom: 0.5rem;
        }
        .row-style {
            display: flex;
            align-items: center;
            min-height: 2.5rem; /* Aligns text with dropdown height */
            padding: 0.1rem 0.25rem;
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
        label = get_unit_label(uk)

        # Apply file name filter
        if file_filter_text and file_filter_text.lower() not in label.lower():
            continue

        current_selection = st.session_state.new_group_selections.get(uk, "")

        # Filter logic: if toggle is on (and not editing), skip rows that have a selection
        if show_uncovered_only and not is_editing_mode and current_selection:
            continue

        row_cols = st.columns([2, 2])
        
        # Filter the headers for this unit to exclude already used columns
        all_headers = get_headers_for_unit(uk)
        used_for_this_unit = used_columns_map.get(uk, set())
        # When editing, the column used by the current group should still be available
        if is_editing_mode and current_selection in used_for_this_unit:
            used_for_this_unit.remove(current_selection)
        available_headers = [h for h in all_headers if h not in used_for_this_unit]
        
        with row_cols[0]:
            st.markdown(f'<div class="row-style">{label}</div>', unsafe_allow_html=True)
        
        with row_cols[1]:
            st.selectbox(
                label=f"selectbox_for_{uk}", # Hidden label for unique widget ID
                label_visibility="collapsed",
                options=[""] + available_headers,
                key=f"new_group_select_{uk}"
            )

    # --- Save Group Logic ---
    def save_group_callback():
        # THIS IS THE FIX:
        # Read from the 'new_group_selections' dictionary which was synced
        # at the top of the script, BEFORE any widgets were filtered out.
        # This dictionary is the reliable source of truth at this point.
        current_selections = st.session_state.get("new_group_selections", {})
        
        coverage_count = sum(1 for sel in current_selections.values() if sel)

        # Validation
        if not st.session_state.new_group_name.strip():
            st.toast("Please enter a group/output column name.", icon="⚠️")
            return
        if coverage_count == 0:
            st.toast("Select at least one column to map for this group.", icon="⚠️")
            return

        # Save the group with the new, precise mapping structure
        st.session_state.header_groups.append({
            "output_name": st.session_state.new_group_name.strip(),
            "mapping": {uk: sel for uk, sel in current_selections.items() if sel}
        })

        # Clear the form fields by resetting their state variables
        st.session_state.new_group_name = ""
        st.session_state.new_group_pattern = ""
        st.session_state.file_filter_text = ""
        st.session_state.new_group_selections = {} # Clear the master dictionary
        # Also clear the individual selectbox widgets so they reset visually
        for uk in st.session_state.included_units:
            if f"new_group_select_{uk}" in st.session_state:
                st.session_state[f"new_group_select_{uk}"] = ""
        
        # Reset editing state
        if 'is_editing' in st.session_state:
            del st.session_state['is_editing']
        
        st.toast("Saved header group.", icon="✅")

    # Display any messages that were set by the callback
    if st.session_state.get("form_error"):
        st.error(st.session_state.form_error)
        del st.session_state.form_error
    
    if st.session_state.get("form_success"):
        st.success(st.session_state.form_success)
        del st.session_state.form_success

    # Save group button that triggers the callback
    st.button("💾 Save group", on_click=save_group_callback)

# ---------- Saved groups (expand for full mapping table) ----------
def edit_group_callback(group_index):
    # Get the group to edit
    group_to_edit = st.session_state.header_groups[group_index]

    # Set editing mode flag
    st.session_state.is_editing = True

    # Populate the mapping builder form
    st.session_state.new_group_name = group_to_edit['output_name']
    st.session_state.new_group_pattern = "" # Clear pattern
    
    # Reconstruct the selections dictionary AND update the individual widget states
    # using the new 'mapping' structure
    selections = {}
    for uk in st.session_state.included_units:
        # Get the saved selection for this unit, or "" if not present
        selection_value = group_to_edit.get("mapping", {}).get(uk, "")
        selections[uk] = selection_value
        st.session_state[f"new_group_select_{uk}"] = selection_value

    st.session_state.new_group_selections = selections

    # Remove the group from the list so saving it again doesn't create a duplicate
    st.session_state.header_groups.pop(group_index)

with st.expander(f"Saved Header Groups ({len(st.session_state.header_groups)})", expanded=False):
    if st.session_state.header_groups and st.session_state.included_units:
        # Iterate backwards so popping doesn't mess up indices
        for i in range(len(st.session_state.header_groups) - 1, -1, -1):
            grp = st.session_state.header_groups[i]
            # Get the unique header variants from the mapping for display
            header_variants = list(set(val for val in grp.get("mapping", {}).values() if val))
            
            with st.expander(f"{i+1}. {grp['output_name']} — {len(header_variants)} header variant(s)", expanded=False):
                grp["output_name"] = st.text_input(f"Rename output column (group {i+1})", value=grp["output_name"], key=f"rename_{i}")
                # Build mapping table: Original Column | Output (Group) | File | Sheet
                rows = []
                for uk in sorted(st.session_state.included_units):
                    meta = st.session_state.unit_meta[uk]
                    # Directly get the match from the group's mapping
                    match = grp.get("mapping", {}).get(uk, "")
                    rows.append({
                        "Original Column": match,
                        "Output (Group)": grp["output_name"],
                        "File": os.path.relpath(meta["file"], st.session_state.input_dir) if st.session_state.input_dir else os.path.basename(meta["file"]),
                        "Sheet": meta["sheet"] or ""
                    })
                df = pd.DataFrame(rows).sort_values(by=["File","Sheet"]).reset_index(drop=True)
                st.dataframe(df, use_container_width=True)

                st.markdown("**Group variants (headers in this group):**")
                st.dataframe(pd.DataFrame({"headers": header_variants}), use_container_width=True)

                b_col1, b_col2 = st.columns(2)
                with b_col1:
                    st.button(f"✏️ Edit group", key=f"edit_{i}", on_click=edit_group_callback, args=(i,))
                with b_col2:
                    if st.button(f"Remove group", key=f"remove_{i}"):
                        st.session_state.header_groups.pop(i)
                        st.toast("Removed group.", icon="🗑️")
                        st.rerun()
    else:
        st.caption("No groups yet, or no included units.")

# ---------- Validate ----------
with st.expander("Validate before merging", expanded=False):
    def validate_groups() -> pd.DataFrame:
        rows = []
        inc = st.session_state.included_units or set()
        for grp in st.session_state.header_groups:
            for uk in inc:
                meta = st.session_state.unit_meta[uk]
                # Directly get the match from the group's mapping
                match = grp.get("mapping", {}).get(uk, "")
                rows.append({
                    "file": os.path.relpath(meta["file"], st.session_state.input_dir) if st.session_state.input_dir else os.path.basename(meta["file"]),
                    "sheet": meta["sheet"] or "",
                    "group": grp["output_name"],
                    "original_header": match,
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
        st.toast("Create at least one header group before merging.", icon="⚠️")
    elif not is_valid_dir(st.session_state.output_dir):
        st.toast("Please set a valid output folder.", icon="⚠️")
    else:
        need_confirm = False
        if not st.session_state.validation_report.empty and (st.session_state.validation_report["covered"]==False).any():
            need_confirm = True
        if need_confirm and not st.session_state.proceed_with_missing:
            st.toast("Run Validate and check the confirmation box to proceed with blanks.", icon="⚠️")
        else:
            inc = sorted(list(st.session_state.included_units or st.session_state.units))
            add_log(f"Starting merge for {len(inc)} unit(s)..."); start = time.time()

            dfs = {}
            for uk in inc:
                fpath, sheet = split_unit_key(uk)
                try:
                    dfs[uk] = read_full_unit(fpath, sheet)
                except Exception as e:
                    add_log(f"[ERROR] Failed reading {unit_label(fpath, sheet, st.session_state.input_dir)}: {e}")

            merged_parts = []
            for uk in inc:
                if uk not in dfs: continue
                df = dfs[uk]
                
                from collections import OrderedDict
                out_cols = OrderedDict()
                
                fpath, sheet = split_unit_key(uk)
                source_label = unit_label(fpath, sheet, st.session_state.input_dir)

                for grp in st.session_state.header_groups:
                    original_header = grp.get("mapping", {}).get(uk)

                    if original_header and original_header in df.columns:
                        data_series = df[original_header]
                    else:
                        data_series = pd.Series([pd.NA] * len(df), index=df.index)

                    if st.session_state.add_source_cols:
                        out_cols[f"{grp['output_name']}_source"] = source_label
                    
                    out_cols[grp['output_name']] = data_series

                part = pd.DataFrame(out_cols)
                merged_parts.append(part)

            merged_df = pd.concat(merged_parts, ignore_index=True) if merged_parts else pd.DataFrame()
            elapsed = time.time() - start
            add_log(f"Merge finished in {elapsed:.2f}s. Rows total: {len(merged_df)}.")

            if not merged_df.empty:
                out_path = os.path.join(st.session_state.output_dir, st.session_state.output_name)
                try:
                    merged_df.to_csv(out_path, index=False)
                    add_log(f"[SAVE] Wrote CSV: {out_path}")
                    st.toast(f"CSV saved: {out_path}", icon="🎉")
                    with open(out_path, "rb") as f:
                        st.download_button("Download CSV", data=f.read(), file_name=os.path.basename(out_path))
                except Exception as e:
                    st.toast(f"Failed to save CSV: {e}", icon="🚨"); add_log(f"[ERROR] Save failed: {e}")
            else:
                st.toast("No data produced. Check groups and units, then try again.", icon="⚠️")

# ---------- Logs ----------
with st.expander("Logs", expanded=False):
    if st.session_state.get("logs"): st.code("\n".join(st.session_state["logs"]), language="text")
    else: st.write("No logs yet.")