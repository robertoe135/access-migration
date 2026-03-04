# streamlit_app.py
import os
import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
import openpyxl

APP_TITLE = "Fixture Schedule Migration to Access"
TEMPLATE_SHEETNAME = "tbeFixtureTypeDetails"
TEMPLATE_CANDIDATE_PATHS = (
    "TEMPLATE.xlsx",
    "./TEMPLATE.xlsx",
    "assets/TEMPLATE.xlsx",
    "./assets/TEMPLATE.xlsx",
)
FIXTURE_CODE_RE = re.compile(r"^[A-Z]{1,3}(?:-\d+)?$")


def _s(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and np.isnan(x):
        return ""
    return str(x)


def is_fixture_code(x) -> bool:
    return bool(FIXTURE_CODE_RE.fullmatch(_s(x).strip()))


def resolve_template_path() -> str:
    # Prefer repo root (Streamlit Cloud uses os.getcwd() as repo root)
    for p in TEMPLATE_CANDIDATE_PATHS:
        abs_p = os.path.join(os.getcwd(), p) if not os.path.isabs(p) else p
        if os.path.exists(abs_p):
            return abs_p

    # Fallback to script directory (in case cwd differs)
    script_dir = os.path.dirname(__file__)
    for p in TEMPLATE_CANDIDATE_PATHS:
        abs_p = os.path.join(script_dir, p)
        if os.path.exists(abs_p):
            return abs_p

    raise FileNotFoundError(
        "TEMPLATE.xlsx not found. Add it to the repo root as 'TEMPLATE.xlsx' "
        "or to 'assets/TEMPLATE.xlsx'."
    )


@st.cache_data(show_spinner=False)
def load_template_defaults() -> Tuple[List[str], Dict[str, object]]:
    template_path = resolve_template_path()
    wb = openpyxl.load_workbook(template_path, data_only=True)
    if TEMPLATE_SHEETNAME not in wb.sheetnames:
        raise ValueError(f"Template must contain a sheet named '{TEMPLATE_SHEETNAME}'.")
    ws = wb[TEMPLATE_SHEETNAME]

    headers = [c.value for c in ws[1]]
    defaults = [c.value for c in ws[2]]

    if not headers or any(h is None for h in headers):
        raise ValueError("Template header row (row 1) appears to be missing/invalid.")

    return headers, dict(zip(headers, defaults))


def block_contains_void(block: pd.DataFrame) -> bool:
    for col in block.columns:
        for v in block[col].values:
            t = _s(v)
            if t and re.search(r"\bVOID\b", t, flags=re.IGNORECASE):
                return True
    return False


def pick_labeled_value(
    block: pd.DataFrame,
    label: str,
    label_col: str = "Unnamed: 3",
    value_col: str = "Unnamed: 4",
) -> str:
    for _, r in block.iterrows():
        lab = _s(r.get(label_col, "")).strip()
        if lab.upper().startswith(label.upper()):
            return _s(r.get(value_col, ""))
    return ""


def pick_lumens(block: pd.DataFrame) -> str:
    col = "Unnamed: 7"
    if col not in block.columns:
        return ""
    for v in block[col].values:
        t = _s(v).strip()
        if t and re.fullmatch(r"\d+(\.\d+)?", t):
            return t
    for v in block[col].values:
        t = _s(v).strip()
        if t:
            return t
    return ""


def pick_input_load(block: pd.DataFrame) -> str:
    for _, r in block.iterrows():
        for c in block.columns:
            t = _s(r.get(c, "")).strip()
            if not t:
                continue
            m = re.search(r"\b(\d+(?:\.\d+)?)\s*[Ww]\b", t)
            if m:
                return m.group(1)
    return ""



def pick_unit(block: pd.DataFrame) -> str:
    unit_cols = ["Unnamed: 8", "Unnamed: 6", "Unnamed: 9"]
    raw = ""
    for c in unit_cols:
        if c in block.columns:
            for v in block[c].values:
                t = _s(v).strip()
                if t:
                    raw = t
                    break
        if raw:
            break
    low = raw.lower()
    if "ln.ft" in low or "ln/ft" in low or "linear" in low:
        return "ln.ft"
    return "each"


def pick_catalog_scored_exact(block: pd.DataFrame) -> str:
    if "Unnamed: 1" not in block.columns:
        return ""
    candidates = []
    for k in range(1, len(block)):
        val = block.iloc[k].get("Unnamed: 1", "")
        if pd.isna(val):
            continue
        txt = str(val)  # preserve exact text (incl notes)
        t = txt.strip()
        if not t:
            continue
        low = t.lower()
        if "http" in low or "www" in low or "@" in t or "#ref" in low:
            continue
        if re.search(r"\(\d{3}\)", t) or re.search(r"\d{3}[-\s]\d{3}[-\s]\d{4}", t):
            continue

        score = 0
        if "_" in t:
            score += 10
        if re.search(r"\d", t):
            score += 3
        if "*" in t:
            score += 2
        if len(t) > 18:
            score += 2
        if re.search(r"[A-Z]{2,}\d", t):
            score += 2
        if re.fullmatch(r"[A-Za-z]+\s*\d+", t) and "_" not in t:
            score -= 6

        candidates.append((score, txt))

    if not candidates:
        return ""
    candidates.sort(key=lambda x: (-x[0], -len(x[1])))
    return candidates[0][1]


def find_best_unit_field(headers: List[str]) -> Optional[str]:
    priority = ["UnitType", "Unit", "Units", "QtyUnit", "QuantityUnit"]
    for p in priority:
        if p in headers:
            return p
    unitish = [h for h in headers if h and re.search(r"unit", str(h), re.IGNORECASE)]
    if not unitish:
        return None
    for h in unitish:
        if re.search(r"type", str(h), re.IGNORECASE):
            return h
    return unitish[0]


def transform(schedule_df: pd.DataFrame, headers: List[str], default_row: Dict[str, object]) -> pd.DataFrame:
    first_col = schedule_df.columns[0]
    fixture_idx = [i for i, v in enumerate(schedule_df[first_col]) if is_fixture_code(v)]
    if not fixture_idx:
        raise ValueError("No fixture designations found.")

    unit_field = find_best_unit_field(headers)
    out_rows = []

    for j, start in enumerate(fixture_idx):
        end = fixture_idx[j + 1] if j + 1 < len(fixture_idx) else len(schedule_df)
        block = schedule_df.iloc[start:end]

        fixture_type = _s(block.iloc[0][first_col]).strip()
        manufacturer = _s(block.iloc[0].get("Unnamed: 1", "")).strip()
        base_desc = _s(block.iloc[0].get("Unnamed: 3", "")).strip()
        dim_proto = _s(block.iloc[0].get("Unnamed: 17", "")).strip()
        if dim_proto.upper() == "0-10V":
            dim_proto = "0-10v (TVI)"

        catalog = pick_catalog_scored_exact(block)
        protection = pick_labeled_value(block, "PROTECTION:").strip()
        location = pick_labeled_value(block, "LOCATION:").strip()
        lumens = pick_lumens(block).strip()
        input_load = pick_input_load(block).strip()
        unit_value = pick_unit(block)
        void_flag = block_contains_void(block)

        out = dict(default_row)

        if "Type" in out:
            out["Type"] = fixture_type
        if "Manufacturer" in out and manufacturer:
            out["Manufacturer"] = manufacturer
        if "CatalogNo" in out and catalog:
            out["CatalogNo"] = catalog
        if "BaseDescription" in out and base_desc:
            out["BaseDescription"] = base_desc
        if "Protection" in out and protection:
            out["Protection"] = protection
        if "Location" in out and location:
            out["Location"] = location
        if "DimProtocol" in out and dim_proto:
            out["DimProtocol"] = dim_proto
        if "LumenOutput" in out and lumens:
            out["LumenOutput"] = lumens

        if input_load:
            if "InputLoad" in out:
                out["InputLoad"] = input_load
            elif "Lamp1InputLoad" in out:
                out["Lamp1InputLoad"] = input_load

        if unit_field and unit_field in out:
            out[unit_field] = unit_value

        if "VOID" in out:
            out["VOID"] = True if void_flag else bool(out["VOID"])

        out_rows.append(out)

    return pd.DataFrame(out_rows, columns=headers)


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)

schedule_file = st.file_uploader("Fixture schedule CSV", type=["csv"], label_visibility="collapsed")
show_detected = st.checkbox("Show preview of detected fixture rows", value=False)
show_output = st.checkbox("Output preview", value=True)

if schedule_file is not None:
    headers, default_row = load_template_defaults()
    schedule_df = pd.read_csv(schedule_file, dtype=str)

    if show_detected:
        first_col = schedule_df.columns[0]
        mask = schedule_df[first_col].apply(is_fixture_code)
        st.dataframe(schedule_df.loc[mask, [first_col]].head(200), use_container_width=True)

    out_df = transform(schedule_df, headers, default_row)

    if show_output:
        st.dataframe(out_df.head(200), use_container_width=True)

    st.download_button(
        "Download CSV",
        data=out_df.to_csv(index=False).encode("utf-8"),
        file_name="fixture_template_export.csv",
        mime="text/csv",
    )
