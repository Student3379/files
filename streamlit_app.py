import io
import warnings
from typing import Optional

import pandas as pd
import streamlit as st

# --- Streamlit / App Setup ----------------------------------------------------
warnings.filterwarnings("ignore")
st.set_page_config(page_title="DEAL File", page_icon="📂", layout="wide")

# 🔒 Hide all Streamlit chrome (floating toolbar, top-right icons, menu, header, footer)
_hide_all_streamlit_ui = """
    <style>
    /* Floating toolbar / deploy / status */
    [data-testid="stToolbar"] {visibility: hidden !important;}
    [data-testid="stDecoration"] {visibility: hidden !important;}
    [data-testid="stStatusWidget"] {visibility: hidden !important;}
    .viewerBadge_container__1QSob {display: none !important;}
    .stAppDeployButton {display: none !important;}

    /* Top-right action icons (Fork / GitHub / …) */
    [data-testid="stHeaderActionElements"] {display: none !important;}

    /* Streamlit menu, header, footer */
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    </style>
"""
st.markdown(_hide_all_streamlit_ui, unsafe_allow_html=True)

PREVIEW_ROWS = 100

# --- Utilities ----------------------------------------------------------------
def _arrow_safe_df(df: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
    if df is None or not isinstance(df, pd.DataFrame):
        return df
    out = df.copy()
    cols = pd.Index([str(c) for c in out.columns])
    mask = ~cols.str.match(r"^Unnamed(:\s*\d+)?$")
    out = out.loc[:, mask]
    out.columns = [str(c) for c in out.columns]
    for c in out.columns:
        if pd.api.types.is_object_dtype(out[c]):
            try:
                if out[c].map(type).nunique(dropna=False) > 1:
                    out[c] = out[c].astype(str)
            except Exception:
                out[c] = out[c].astype(str)
    return out

def _excel_engine_for_name(lower_name: str) -> Optional[str]:
    if lower_name.endswith(".xls"):  return "xlrd"
    if lower_name.endswith(".xlsx"): return "openpyxl"
    return None

def _read_excel_generic(content_bytes: bytes, *, file_name: str, skiprows: int = 0, sheet_name=0, nrows: Optional[int] = None):
    lower = file_name.lower() if file_name else ""
    engine = _excel_engine_for_name(lower)
    buf = io.BytesIO(content_bytes)
    try:
        return pd.read_excel(buf, skiprows=skiprows, nrows=nrows, sheet_name=sheet_name, engine=engine)
    except Exception:
        try:
            buf.seek(0)
            return pd.read_excel(buf, skiprows=skiprows, nrows=nrows, sheet_name=sheet_name)
        except Exception as e2:
            if lower.endswith(".xls"):
                st.error("Failed to read .xls. Install xlrd: `pip install xlrd`")
            elif lower.endswith(".xlsx"):
                st.error("Failed to read .xlsx. Install openpyxl: `pip install openpyxl`")
            raise e2

# --- Cached Readers ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _read_csv_preview(content_bytes: bytes, skiprows: int = 0, nrows: int = PREVIEW_ROWS):
    buf = io.BytesIO(content_bytes)
    try:
        return pd.read_csv(buf, skiprows=skiprows, nrows=nrows)
    except Exception:
        buf.seek(0)
        return pd.read_csv(buf, skiprows=skiprows).head(nrows)

@st.cache_data(show_spinner=False)
def _read_excel_preview(content_bytes: bytes, file_name: str, skiprows: int = 0, nrows: int = PREVIEW_ROWS, sheet_name=0):
    try:
        return _read_excel_generic(content_bytes, file_name=file_name, skiprows=skiprows, nrows=nrows, sheet_name=sheet_name)
    except Exception:
        buf = io.BytesIO(content_bytes)
        return pd.read_excel(buf, skiprows=skiprows, sheet_name=sheet_name).head(nrows)

@st.cache_data(show_spinner=False)
def _read_csv_full(content_bytes: bytes, skiprows: int = 0):
    buf = io.BytesIO(content_bytes)
    return pd.read_csv(buf, skiprows=skiprows)

@st.cache_data(show_spinner=False)
def _read_excel_full(content_bytes: bytes, file_name: str, skiprows: int = 0, sheet_name=0):
    return _read_excel_generic(content_bytes, file_name=file_name, skiprows=skiprows, nrows=None, sheet_name=sheet_name)

# --- File helpers --------------------------------------------------------------
def _file_to_bytes(uploaded):  return uploaded.getvalue() if uploaded else None
def _safe_cols(df: pd.DataFrame) -> list[str]: return [str(c) for c in df.columns]

def _read_preview(uploaded, skiprows: int = 0):
    if not uploaded: return None
    data = _file_to_bytes(uploaded); name = uploaded.name
    if name.lower().endswith(".csv"): return _read_csv_preview(data, skiprows=skiprows, nrows=PREVIEW_ROWS)
    return _read_excel_preview(data, file_name=name, skiprows=skiprows, nrows=PREVIEW_ROWS)

def _read_full(uploaded, skiprows: int = 0):
    if not uploaded: return None
    data = _file_to_bytes(uploaded); name = uploaded.name
    if name.lower().endswith(".csv"): return _read_csv_full(data, skiprows=skiprows)
    return _read_excel_full(data, file_name=name, skiprows=skiprows)

# --- Key normalization / alignment --------------------------------------------
def _clean_text_like(s: pd.Series, *, lower: bool = True, strip_all_ws: bool = True) -> pd.Series:
    out = s.astype(str)
    out = out.str.replace(r"\s+", " ", regex=True).str.strip()
    out = out.str.replace(r"\.0$", "", regex=True)
    if strip_all_ws:
        out = out.str.replace(" ", "", regex=False)
    if lower:
        out = out.str.lower()
    return out

def _smart_align_keys(left: pd.Series, right: pd.Series):
    l_num = pd.to_numeric(left, errors="coerce")
    r_num = pd.to_numeric(right, errors="coerce")
    if not l_num.isna().all() and not r_num.isna().all():
        l_int_ok = ((l_num.dropna() % 1) == 0).all()
        r_int_ok = ((r_num.dropna() % 1) == 0).all()
        if l_int_ok and r_int_ok:
            return l_num.astype("Int64"), r_num.astype("Int64"), "auto:number(Int64)"
        return l_num, r_num, "auto:number(float)"
    lt = _clean_text_like(left, lower=True, strip_all_ws=True)
    rt = _clean_text_like(right, lower=True, strip_all_ws=True)
    return lt, rt, "auto:text(case+whitespace-insensitive)"

# --- Sidebar / Controls --------------------------------------------------------
st.sidebar.header("Upload Files")

file1 = st.sidebar.file_uploader("First File", type=["csv", "xlsx", "xls"], key="file1")
if file1: st.sidebar.write(f"📄 File 1 uploaded: **{file1.name}**")
skip1 = st.sidebar.number_input("Skip rows (File 1)", 0, 100000, 0, 1)

file2 = st.sidebar.file_uploader("Second File", type=["csv", "xlsx", "xls"], key="file2")
if file2: st.sidebar.write(f"📄 File 2 uploaded: **{file2.name}**")
skip2 = st.sidebar.number_input("Skip rows (File 2)", 0, 100000, 0, 1)

if "show_vlookup" not in st.session_state: st.session_state.show_vlookup = False
if "show_merge"   not in st.session_state: st.session_state.show_merge   = False

with st.container():
    st.markdown('<div class="topbar-sticky"><div class="topbar-card">', unsafe_allow_html=True)
    top_l, top_r = st.columns([0.75, 0.25])
    with top_r:
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("🔍 VLOOKUP", use_container_width=True):
                st.session_state.show_vlookup = not st.session_state.show_vlookup
        with col_b:
            if st.button("🧾 Merge Files", use_container_width=True):
                st.session_state.show_merge = not st.session_state.show_merge
    st.markdown('</div></div>', unsafe_allow_html=True)

# --- VLOOKUP -------------------------------------------------------------------
if st.session_state.show_vlookup:
    st.markdown("---")
    st.subheader("🔍 VLOOKUP (File 1 → File 2)")

    if not (file1 and file2):
        st.info("Upload both files first to use VLOOKUP.")
    else:
        try:
            df1_full = _read_full(file1, skip1)
            df2_full = _read_full(file2, skip2)
        except Exception as e:
            st.error(f"Error loading full data for VLOOKUP: {e}")
            df1_full, df2_full = None, None

        if isinstance(df1_full, pd.DataFrame) and isinstance(df2_full, pd.DataFrame):
            cols1, cols2 = _safe_cols(df1_full), _safe_cols(df2_full)

            with st.form("vlookup_form", clear_on_submit=False):
                a, b, c = st.columns(3)
                with a:
                    left_key  = st.selectbox("Key column in File 1", options=cols1, key="vk_left")
                with b:
                    right_key = st.selectbox("Key column in File 2", options=cols2, key="vk_right")
                with c:
                    fetch_cols = st.multiselect(
                        "Columns to bring from File 2",
                        options=[col for col in cols2 if col != right_key],
                        default=[]
                    )
                run = st.form_submit_button("Apply VLOOKUP")

            if run:
                try:
                    right = df2_full.drop_duplicates(subset=[right_key], keep="first")

                    # Auto-align keys (numeric if both numeric; else case+whitespace-insensitive text)
                    lnorm, rnorm, _ = _smart_align_keys(df1_full[left_key], right[right_key])

                    df1_join = df1_full.copy()
                    right = right.copy()
                    df1_join["_join_key_"] = lnorm
                    right["_join_key_"] = rnorm

                    # Avoid name clashes for fetched columns
                    clashes = [c for c in fetch_cols if c in df1_join.columns]
                    renames = {c: f"{c}_from_file2" for c in clashes}
                    right = right.rename(columns=renames)
                    use_cols = ["_join_key_"] + [renames.get(c, c) for c in fetch_cols]

                    merged = df1_join.merge(
                        right[use_cols],
                        how="left",
                        left_on="_join_key_",
                        right_on="_join_key_",
                        suffixes=("", " ")
                    ).drop(columns=["_join_key_"], errors="ignore")

                    merged_preview = _arrow_safe_df(merged.head(PREVIEW_ROWS))
                    st.dataframe(merged_preview, width="stretch", height=440)

                    f1 = (file1.name.rsplit('.', 1)[0] if file1 else "file1").strip().replace(" ", "_")
                    f2 = (file2.name.rsplit('.', 1)[0] if file2 else "file2").strip().replace(" ", "_")
                    output_name = f"{f1}_{f2}.xlsx"

                    xbuf = io.BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
                        merged.to_excel(writer, index=False, sheet_name="VLOOKUP")
                    xbuf.seek(0)
                    st.download_button(
                        "⬇️ Download",
                        data=xbuf,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"VLOOKUP failed: {e}")

# --- Merge Files ---------------------------------------------------------------
if st.session_state.show_merge:
    st.markdown("---")
    st.subheader("🧾 Merge Files into Excel")

    with st.form("merge_form_simple", clear_on_submit=False):
        files_to_merge = st.file_uploader(
            "Drag & Drop multiple in (CSV/XLSX/XLS)",
            type=["csv", "xlsx", "xls"],
            accept_multiple_files=True,
            key="merge_files_uploader_simple",
        )
        run_merge = st.form_submit_button("Excel")

    if run_merge:
        try:
            selected = files_to_merge or []
            if not selected:
                st.warning("Please select at least one file.")
            else:
                frames, all_cols = [], set()
                for up in selected:
                    name = up.name
                    data = _file_to_bytes(up)
                    if name.lower().endswith(".csv"):
                        df = pd.read_csv(io.BytesIO(data))
                    else:
                        df = _read_excel_generic(data, file_name=name)
                    df = _arrow_safe_df(df)
                    all_cols |= set(map(str, df.columns))
                    frames.append((name, df))

                all_cols = list(all_cols)
                combined = []
                for _, df in frames:
                    tmp = df.copy()
                    for c in all_cols:
                        if c not in tmp.columns:
                            tmp[c] = pd.NA
                    tmp = tmp[all_cols]
                    combined.append(tmp)

                combined_df = pd.concat(combined, ignore_index=True)

                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="openpyxl") as writer:
                    combined_df.to_excel(writer, index=False, sheet_name="Combined")
                out.seek(0)
                out_name = "Merged.xlsx"
                st.success("Excel File Generated")
                st.caption("Preview of Combined File")
                st.dataframe(_arrow_safe_df(combined_df.head(PREVIEW_ROWS)), width="stretch", height=420)
                st.download_button(
                                    "⬇️ Download Combined Excel",
                                    data=out,
                                    file_name=out_name,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                )
        except Exception as e:
            st.error(f"Merge failed: {e}")

# --- Bottom area: Previews -----------------------------------------------------
st.title("📂")

df1_prev = _read_preview(file1, skip1) if file1 else None
df2_prev = _read_preview(file2, skip2) if file2 else None

df1_prev_safe = _arrow_safe_df(df1_prev) if isinstance(df1_prev, pd.DataFrame) else None
df2_prev_safe = _arrow_safe_df(df2_prev) if isinstance(df2_prev, pd.DataFrame) else None

c1, c2 = st.columns(2)
if isinstance(df1_prev_safe, pd.DataFrame):
    with c1:
        st.markdown(f"### 📄 File 1: `{file1.name}` (skip {skip1})")
        st.dataframe(df1_prev_safe, width="stretch", height=420)

if isinstance(df2_prev_safe, pd.DataFrame):
    with c2:
        st.markdown(f"### 📄 File 2: `{file2.name}` (skip {skip2})")
        st.dataframe(df2_prev_safe, width="stretch", height=420)
