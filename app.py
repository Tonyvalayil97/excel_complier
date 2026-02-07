import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Excel/CSV Compiler", layout="wide")

st.title("📎 Compile Multiple Excel/CSV Files into One Excel")
st.caption("Upload files one-by-one. All files must have the same columns. Download a single compiled .xlsx at the end.")

# --- Session state init ---
if "expected_cols" not in st.session_state:
    st.session_state.expected_cols = None
if "compiled_df" not in st.session_state:
    st.session_state.compiled_df = None
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []


def read_any_table(uploaded_file) -> pd.DataFrame:
    """Read .xlsx/.xls/.csv safely into a DataFrame."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    elif name.endswith(".xlsx") or name.endswith(".xls"):
        # reads first sheet by default
        return pd.read_excel(uploaded_file, engine="openpyxl")
    else:
        raise ValueError("Unsupported file type. Please upload .csv, .xlsx, or .xls")


def normalize_columns(cols):
    """Keep columns as-is but ensure they are strings and strip whitespace."""
    return [str(c).strip() for c in cols]


def cols_match(expected, incoming):
    """Exact match in same order."""
    return expected == incoming


def diff_cols(expected, incoming):
    exp_set = set(expected)
    inc_set = set(incoming)
    missing = [c for c in expected if c not in inc_set]
    extra = [c for c in incoming if c not in exp_set]
    order_issue = exp_set == inc_set and expected != incoming
    return missing, extra, order_issue


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Convert dataframe to Excel bytes."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="compiled")
    return output.getvalue()


with st.sidebar:
    st.header("Controls")
    add_source_col = st.checkbox("Add source filename column", value=True)
    source_col_name = st.text_input("Source column name", value="source_file", disabled=not add_source_col)

    st.divider()
    if st.button("🧹 Reset / Clear all", type="secondary"):
        st.session_state.expected_cols = None
        st.session_state.compiled_df = None
        st.session_state.uploaded_files = []
        st.rerun()

st.subheader("1) Upload one file at a time")
uploaded = st.file_uploader(
    "Upload a CSV or Excel file",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=False
)

if uploaded is not None:
    try:
        df_new = read_any_table(uploaded)
        df_new.columns = normalize_columns(df_new.columns)

        if st.session_state.expected_cols is None:
            st.session_state.expected_cols = df_new.columns.tolist()

            if add_source_col:
                df_new[source_col_name] = uploaded.name

            st.session_state.compiled_df = df_new.copy()
            st.session_state.uploaded_files.append(uploaded.name)
            st.success(f"✅ First file loaded. Columns locked ({len(st.session_state.expected_cols)} columns). Appended {len(df_new)} rows.")

        else:
            incoming_cols = df_new.columns.tolist()
            expected = st.session_state.expected_cols

            if cols_match(expected, incoming_cols):
                if add_source_col:
                    df_new[source_col_name] = uploaded.name

                st.session_state.compiled_df = pd.concat(
                    [st.session_state.compiled_df, df_new],
                    ignore_index=True
                )
                st.session_state.uploaded_files.append(uploaded.name)
                st.success(f"✅ Appended {len(df_new)} rows from {uploaded.name}.")
            else:
                missing, extra, order_issue = diff_cols(expected, incoming_cols)

                st.error("❌ Column mismatch. This file was NOT added.")
                c1, c2, c3 = st.columns(3)
                with c1:
                    st.write("**Missing columns (expected but not found):**")
                    st.write(missing if missing else "None")
                with c2:
                    st.write("**Extra columns (found but not expected):**")
                    st.write(extra if extra else "None")
                with c3:
                    st.write("**Order issue:**")
                    st.write("Yes (same columns, different order)" if order_issue else "No")

                with st.expander("Show expected columns"):
                    st.write(expected)
                with st.expander("Show incoming columns"):
                    st.write(incoming_cols)

    except Exception as e:
        st.error(f"Failed to read file: {e}")

st.divider()

st.subheader("2) Current compiled output")

if st.session_state.compiled_df is None:
    st.info("No data yet. Upload your first file to start compiling.")
else:
    df_all = st.session_state.compiled_df

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Files added", len(st.session_state.uploaded_files))
    with col2:
        st.metric("Total rows", len(df_all))
    with col3:
        st.metric("Total columns", len(df_all.columns))

    st.write("**Uploaded files:**", ", ".join(st.session_state.uploaded_files))

    st.write("Preview (first 200 rows):")
    st.dataframe(df_all.head(200), use_container_width=True)

    st.subheader("3) Download compiled Excel")
    default_name = f"compiled_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    filename = st.text_input("Output filename", value=default_name)

    excel_bytes = to_excel_bytes(df_all)
    st.download_button(
        label="⬇️ Download compiled .xlsx",
        data=excel_bytes,
        file_name=filename if filename.lower().endswith(".xlsx") else f"{filename}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
