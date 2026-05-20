import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Daily Team Plan Generator", layout="wide")

REQUIRED_COLUMNS = {
    "Day No": ["Day No", "Day", "DayNo", "DAY NO"],
    "Team No": ["Team No", "Team", "TeamNo", "TEAM NO"],
    "Visit Sequence": ["Visit Sequence", "Sequence", "Visit Seq", "Seq", "VISIT SEQUENCE"],
    "Team Day Town": ["Team Day Town", "Town", "Tehsil", "Area", "TEAM DAY TOWN"],
    "Business Name": ["Business Name", "Business", "Shop Name", "Outlet Name", "BUSINESS NAME"],
    "Address": ["Address", "Location", "ADDRESS"],
    "FBO Name": ["FBO Name", "FBO", "Owner Name", "Contact Person", "FBO NAME"],
    "Contact": ["Contact", "Phone", "Mobile", "Contact No", "CONTACT"],
    "Planned Category": ["Planned Category", "Category", "Inspection Category", "PLANNED CATEGORY"],
}

OUTPUT_ORDER = [
    "Visit Sequence",
    "Team Day Town",
    "Business Name",
    "Address",
    "FBO Name",
    "Contact",
    "Planned Category",
]

DISPLAY_NAMES = {
    "Visit Sequence": "Visit Sequence",
    "Team Day Town": "Town",
    "Business Name": "Business Name",
    "Address": "Address",
    "FBO Name": "FBO Name",
    "Contact": "Contact",
    "Planned Category": "Planned Category",
}


def normalize_text(value):
    return str(value).strip().lower().replace("_", " ").replace("-", " ")


def find_column(df_columns, possible_names):
    normalized = {normalize_text(col): col for col in df_columns}
    for name in possible_names:
        key = normalize_text(name)
        if key in normalized:
            return normalized[key]
    return None


@st.cache_data(show_spinner=False)
def read_excel_file(uploaded_file_bytes):
    workbook = pd.ExcelFile(io.BytesIO(uploaded_file_bytes))
    return workbook.sheet_names, uploaded_file_bytes


def load_sheet(uploaded_file_bytes, sheet_name):
    df = pd.read_excel(io.BytesIO(uploaded_file_bytes), sheet_name=sheet_name)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how="all")
    return df


def auto_map_columns(df):
    mapping = {}
    missing = []
    for standard_name, possible_names in REQUIRED_COLUMNS.items():
        actual = find_column(df.columns, possible_names)
        if actual is None:
            missing.append(standard_name)
        else:
            mapping[standard_name] = actual
    return mapping, missing


def make_daily_plan(df, mapping, day_no, team_no):
    day_col = mapping["Day No"]
    team_col = mapping["Team No"]
    seq_col = mapping["Visit Sequence"]

    working = df.copy()
    working["__day__"] = pd.to_numeric(working[day_col], errors="coerce")
    working["__team__"] = pd.to_numeric(working[team_col], errors="coerce")
    working["__seq__"] = pd.to_numeric(working[seq_col], errors="coerce")

    filtered = working[(working["__day__"] == float(day_no)) & (working["__team__"] == float(team_no))]
    filtered = filtered.sort_values("__seq__", na_position="last")

    output = pd.DataFrame()
    for standard_col in OUTPUT_ORDER:
        actual_col = mapping.get(standard_col)
        output[DISPLAY_NAMES[standard_col]] = filtered[actual_col] if actual_col in filtered.columns else ""
    output["Remarks"] = ""
    return output.reset_index(drop=True)


def plan_to_excel_bytes(plan_df, day_no, team_no):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        sheet_name = f"Day {day_no} Team {team_no}"[:31]
        plan_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=3)
        workbook = writer.book
        worksheet = writer.sheets[sheet_name]

        title_format = workbook.add_format({"bold": True, "font_size": 16, "align": "center"})
        header_format = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1, "text_wrap": True})
        cell_format = workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"})

        worksheet.merge_range(0, 0, 0, len(plan_df.columns) - 1, f"Daily Inspection Plan - Day {day_no} - Team {team_no}", title_format)
        worksheet.write(1, 0, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

        for col_num, col_name in enumerate(plan_df.columns):
            worksheet.write(3, col_num, col_name, header_format)
            if col_name in ["Address", "Business Name"]:
                worksheet.set_column(col_num, col_num, 35, cell_format)
            elif col_name == "Remarks":
                worksheet.set_column(col_num, col_num, 22, cell_format)
            else:
                worksheet.set_column(col_num, col_num, 18, cell_format)

        for row_num in range(4, 4 + len(plan_df)):
            worksheet.set_row(row_num, 45)
            for col_num in range(len(plan_df.columns)):
                value = plan_df.iloc[row_num - 4, col_num]
                if pd.isna(value):
                    value = ""
                worksheet.write(row_num, col_num, value, cell_format)

        worksheet.freeze_panes(4, 0)
        worksheet.set_landscape()
        worksheet.fit_to_pages(1, 0)
        worksheet.set_margins(left=0.25, right=0.25, top=0.5, bottom=0.5)
    output.seek(0)
    return output.getvalue()


def plans_zip_bytes(df, mapping, day_no, teams):
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for team in teams:
            plan_df = make_daily_plan(df, mapping, day_no, team)
            excel_bytes = plan_to_excel_bytes(plan_df, day_no, team)
            filename = f"Day_{day_no}_Team_{team}_Plan.xlsx"
            zip_file.writestr(filename, excel_bytes)
    output.seek(0)
    return output.getvalue()


st.title("Daily Team Plan Generator")
st.caption("Upload the master Excel file, choose the day and team, then download a clean daily plan.")

with st.expander("How to use this app", expanded=False):
    st.markdown(
        """
        1. Upload the master Excel file.  
        2. Select the sheet containing the inspection plan.  
        3. Confirm that the app detected the required columns.  
        4. Select Day No and Team No.  
        5. Click the download button to export the team plan.
        """
    )

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if not uploaded_file:
    st.info("Upload your Excel file to begin.")
    st.stop()

uploaded_file_bytes = uploaded_file.getvalue()
sheet_names, stored_bytes = read_excel_file(uploaded_file_bytes)

preferred_sheet_index = 0
for i, name in enumerate(sheet_names):
    if name.strip().lower() == "revised inspection plan":
        preferred_sheet_index = i
        break

sheet_name = st.selectbox("Select sheet", sheet_names, index=preferred_sheet_index)
df = load_sheet(stored_bytes, sheet_name)

st.subheader("Data Preview")
st.write(f"Rows loaded: **{len(df):,}** | Columns loaded: **{len(df.columns):,}**")
st.dataframe(df.head(10), use_container_width=True)

mapping, missing = auto_map_columns(df)

if missing:
    st.warning("Some required columns were not detected automatically. Please map them manually below.")
    for col in missing:
        mapping[col] = st.selectbox(f"Select column for: {col}", [""] + list(df.columns), key=f"map_{col}")
else:
    st.success("All required columns detected automatically.")

with st.expander("Column mapping", expanded=False):
    mapping_table = pd.DataFrame([{"Required Field": k, "Detected Excel Column": v} for k, v in mapping.items()])
    st.dataframe(mapping_table, use_container_width=True)

if any(not mapping.get(col) for col in REQUIRED_COLUMNS):
    st.error("Please map all required columns before generating plans.")
    st.stop()

# Prepare dropdown values
day_values = sorted(pd.to_numeric(df[mapping["Day No"]], errors="coerce").dropna().astype(int).unique().tolist())
team_values = sorted(pd.to_numeric(df[mapping["Team No"]], errors="coerce").dropna().astype(int).unique().tolist())

if not day_values or not team_values:
    st.error("Could not find valid Day No or Team No values in the selected sheet.")
    st.stop()

left, right = st.columns(2)
with left:
    day_no = st.selectbox("Select Day No", day_values)
with right:
    team_no = st.selectbox("Select Team No", team_values)

plan_df = make_daily_plan(df, mapping, day_no, team_no)

st.subheader(f"Daily Inspection Plan - Day {day_no} - Team {team_no}")
st.write(f"Total planned visits: **{len(plan_df):,}**")
st.dataframe(plan_df, use_container_width=True, hide_index=True)

selected_excel = plan_to_excel_bytes(plan_df, day_no, team_no)
st.download_button(
    label="Download selected team plan as Excel",
    data=selected_excel,
    file_name=f"Day_{day_no}_Team_{team_no}_Plan.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

zip_bytes = plans_zip_bytes(df, mapping, day_no, team_values)
st.download_button(
    label="Download all team plans for selected day as ZIP",
    data=zip_bytes,
    file_name=f"Day_{day_no}_All_Team_Plans.zip",
    mime="application/zip",
)

st.markdown("---")
st.subheader("Print View")
st.caption("Use your browser print command if you want a paper copy of the table shown below.")
st.table(plan_df)
