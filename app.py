import streamlit as st
import pandas as pd
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="Eastern Region Clock Report", layout="centered")

# --- SIMPLE DARK MODE CSS ---
simple_dark_style = """
    <style>
    /* 1. Main Background - Standard Dark Mode */
    .stApp {
        background-color: #0E1117;
        color: #FAFAFA;
    }

    /* 2. Hide Streamlit Branding (Privacy) */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* 3. Simple Button Style (Standard Streamlit Dark Look) */
    .stButton > button {
        width: 100%;
    }
    </style>
    """
st.markdown(simple_dark_style, unsafe_allow_html=True)

# --- APP LOGIC STARTS HERE ---

st.header("Clock Report Processor")

def create_pivot_view(df_input, group_cols):
    """
    Simulates a Pivot Table "Tabular View" by sorting and hiding repeated labels.
    """
    df_sorted = df_input.fillna("").sort_values(by=group_cols).copy()
    df_display = df_sorted.astype(str).copy()
    
    mask_cols = group_cols[:-1] 
    prev_row = {col: None for col in mask_cols}
    formatted_rows = []
    
    for _, row in df_sorted.iterrows():
        current_row = []
        is_parent_same = True 
        
        for col in group_cols:
            val = row[col]
            if col in mask_cols:
                if is_parent_same and val == prev_row[col]:
                    current_row.append("") 
                else:
                    current_row.append(val) 
                    is_parent_same = False 
                prev_row[col] = val
            else:
                current_row.append(val)
        formatted_rows.append(current_row)
        
    return df_sorted, pd.DataFrame(formatted_rows, columns=group_cols)

# 1. File Upload
# Reverted to standard label visibility for clarity in simple mode
uploaded_file = st.file_uploader("Upload 'Clock Detail Report' (xlsx)", type="xlsx")

if uploaded_file:
    try:
        # Load ALL sheets
        all_sheets = pd.read_excel(uploaded_file, sheet_name=None)
        
        source_sheet_name = "Clock Detail Report"
        if source_sheet_name not in all_sheets:
            st.error(f"Error: The file must contain a sheet named '{source_sheet_name}'.")
            st.stop()
            
        df_source = all_sheets[source_sheet_name]
        df_source.columns = df_source.columns.astype(str).str.strip()
        
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # --- DEFINE FORMATS ---
            base_props = {'border': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True}
            thick_top_props = {'top': 2, 'bottom': 1, 'left': 1, 'right': 1, 'align': 'left', 'valign': 'top', 'text_wrap': True}
            
            fmt_std = workbook.add_format(base_props)
            fmt_std_thick = workbook.add_format(thick_top_props)
            fmt_bold = workbook.add_format({**base_props, 'bold': True})
            fmt_bold_thick = workbook.add_format({**thick_top_props, 'bold': True})
            fmt_orange = workbook.add_format({**base_props, 'bg_color': '#FFC000', 'font_color': '#000000'})
            fmt_orange_thick = workbook.add_format({**thick_top_props, 'bg_color': '#FFC000', 'font_color': '#000000'})
            
            header_fmt = workbook.add_format({
                'bold': True, 'text_wrap': False, 'valign': 'vcenter', 'align': 'center',
                'fg_color': '#D9E1F2', 'border': 1
            })

            # Write Original Sheets
            for sheet_name, df in all_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            categories = ["ECNB", "ECMW"]
            
            for category in categories:
                if len(df_source.columns) < 9:
                    st.error("Error: File has fewer than 9 columns.")
                    st.stop()

                # --- FILTERING LOGIC ---
                mask_category = df_source.iloc[:, 8].astype(str).str.contains(category, case=False, na=False)
                
                if category == "ECNB":
                    dist_numeric = pd.to_numeric(df_source.iloc[:, 5], errors='coerce')
                    mask_distance = dist_numeric <= 500
                    df_filtered = df_source[mask_category & mask_distance]
                else:
                    df_filtered = df_source[mask_category]
                
                df_filtered.to_excel(writer, sheet_name=f"Data {category}", index=False)
                
                # --- PIVOT PREPARATION ---
                pivot_cols_base = ["Company", "Name", "Account", "DU ID"]
                time_col_name = df_source.columns[4] 
                
                missing = [c for c in pivot_cols_base + [time_col_name] if c not in df_filtered.columns]
                if missing:
                    st.error(f"Missing columns: {missing}")
                    st.stop()
                
                df_working = df_filtered.copy()
                df_working[time_col_name] = pd.to_datetime(df_working[time_col_name], errors='coerce')
                df_aggregated = df_working.groupby(pivot_cols_base)[time_col_name].min().reset_index()
                df_aggregated[time_col_name] = df_aggregated[time_col_name].dt.strftime('%H:%M:%S').fillna("-")
                
                pivot_cols_final = pivot_cols_base + [time_col_name]
                
                df_sorted, df_display = create_pivot_view(df_aggregated, pivot_cols_final)
                
                pivot_sheet_name = f"Pivot {category}"
                worksheet = workbook.add_worksheet(pivot_sheet_name)
                writer.sheets[pivot_sheet_name] = worksheet
                
                worksheet.freeze_panes(3, 0)
                worksheet.autofilter(2, 0, 2 + len(df_display), len(pivot_cols_final) - 1)
                
                for col_num, val in enumerate(pivot_cols_final):
                    worksheet.write(2, col_num, val, header_fmt)
                
                for row_idx, row_data in df_display.iterrows():
                    is_new_subcon = (row_data[0] != "") and (row_idx > 0)
                    actual_du_id = df_sorted.iloc[row_idx]["DU ID"]
                    is_duplicate_du = len(df_sorted[df_sorted["DU ID"] == actual_du_id]) > 1
                    
                    excel_row = row_idx + 3
                    
                    for col_idx, cell_value in enumerate(row_data):
                        cell_fmt = fmt_std 
                        if is_new_subcon:
                            if col_idx == 0 and cell_value != "": cell_fmt = fmt_bold_thick
                            elif col_idx == 3 and is_duplicate_du: cell_fmt = fmt_orange_thick
                            else: cell_fmt = fmt_std_thick
                        else:
                            if col_idx == 0 and cell_value != "": cell_fmt = fmt_bold
                            elif col_idx == 3 and is_duplicate_du: cell_fmt = fmt_orange
                            else: cell_fmt = fmt_std

                        worksheet.write(excel_row, col_idx, cell_value, cell_fmt)

                worksheet.set_column(0, 0, 40) 
                worksheet.set_column(1, 1, 30) 
                worksheet.set_column(2, 2, 20) 
                worksheet.set_column(3, 3, 25) 
                worksheet.set_column(4, 4, 15) 
                
                summary = df_filtered.groupby("Company")["Name"].nunique().reset_index()
                summary.columns = ["Company", "Count of Name"]
                
                worksheet.write("G3", "Company", header_fmt)
                worksheet.write("H3", "Count of Name", header_fmt)
                
                last_row = 3
                for idx, row in summary.iterrows():
                    last_row = idx + 3
                    worksheet.write(last_row, 6, row["Company"], fmt_std)
                    worksheet.write(last_row, 7, row["Count of Name"], fmt_std)
                
                total_row = last_row + 1
                total_count = summary["Count of Name"].sum()
                worksheet.write(total_row, 6, "Grand Total", header_fmt)
                worksheet.write(total_row, 7, total_count, header_fmt)
                
                worksheet.set_column(6, 6, 40)
                worksheet.set_column(7, 7, 15)

        output.seek(0)
        
        # Add a little spacing
        st.write("")
        st.write("")
        
        # 2. Download Button
        st.download_button(
            label="Download Processed Report",
            data=output,
            file_name="Processed_ClockReport.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Simple success text below
        st.success("Processing Complete")
        
    except Exception as e:
        st.error(f"An error occurred: {e}")
