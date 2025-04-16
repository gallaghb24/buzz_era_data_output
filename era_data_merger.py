import streamlit as st
import pandas as pd
import io
from datetime import datetime
from parts_data import PARTS_DATA_RAW

st.set_page_config(page_title="ERA Data Merger", layout="centered")
st.title("🔄 ERA Club & Production Data Merger")

# --- PARTS LOOKUP TRANSFORM ---
PARTS_DATA = {}
for part_data in PARTS_DATA_RAW.values():
    PARTS_DATA[part_data["Part Name"]] = {
        "Size": f"{part_data['Height (mm)']}x{part_data['Width (mm)']}",
        "Pagination": part_data["No. of Pages"],
        "Material": part_data["Materials"],
        "Finishing": part_data.get("Finishing", "")
    }

# --- SESSION STATE SETUP ---
if "step" not in st.session_state:
    st.session_state.step = 1

if "club_files" not in st.session_state:
    st.session_state.club_files = []

if "prod_files" not in st.session_state:
    st.session_state.prod_files = []

# --- STEP 1: CLUB ORDERS ---
st.header("Step 1: Upload Club Orders")
st.caption("Upload one or more club order files")
club_files = st.file_uploader("Choose Club Orders Excel file(s)", type="xlsx", accept_multiple_files=True)

if club_files:
    st.session_state.club_files = club_files
    st.success("Club Orders uploaded. Proceed to Step 2")
    st.session_state.step = 2

# --- STEP 2: PRODUCTION DATA ---
if st.session_state.step >= 2:
    st.header("Step 2: Upload Production Data")
    st.caption("Upload BOTH the Print and C&P files")
    prod_files = st.file_uploader("Choose Production Excel files (2 required)", type="xlsx", accept_multiple_files=True)

    if prod_files and len(prod_files) == 2:
        st.session_state.prod_files = prod_files
        st.success("Production files uploaded. Proceed to generate output")
        st.session_state.step = 3
    elif prod_files and len(prod_files) != 2:
        st.error("Please upload exactly 2 production files: 1 Print and 1 C&P file")

# --- STEP 3: PROCESS & DOWNLOAD ---
if st.session_state.step == 3:
    if st.button("Generate Combined Output"):
        with st.spinner("Processing your files..."):
            try:
                today = datetime.today().strftime("%Y-%m-%d")

                # --- Load club files ---
                club_dfs = []
                for f in st.session_state.club_files:
                    df = pd.read_excel(f, header=1)
                    # Calculate total sell price from all price columns
                    price_columns = [
                        'Sell Price (DT)',
                        'Sell Price (FT)',
                        'Sell Price (STK)',
                        'Sell Price (STA)',
                        'Sell Price (RFP)'
                    ]
                    df['Total_Print_Sell'] = df[price_columns].sum(axis=1, skipna=True)
                    club_dfs.append(df)
                
                club_df = pd.concat(club_dfs, ignore_index=True)
                club_df.columns = club_df.columns.str.strip()

                # --- Load production files ---
                df1 = pd.read_excel(st.session_state.prod_files[0], header=1)
                df2 = pd.read_excel(st.session_state.prod_files[1], header=1)
                df1.columns = df1.columns.str.strip()
                df2.columns = df2.columns.str.strip()

                df_cp = df1 if "Collate And Pack Cost Price" in df1.columns else df2
                df_print = df2 if df_cp is df1 else df1

                cp_lookup = df_cp.set_index("Project Ref")["Collate And Pack Cost Price"].to_dict()

                # --- Transform Club Orders ---
                club_rows = []
                index_counter = 1

                for _, row in club_df.iterrows():
                    part = row.get("Part")
                    part_info = PARTS_DATA.get(part, {})

                    try:
                        p_val = float(row.get("Total_Print_Sell", 0))
                        cp_val = 2.35
                        d_val = 5.33
                    except:
                        p_val, cp_val, d_val = 0, 2.35, 5.33

                    # Skip this row if total print sell is zero
                    if p_val == 0:
                        continue

                    # Row 1
                    club_rows.append({
                        "index no": index_counter,
                        "Matrix": "",
                        "Matrix URN": "",
                        "Order type": "Club",
                        "Project Ref": row.get("Local Marketing Order Ref", ""),
                        "Project name": "Club",
                        "Brief ref": row.get("Local Marketing Order Line Ref", ""),
                        "Product": part,
                        "Size": part_info.get("Size", ""),
                        "Pagination": part_info.get("Pagination", ""),
                        "Material": part_info.get("Material", ""),
                        "Finishing": part_info.get("Finishing", ""),
                        "Quantity": row.get("Quantity", ""),
                        "Print Matrix": "",
                        "Print Sell": p_val,
                        "Sell": "",
                        "Number of clubs": 1,
                        "Comment": "",
                        "ERA Comments": "",
                        "ITG Comment": "",
                        "Credit": ""
                    })

                    # Row 2: C&P
                    club_rows.append({
                        "index no": index_counter,
                        "Matrix": "",
                        "Matrix URN": "",
                        "Order type": "Club",
                        "Project Ref": row.get("Local Marketing Order Ref", ""),
                        "Project name": "Club",
                        "Brief ref": row.get("Local Marketing Order Line Ref", ""),
                        "Product": "C&P",
                        "Size": "",
                        "Pagination": "",
                        "Material": "",
                        "Finishing": "C&P",
                        "Quantity": 1,
                        "Print Matrix": "",
                        "Print Sell": cp_val,
                        "Sell": "",
                        "Number of clubs": 1,
                        "Comment": "",
                        "ERA Comments": "",
                        "ITG Comment": "",
                        "Credit": ""
                    })

                    # Row 3: Delivery
                    club_rows.append({
                        "index no": index_counter,
                        "Matrix": "",
                        "Matrix URN": "",
                        "Order type": "Club",
                        "Project Ref": row.get("Local Marketing Order Ref", ""),
                        "Project name": "Club",
                        "Brief ref": row.get("Local Marketing Order Line Ref", ""),
                        "Product": "Delivery",
                        "Size": "",
                        "Pagination": "",
                        "Material": "",
                        "Finishing": "Delivery",
                        "Quantity": 1,
                        "Print Matrix": "",
                        "Print Sell": d_val,
                        "Sell": p_val + cp_val + d_val,
                        "Number of clubs": 1,
                        "Comment": "",
                        "ERA Comments": "",
                        "ITG Comment": "",
                        "Credit": ""
                    })

                    index_counter += 1

                df_club_out = pd.DataFrame(club_rows)

                # --- Transform Production Data ---
                grouped = df_print.groupby("Project Ref")
                prod_rows = []

                for project_ref, group in grouped:
                    current_index = index_counter
                    project_name = group["Project Description"].iloc[0]
                    cp_cost = cp_lookup.get(project_ref, "NOT FOUND")
                    total_sell = 0

                    for _, row in group.iterrows():
                        try:
                            sell = float(row.get("Production Sell Price", 0))
                        except:
                            sell = 0
                        total_sell += sell

                        prod_rows.append({
                            "index no": current_index,
                            "Matrix": "Matrix",
                            "Matrix URN": "",
                            "Order type": "Camp / Misc",
                            "Project Ref": row.get("Project Ref", ""),
                            "Project name": row.get("Project Description", ""),
                            "Brief ref": row.get("Brief Ref", ""),
                            "Product": row.get("Part", ""),
                            "Size": f"{row.get('Height', '')}x{row.get('Width', '')}",
                            "Pagination": row.get("No of Pages", ""),
                            "Material": row.get("Material", ""),
                            "Finishing": row.get("Production Finishing Notes", ""),
                            "Quantity": row.get("Total including Spares", ""),
                            "Print Matrix": "",
                            "Print Sell": sell,
                            "Sell": "",
                            "Number of clubs": row.get("No of Clubs", ""),
                            "Comment": "",
                            "ERA Comments": "",
                            "ITG Comment": "",
                            "Credit": ""
                        })

                    prod_rows.append({
                        "index no": current_index,
                        "Matrix": "Matrix",
                        "Matrix URN": "",
                        "Order type": "Camp / Misc",
                        "Project Ref": project_ref,
                        "Project name": project_name,
                        "Brief ref": "",
                        "Product": "C&P",
                        "Size": "",
                        "Pagination": "",
                        "Material": "",
                        "Finishing": "C&P",
                        "Quantity": "",
                        "Print Matrix": "",
                        "Print Sell": cp_cost,
                        "Sell": total_sell + (cp_cost if isinstance(cp_cost, (int, float)) else 0),
                        "Number of clubs": group["No of Clubs"].iloc[0],
                        "Comment": "",
                        "ERA Comments": "",
                        "ITG Comment": "",
                        "Credit": ""
                    })

                    index_counter += 1

                df_prod_out = pd.DataFrame(prod_rows)

                # --- Combine outputs ---
                final_df = pd.concat([df_club_out, df_prod_out], ignore_index=True)

                # Save to buffer
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    final_df.to_excel(writer, sheet_name='Combined Data', index=False)
                    
                    # Get the xlsxwriter workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets['Combined Data']
                    
                    # Define formats
                    border_fmt = workbook.add_format({
                        'border': 1,  # Thin border
                        'border_color': '#000000'  # Black color
                    })
                    
                    header_fmt = workbook.add_format({
                        'border': 1,
                        'border_color': '#000000',
                        'bold': True,
                        'text_wrap': True,
                        'valign': 'top'
                    })
                    
                    # Apply formats to the header row
                    for col_num, value in enumerate(final_df.columns.values):
                        worksheet.write(0, col_num, value, header_fmt)
                    
                    # Auto-fit column widths based on data
                    for idx, col in enumerate(final_df.columns):
                        series = final_df[col]
                        max_len = max(
                            series.astype(str).apply(len).max(),  # max length of values
                            len(str(series.name))  # length of column name
                        ) + 2  # adding a little extra space
                        worksheet.set_column(idx, idx, max_len)
                    
                    # Apply borders to the data range
                    worksheet.conditional_format(
                        0, 0, len(final_df), len(final_df.columns)-1,
                        {'type': 'no_blanks', 'format': border_fmt}
                    )
                    worksheet.conditional_format(
                        0, 0, len(final_df), len(final_df.columns)-1,
                        {'type': 'blanks', 'format': border_fmt}
                    )
                
                buffer.seek(0)
                st.download_button(
                    label="Download Combined Data",
                    data=buffer,
                    file_name=f"ERA_Combined_Data_{today}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                st.success("Data processing complete! Click the button above to download.")
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
