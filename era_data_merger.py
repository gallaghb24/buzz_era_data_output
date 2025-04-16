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
                club_dfs = [pd.read_excel(f, header=1) for f in st.session_state.club_files]
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
                        p_val = float(row.get("Sell Price (FT)", 0))
                        cp_val = 2.35
                        d_val = 5.33
                    except:
                        p_val, cp_val, d_val = 0, 2.35, 5.33

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