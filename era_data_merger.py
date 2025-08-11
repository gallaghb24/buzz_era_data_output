import streamlit as st
import pandas as pd
import io
import warnings
from datetime import datetime
from parts_data import PARTS_DATA_RAW

# Constants for general report parsing
OLD_PRICE_COLUMNS = [
    "Sell Price (DT)",
    "Sell Price (FT)",
    "Sell Price (STK)",
    "Sell Price (STA)",
    "Sell Price (RFP)",
]

BASE_COLUMNS = [
    "Order Owner",
    "Order Status",
    "Local Marketing Order Ref",
    "Stock Order Ref",
    "Order Placed Date",
    "Order Line Reference",
    "Local Marketing Asset",
    "Local Marketing Order Line Ref",
    "Stock Item",
    "Stock Order Line Ref",
    "Part",
    "Quantity",
    "Date Approved",
    "Location",
    "Workflow Reference Number",
]

COST_COLUMNS = [
    "If tender pre 5.25%",
    "Print",
    "Collate & pack",
    "Despatch",
    "Total",
]

FINAL_COLUMNS = BASE_COLUMNS + COST_COLUMNS


def parse_general_report(df: pd.DataFrame) -> pd.DataFrame:
    """Parse a general_report sheet into a unified schema."""

    # --- Promote header row ---
    header_row_idx = None
    for i in range(len(df)):
        if df.iloc[i].notna().any():
            header_row_idx = i
            break
    if header_row_idx is None:
        raise ValueError("Could not determine header row")

    header = df.iloc[header_row_idx].astype(str).str.strip()
    df = df.iloc[header_row_idx + 1 :].reset_index(drop=True)
    df.columns = header
    df.columns = df.columns.astype(str).str.strip()
    df = df.dropna(how="all")

    # --- Detect format ---
    is_old = any(col in df.columns for col in OLD_PRICE_COLUMNS)
    is_new = ("Total" in df.columns) and any(
        col in df.columns for col in ["Print", "Collate & pack", "Despatch"]
    )
    if not (is_old or is_new):
        raise ValueError(
            "Unrecognized general_report format. Expected Sell Price columns or Total with Print/Collate & pack/Despatch columns."
        )

    # --- Validate required columns ---
    missing = [c for c in BASE_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    df = df.copy()

    # --- Map to unified schema ---
    if is_old:
        # Create cost columns with zeros
        for col in ["If tender pre 5.25%", "Print", "Collate & pack", "Despatch"]:
            df[col] = 0
        df["Total"] = (
            df[OLD_PRICE_COLUMNS].apply(pd.to_numeric, errors="coerce").sum(axis=1, skipna=True)
        )
    else:
        # Ensure all cost columns exist
        for col in COST_COLUMNS:
            if col not in df.columns:
                df[col] = 0

    # --- Type coercion ---
    for col in COST_COLUMNS:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    df["Order Placed Date"] = pd.to_datetime(
        df["Order Placed Date"], dayfirst=True, errors="coerce"
    )
    df["Date Approved"] = pd.to_datetime(
        df["Date Approved"], dayfirst=True, errors="coerce"
    )

    quantity_num = pd.to_numeric(df["Quantity"], errors="coerce")
    if quantity_num.isna().any():
        warnings.warn("Non-numeric Quantity values coerced to 0")
    df["Quantity"] = quantity_num.fillna(0).astype(int)

    if (df["Total"] < 0).any():
        warnings.warn("Negative Total values found")

    df = df[FINAL_COLUMNS]
    return df


def main():
    st.set_page_config(page_title="ERA Data Merger", layout="centered")
    st.title("🔄 ERA Club & Production Data Merger")

    # --- PARTS LOOKUP TRANSFORM ---
    PARTS_DATA = {}
    for part_data in PARTS_DATA_RAW.values():
        PARTS_DATA[part_data["Part Name"]] = {
            "Size": f"{part_data['Height (mm)']}x{part_data['Width (mm)']}",
            "Pagination": part_data["No. of Pages"],
            "Material": part_data["Materials"],
            "Finishing": part_data.get("Finishing", ""),
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
    club_files = st.file_uploader(
        "Choose Club Orders Excel file(s)", type="xlsx", accept_multiple_files=True
    )

    if club_files:
        st.session_state.club_files = club_files
        st.success("Club Orders uploaded. Proceed to Step 2")
        st.session_state.step = 2

    # --- STEP 2: PRODUCTION DATA ---
    if st.session_state.step >= 2:
        st.header("Step 2: Upload Production Data")
        st.caption("Upload BOTH the Print and C&P files")
        prod_files = st.file_uploader(
            "Choose Production Excel files (2 required)",
            type="xlsx",
            accept_multiple_files=True,
        )

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
                        raw_df = pd.read_excel(f, header=None)
                        df = parse_general_report(raw_df)
                        club_dfs.append(df)

                    club_df = pd.concat(club_dfs, ignore_index=True)

                    # --- Load production files ---
                    df1 = pd.read_excel(st.session_state.prod_files[0], header=1)
                    df2 = pd.read_excel(st.session_state.prod_files[1], header=1)
                    df1.columns = df1.columns.str.strip()
                    df2.columns = df2.columns.str.strip()

                    df_cp = df1 if "Collate And Pack Cost Price" in df1.columns else df2
                    df_print = df2 if df_cp is df1 else df1

                    cp_lookup = df_cp.set_index("Project Ref")[
                        "Collate And Pack Cost Price"
                    ].to_dict()

                    # --- Transform Club Orders ---
                    club_rows = []
                    index_counter = 1

                    for _, row in club_df.iterrows():
                        part = row.get("Part")
                        part_info = PARTS_DATA.get(part, {})

                        print_val = row.get("Print", 0)
                        cp_raw = row.get("Collate & pack", 0)
                        d_raw = row.get("Despatch", 0)

                        if pd.isna(print_val):
                            print_val = 0

                        if print_val == 0 and cp_raw == 0 and d_raw == 0:
                            print_val = row.get("Total", 0)

                        cp_val = cp_raw if cp_raw != 0 else 2.35
                        d_val = d_raw if d_raw != 0 else 5.33

                        sell_total = row.get("Total", 0)
                        if pd.isna(sell_total) or sell_total == 0:
                            sell_total = print_val + cp_val + d_val

                        # Row 1
                        club_rows.append(
                            {
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
                                "Print Sell": print_val,
                                "Sell": "",
                                "Number of clubs": 1,
                                "Comment": "",
                                "ERA Comments": "",
                                "ITG Comment": "",
                                "Credit": "",
                            }
                        )

                        # Row 2: C&P
                        club_rows.append(
                            {
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
                                "Credit": "",
                            }
                        )

                        # Row 3: Delivery
                        club_rows.append(
                            {
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
                                "Sell": sell_total,
                                "Number of clubs": 1,
                                "Comment": "",
                                "ERA Comments": "",
                                "ITG Comment": "",
                                "Credit": "",
                            }
                        )

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
                            except Exception:
                                sell = 0
                            total_sell += sell

                            prod_rows.append(
                                {
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
                                    "Credit": "",
                                }
                            )

                        prod_rows.append(
                            {
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
                                "Sell": total_sell
                                + (cp_cost if isinstance(cp_cost, (int, float)) else 0),
                                "Number of clubs": group["No of Clubs"].iloc[0],
                                "Comment": "",
                                "ERA Comments": "",
                                "ITG Comment": "",
                                "Credit": "",
                            }
                        )

                        index_counter += 1

                    df_prod_out = pd.DataFrame(prod_rows)

                    # --- Combine outputs ---
                    final_df = pd.concat([df_club_out, df_prod_out], ignore_index=True)

                    # Save to buffer
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        final_df.to_excel(writer, sheet_name="Combined Data", index=False)

                        workbook = writer.book
                        worksheet = writer.sheets["Combined Data"]

                        border_fmt = workbook.add_format(
                            {
                                "border": 1,
                                "border_color": "#000000",
                            }
                        )

                        header_fmt = workbook.add_format(
                            {
                                "border": 1,
                                "border_color": "#000000",
                                "bold": True,
                                "text_wrap": True,
                                "valign": "top",
                            }
                        )

                        for col_num, value in enumerate(final_df.columns.values):
                            worksheet.write(0, col_num, value, header_fmt)

                        for idx, col in enumerate(final_df.columns):
                            series = final_df[col]
                            max_len = max(
                                series.astype(str).apply(len).max(),
                                len(str(series.name)),
                            ) + 2
                            worksheet.set_column(idx, idx, max_len)

                        worksheet.conditional_format(
                            0,
                            0,
                            len(final_df),
                            len(final_df.columns) - 1,
                            {"type": "no_blanks", "format": border_fmt},
                        )
                        worksheet.conditional_format(
                            0,
                            0,
                            len(final_df),
                            len(final_df.columns) - 1,
                            {"type": "blanks", "format": border_fmt},
                        )

                    buffer.seek(0)
                    st.download_button(
                        label="Download Combined Data",
                        data=buffer,
                        file_name=f"ERA_Combined_Data_{today}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                    st.success("Data processing complete! Click the button above to download.")
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")


if __name__ == "__main__":
    main()

