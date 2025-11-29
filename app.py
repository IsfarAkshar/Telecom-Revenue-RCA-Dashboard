# import streamlit as st
# import pandas as pd
# import os
# from io import BytesIO

# # import your RCA functions from rca_agent.py
# from rca_agent_new import (
#     read_multiple_tables,
#     process_rca,
#     add_kpi_label_column,
#     plot_rca_drivers,
#     generate_structured_rca_text,
#     save_rca_text,
# )


# st.set_page_config(
#     page_title="Telecom Revenue RCA Dashboard",
#     layout="wide"
# )

# st.title("Telecom Revenue RCA Dashboard")

# st.sidebar.header("1. Upload KPI file")
# uploaded_file = st.sidebar.file_uploader(
#     "Upload Excel (MTD vs LMTD)",
#     type=["xlsx"]
# )

# sheet_name = st.sidebar.text_input(
#     "Sheet name or index (e.g., 0 for first sheet, 'Robi')",
#     value="0"
# )

# top_n = st.sidebar.slider(
#     "Top drivers in charts",
#     min_value=5,
#     max_value=20,
#     value=10,
#     step=1
# )

# brand_name = st.sidebar.text_input(
#     "Brand name for narrative",
#     value="Robi"
# )

# run_button = st.sidebar.button("Run RCA")


# def parse_sheet_name(raw):
#     try:
#         return int(raw)
#     except ValueError:
#         return raw  # treat as sheet name string


# if run_button and uploaded_file is not None:
#     st.info("Running RCA analysis...")

#     # Streamlit gives a file-like object; pandas.read_excel can handle it
#     # directly or via BytesIO
#     file_bytes = uploaded_file.read()
#     excel_io = BytesIO(file_bytes)

#     tables = read_multiple_tables(excel_io, sheet_name=parse_sheet_name(sheet_name))
#     rca_results = process_rca(tables)

#     if rca_results.empty:
#         st.error("No valid RCA analysis found in this sheet.")
#     else:
#         rca_results = add_kpi_label_column(rca_results)

#         # Ensure output folder exists
#         os.makedirs("output_new", exist_ok=True)

#         # Save RCA results Excel and insights text
#         excel_path = os.path.join("output_new", "rca_results_streamlit.xlsx")
#         rca_results.to_excel(excel_path, index=False)

#         rca_text = generate_structured_rca_text(rca_results, brand_name=brand_name)
#         txt_path = save_rca_text(
#             rca_text,
#             output_folder="output_new",
#             filename="rca_insights_streamlit.txt"
#         )

#         # Layout: 2 columns for metrics & narrative
#         col1, col2 = st.columns([1, 1])

#         with col1:
#             st.subheader("RCA Narrative (Key Factors)")
#             st.text(rca_text)

#         with col2:
#             st.subheader("RCA Results Preview")
#             st.dataframe(
#                 rca_results.head(30),
#                 use_container_width=True
#             )

#         # Charts
#         st.subheader("Top Revenue Drivers")
#         chart_pos, chart_neg = plot_rca_drivers(
#             rca_results,
#             top_n=top_n,
#             output_folder="output_new"
#         )
#         st.write("Top Positive Drivers")
#         st.image(chart_pos, use_column_width=True)
#         st.write("Top Negative Drivers")
#         st.image(chart_neg, use_column_width=True)

#         # Download buttons
#         st.subheader("Downloads")
#         with open(excel_path, "rb") as f:
#             st.download_button(
#                 label="Download RCA Results (Excel)",
#                 data=f,
#                 file_name="rca_results.xlsx",
#                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#             )

#         with open(txt_path, "rb") as f:
#             st.download_button(
#                 label="Download RCA Insights (Text)",
#                 data=f,
#                 file_name="rca_insights.txt",
#                 mime="text/plain"
#             )

# elif run_button and uploaded_file is None:
#     st.warning("Please upload an Excel file first.")


import streamlit as st
import pandas as pd
import os
from io import BytesIO

from rca_agent_new import (
    read_multiple_tables,
    process_rca,
    add_kpi_label_column,
    plot_rca_drivers,
    generate_structured_rca_text,
    save_rca_text,
)


st.set_page_config(
    page_title="Telecom Revenue RCA Dashboard",
    layout="wide"
)


def parse_sheet_name(raw):
    try:
        return int(raw)
    except ValueError:
        return raw


def build_key_column(df):
    df = df.copy()
    df["Key"] = df["Section"].astype(str) + " | " + df["KPI Segment Label"].astype(str)
    return df


st.title("Telecom Revenue RCA Dashboard")

tab_single, tab_compare = st.tabs(["Single File RCA", "Compare Two Files"])

# ---------------- SINGLE FILE RCA TAB ----------------

with tab_single:
    st.sidebar.header("Single File Settings")

    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel (MTD vs LMTD)",
        type=["xlsx"],
        key="single_uploader"
    )

    sheet_name = st.sidebar.text_input(
        "Sheet name or index (e.g., 0 or 'Robi')",
        value="0",
        key="single_sheet"
    )

    top_n_single = st.sidebar.slider(
        "Top drivers in charts",
        min_value=5,
        max_value=20,
        value=10,
        step=1,
        key="single_topn"
    )

    brand_name = st.sidebar.text_input(
        "Brand name for narrative",
        value="Robi",
        key="single_brand"
    )

    run_single = st.sidebar.button("Run RCA (single file)")

    if run_single and uploaded_file is not None:
        st.info("Running RCA for single file...")

        file_bytes = uploaded_file.read()
        excel_io = BytesIO(file_bytes)

        tables = read_multiple_tables(excel_io, sheet_name=parse_sheet_name(sheet_name))
        rca_results = process_rca(tables)

        if rca_results.empty:
            st.error("No valid RCA analysis found in this sheet.")
        else:
            rca_results = add_kpi_label_column(rca_results)
            os.makedirs("output", exist_ok=True)

            excel_path = os.path.join("output", "rca_results_single.xlsx")
            rca_results.to_excel(excel_path, index=False)

            rca_text = generate_structured_rca_text(rca_results, brand_name=brand_name)
            txt_path = save_rca_text(
                rca_text,
                output_folder="output",
                filename="rca_insights_single.txt"
            )

            col1, col2 = st.columns([1, 1])

            with col1:
                st.subheader("RCA Narrative (Key Factors)")
                st.text(rca_text)

            with col2:
                st.subheader("RCA Results Preview")
                st.dataframe(rca_results.head(30), use_container_width=True)

            st.subheader("Top Revenue Drivers")
            chart_pos, chart_neg = plot_rca_drivers(
                rca_results,
                top_n=top_n_single,
                output_folder="output"
            )
            st.write("Top Positive Drivers")
            st.image(chart_pos, use_column_width=True)
            st.write("Top Negative Drivers")
            st.image(chart_neg, use_column_width=True)

            st.subheader("Downloads")
            with open(excel_path, "rb") as f:
                st.download_button(
                    label="Download RCA Results (Excel)",
                    data=f,
                    file_name="rca_results_single.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            with open(txt_path, "rb") as f:
                st.download_button(
                    label="Download RCA Insights (Text)",
                    data=f,
                    file_name="rca_insights_single.txt",
                    mime="text/plain"
                )
    elif run_single and uploaded_file is None:
        st.warning("Please upload an Excel file first for the single-file analysis.")


# ---------------- COMPARE TWO FILES TAB ----------------

# with tab_compare:
#     st.sidebar.header("Comparison Settings")

#     uploaded_file_a = st.sidebar.file_uploader(
#         "Upload Excel A",
#         type=["xlsx"],
#         key="cmp_uploader_a"
#     )
#     sheet_a = st.sidebar.text_input(
#         "Sheet for File A (index or name)",
#         value="0",
#         key="cmp_sheet_a"
#     )

#     uploaded_file_b = st.sidebar.file_uploader(
#         "Upload Excel B",
#         type=["xlsx"],
#         key="cmp_uploader_b"
#     )
#     sheet_b = st.sidebar.text_input(
#         "Sheet for File B (index or name)",
#         value="0",
#         key="cmp_sheet_b"
#     )

#     top_n_cmp = st.sidebar.slider(
#         "Top segments by change in contribution",
#         min_value=5,
#         max_value=30,
#         value=10,
#         step=1,
#         key="cmp_topn"
#     )

#     run_compare = st.sidebar.button("Run RCA Comparison")

#     if run_compare:
#         if uploaded_file_a is None or uploaded_file_b is None:
#             st.warning("Please upload both Excel files for comparison.")
#         else:
#             st.info("Running RCA on both files and building comparison...")

#             # --- File A ---
#             bytes_a = uploaded_file_a.read()
#             excel_io_a = BytesIO(bytes_a)
#             tables_a = read_multiple_tables(excel_io_a, sheet_name=parse_sheet_name(sheet_a))
#             rca_a = process_rca(tables_a)

#             # --- File B ---
#             bytes_b = uploaded_file_b.read()
#             excel_io_b = BytesIO(bytes_b)
#             tables_b = read_multiple_tables(excel_io_b, sheet_name=parse_sheet_name(sheet_b))
#             rca_b = process_rca(tables_b)

#             if rca_a.empty or rca_b.empty:
#                 st.error("RCA results were empty for one or both files.")
#             else:
#                 rca_a = add_kpi_label_column(rca_a)
#                 rca_b = add_kpi_label_column(rca_b)

#                 rca_a = build_key_column(rca_a)
#                 rca_b = build_key_column(rca_b)

#                 # select core columns for comparison
#                 cols_core = [
#                     "Key",
#                     "Section",
#                     "KPI Segment Label",
#                     "Absolute Change",
#                     "Contribution to Absolute Change (%)",
#                     "Contribution to Post (%)"
#                 ]

#                 df_a = rca_a[cols_core].copy()
#                 df_b = rca_b[cols_core].copy()

#                 df_a = df_a.rename(
#                     columns={
#                         "Absolute Change": "AbsChange_A",
#                         "Contribution to Absolute Change (%)": "ContribAbs_A",
#                         "Contribution to Post (%)": "ContribPost_A",
#                     }
#                 )
#                 df_b = df_b.rename(
#                     columns={
#                         "Absolute Change": "AbsChange_B",
#                         "Contribution to Absolute Change (%)": "ContribAbs_B",
#                         "Contribution to Post (%)": "ContribPost_B",
#                     }
#                 )

#                 # outer join on Key
#                 cmp_df = pd.merge(
#                     df_a,
#                     df_b,
#                     on=["Key", "Section", "KPI Segment Label"],
#                     how="outer"
#                 )

#                 # compute deltas (A - B)
#                 cmp_df["Delta_AbsChange"] = cmp_df["AbsChange_A"].fillna(0) - cmp_df["AbsChange_B"].fillna(0)
#                 cmp_df["Delta_ContribAbs"] = cmp_df["ContribAbs_A"].fillna(0) - cmp_df["ContribAbs_B"].fillna(0)
#                 cmp_df["Delta_ContribPost"] = cmp_df["ContribPost_A"].fillna(0) - cmp_df["ContribPost_B"].fillna(0)

#                 # sort by magnitude of change in contribution
#                 cmp_sorted = cmp_df.sort_values("Delta_ContribAbs", key=lambda s: s.abs(), ascending=False)

#                 st.subheader("Top segments where contribution changed most")
#                 st.dataframe(
#                     cmp_sorted.head(top_n_cmp),
#                     use_container_width=True
#                 )

#                 # quick textual summary for top few
#                 st.subheader("Key comparison insights (top segments)")
#                 lines = []
#                 for _, row in cmp_sorted.head(top_n_cmp).iterrows():
#                     label = row["KPI Segment Label"]
#                     d_contrib = row["Delta_ContribAbs"]
#                     d_abs = row["Delta_AbsChange"]
#                     lines.append(
#                         f"{label}: ΔAbsChange={d_abs:+,.0f}, ΔContribution={d_contrib:+.2f} pts"
#                     )
#                 st.text("\n".join(lines))

#                 # allow download of comparison table
#                 os.makedirs("output", exist_ok=True)
#                 cmp_path = os.path.join("output", "rca_comparison.xlsx")
#                 cmp_sorted.to_excel(cmp_path, index=False)

#                 with open(cmp_path, "rb") as f:
#                     st.download_button(
#                         label="Download full comparison (Excel)",
#                         data=f,
#                         file_name="rca_comparison.xlsx",
#                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#                    )


# ---------------- COMPARE TWO FILES TAB ----------------

with tab_compare:
    st.sidebar.header("Comparison Settings")

    uploaded_file_a = st.sidebar.file_uploader(
        "Upload Excel A",
        type=["xlsx"],
        key="cmp_uploader_a"
    )
    sheet_a = st.sidebar.text_input(
        "Sheet for File A (index or name)",
        value="0",
        key="cmp_sheet_a"
    )
    brand_a = st.sidebar.text_input(
        "Brand name for File A",
        value="Brand A",
        key="cmp_brand_a"
    )

    uploaded_file_b = st.sidebar.file_uploader(
        "Upload Excel B",
        type=["xlsx"],
        key="cmp_uploader_b"
    )
    sheet_b = st.sidebar.text_input(
        "Sheet for File B (index or name)",
        value="0",
        key="cmp_sheet_b"
    )
    brand_b = st.sidebar.text_input(
        "Brand name for File B",
        value="Brand B",
        key="cmp_brand_b"
    )

    top_n_cmp = st.sidebar.slider(
        "Top segments by change in contribution",
        min_value=5,
        max_value=30,
        value=10,
        step=1,
        key="cmp_topn"
    )

    run_compare = st.sidebar.button("Run RCA Comparison")

    if run_compare:
        if uploaded_file_a is None or uploaded_file_b is None:
            st.warning("Please upload both Excel files for comparison.")
        else:
            st.info("Running RCA on both files and building comparison...")

            # --- File A ---
            bytes_a = uploaded_file_a.read()
            excel_io_a = BytesIO(bytes_a)
            tables_a = read_multiple_tables(
                excel_io_a,
                sheet_name=parse_sheet_name(sheet_a)
            )
            rca_a = process_rca(tables_a)

            # --- File B ---
            bytes_b = uploaded_file_b.read()
            excel_io_b = BytesIO(bytes_b)
            tables_b = read_multiple_tables(
                excel_io_b,
                sheet_name=parse_sheet_name(sheet_b)
            )
            rca_b = process_rca(tables_b)

            if rca_a.empty or rca_b.empty:
                st.error("RCA results were empty for one or both files.")
            else:
                rca_a = add_kpi_label_column(rca_a)
                rca_b = add_kpi_label_column(rca_b)

                rca_a["Key"] = rca_a["Section"].astype(str) + " | " + rca_a["KPI Segment Label"].astype(str)
                rca_b["Key"] = rca_b["Section"].astype(str) + " | " + rca_b["KPI Segment Label"].astype(str)

                cols_core = [
                    "Key",
                    "Section",
                    "KPI Segment Label",
                    "Absolute Change",
                    "Contribution to Absolute Change (%)",
                    "Contribution to Post (%)"
                ]

                df_a = rca_a[cols_core].copy()
                df_b = rca_b[cols_core].copy()

                df_a = df_a.rename(
                    columns={
                        "Absolute Change": "AbsChange_A",
                        "Contribution to Absolute Change (%)": "ContribAbs_A",
                        "Contribution to Post (%)": "ContribPost_A",
                    }
                )
                df_b = df_b.rename(
                    columns={
                        "Absolute Change": "AbsChange_B",
                        "Contribution to Absolute Change (%)": "ContribAbs_B",
                        "Contribution to Post (%)": "ContribPost_B",
                    }
                )

                cmp_df = pd.merge(
                    df_a,
                    df_b,
                    on=["Key", "Section", "KPI Segment Label"],
                    how="outer"
                )

                cmp_df["Delta_AbsChange"] = cmp_df["AbsChange_A"].fillna(0) - cmp_df["AbsChange_B"].fillna(0)
                cmp_df["Delta_ContribAbs"] = cmp_df["ContribAbs_A"].fillna(0) - cmp_df["ContribAbs_B"].fillna(0)
                cmp_df["Delta_ContribPost"] = cmp_df["ContribPost_A"].fillna(0) - cmp_df["ContribPost_B"].fillna(0)

                cmp_sorted = cmp_df.sort_values(
                    "Delta_ContribAbs",
                    key=lambda s: s.abs(),
                    ascending=False
                )

                # ---------- TABLE WITH GREEN HIGHER VALUE PER BRAND ----------
                st.subheader("Top segments where contribution changed most")

                # Rename contribution columns to show brand names
                disp = cmp_sorted.head(top_n_cmp).copy()
                col_a = f"ContribAbs_{brand_a}"
                col_b = f"ContribAbs_{brand_b}"
                disp = disp.rename(
                    columns={
                        "ContribAbs_A": col_a,
                        "ContribAbs_B": col_b,
                    }
                )

                # Style: make higher contribution (per row) green
                def highlight_higher(row):
                    a = row[col_a]
                    b = row[col_b]
                    styles = []
                    for col in row.index:
                        if col == col_a:
                            styles.append("color: green" if a >= b else "")
                        elif col == col_b:
                            styles.append("color: green" if b > a else "")
                        else:
                            styles.append("")
                    return styles

                styler = disp.style.apply(
                    highlight_higher,
                    axis=1
                )

                st.dataframe(styler, use_container_width=True)

                # ---------- TEXTUAL INSIGHTS WITH BRAND NAMES ----------
                st.subheader("Key comparison insights (top segments)")
                lines = []
                for _, row in cmp_sorted.head(top_n_cmp).iterrows():
                    label = row["KPI Segment Label"]
                    contrib_a = row["ContribAbs_A"] if pd.notna(row["ContribAbs_A"]) else 0
                    contrib_b = row["ContribAbs_B"] if pd.notna(row["ContribAbs_B"]) else 0
                    d_abs = row["Delta_AbsChange"]
                    d_contrib = row["Delta_ContribAbs"]
                    lines.append(
                        f"{label}: "
                        f"{brand_a} Contrib={contrib_a:+.2f}%, "
                        f"{brand_b} Contrib={contrib_b:+.2f}%, "
                        f"ΔAbsChange={d_abs:+,.0f}, ΔContribution={d_contrib:+.2f} pts"
                    )
                st.text("\n".join(lines))

                # ---------- DOWNLOAD FULL COMPARISON ----------
                os.makedirs("output", exist_ok=True)
                cmp_path = os.path.join("output", "rca_comparison.xlsx")
                cmp_sorted.to_excel(cmp_path, index=False)

                with open(cmp_path, "rb") as f:
                    st.download_button(
                        label="Download full comparison (Excel)",
                        data=f,
                        file_name="rca_comparison.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
