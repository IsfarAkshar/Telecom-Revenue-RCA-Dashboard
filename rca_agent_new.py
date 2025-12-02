import os
import pandas as pd
import matplotlib.pyplot as plt


# -------------- DATA LOADING & PREP --------------


def read_multiple_tables(file_path, sheet_name=0):
    """
    Split a sheet into multiple logical tables using blank rows as separators.
    """
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    tables, current_table = [], []
    for _, row in df.iterrows():
        if row.isnull().all():
            if current_table:
                tables.append(pd.DataFrame(current_table).reset_index(drop=True))
                current_table = []
        else:
            current_table.append(row)
    if current_table:
        tables.append(pd.DataFrame(current_table).reset_index(drop=True))
    return tables


def clean_and_prepare_table(table):
    """
    Use the first row as header and clean column names.
    """
    table.columns = table.iloc[0].astype(str).str.strip()
    table = table.drop(table.index[0]).reset_index(drop=True)
    table.columns = table.columns.str.strip()
    return table


def compute_rca_for_table(table):
    """
    Compute contributions, impact scores and RCA priority for a single KPI table.
    Core math stays “pure” (no Multisimmer business overrides here).
    """
    # numeric conversion
    for col in ["Pre", "Post", "Absolute Change", "% Change"]:
        if col in table.columns:
            table[col] = pd.to_numeric(
                table[col].astype(str).str.replace(",", "").str.replace("%", ""),
                errors="coerce",
            )

    section_col = table.columns[0]
    totals_df = table[table[section_col] == "X"]

    if totals_df.empty:
        total_abs_change = table["Absolute Change"].dropna().sum()
        total_post = table["Post"].dropna().sum()
    else:
        totals = totals_df.iloc[0]
        total_abs_change = totals["Absolute Change"]
        total_post = totals["Post"]

    if not totals_df.empty:
        valid_rows = table[
            (table[section_col] != "X")
            & table["Absolute Change"].notna()
            & table["Post"].notna()
        ]
    else:
        valid_rows = table[
            table["Absolute Change"].notna() & table["Post"].notna()
        ]

    # ---- CORE MATH: NO special Multisimmer handling here ----
    valid_rows["Contribution to Absolute Change (%)"] = (
        valid_rows["Absolute Change"] / total_abs_change
    ) * 100

    valid_rows["Contribution to Post (%)"] = (
        valid_rows["Post"] / total_post
    ) * 100

    valid_rows["Combined Impact Score"] = (
        valid_rows["Contribution to Absolute Change (%)"].abs()
        + valid_rows["Contribution to Post (%)"].abs()
    )

    valid_rows["RCA Priority"] = valid_rows["Combined Impact Score"].rank(
        ascending=False
    )

    valid_rows["Section"] = section_col

    return valid_rows


def process_rca(tables):
    """
    Run RCA on all tables and combine them.
    """
    processed_tables = []
    for table in tables:
        table = clean_and_prepare_table(table)
        expected_cols = {"Pre", "Post", "Absolute Change"}
        if not expected_cols.issubset(set(table.columns)):
            continue

        # handle duplicate column names if any
        if table.columns.duplicated().any():
            cols = pd.Series(table.columns)
            for dup in cols[cols.duplicated()].unique():
                dups_idx = cols[cols == dup].index.tolist()
                for i, col_idx in enumerate(dups_idx):
                    if i > 0:
                        cols[col_idx] = f"{dup}_{i}"
            table.columns = cols

        rca_table = compute_rca_for_table(table)
        processed_tables.append(rca_table)

    if not processed_tables:
        return pd.DataFrame()

    combined_df = pd.concat(processed_tables, ignore_index=True)
    return combined_df


# -------------- LABELS --------------


def add_kpi_label_column(rca_df):
    """
    Add a human-readable KPI Segment Label like "Handset Type: Smartphone".
    """
    rca_df = rca_df.copy()

    def get_label(row):
        section = row["Section"]
        value = row.get(section, None)
        return f"{section}: {value}"

    rca_df["KPI Segment Label"] = rca_df.apply(get_label, axis=1)
    return rca_df


# -------------- CHARTS (business view for GP/BL Multisim) --------------


def plot_rca_drivers(rca_df, top_n=10, output_folder="output_new"):
    """
    Plot top positive and negative drivers.

    IMPORTANT:
    - Core math (Contribution to Absolute Change) stays as-is.
    - For charts only, GP_MULTISIM and BL_MULTISIM have inverted sign:
        * If users increase (positive), shown as negative driver in charts.
        * If users decrease (negative), shown as positive driver in charts.
    """
    os.makedirs(output_folder, exist_ok=True)

    # Work on a copy so core RCA math is untouched
    df_plot = rca_df.copy()

    def invert_for_multisim_chart(row):
        label = str(row.get("KPI Segment Label", ""))

        # Only apply special logic to GP_MULTISIM and BL_MULTISIM
        if "GP_MULTISIM" in label or "BL_MULTISIM" in label:
            return -row["Contribution to Absolute Change (%)"]
        else:
            return row["Contribution to Absolute Change (%)"]

    # New column used only for plotting
    df_plot["Contrib_for_chart"] = df_plot.apply(
        invert_for_multisim_chart, axis=1
    )

    # --------------- TOP POSITIVE DRIVERS (chart view) ---------------
    pos = df_plot[df_plot["Contrib_for_chart"] > 0]
    pos = pos.sort_values(
        "Contrib_for_chart", ascending=False
    ).head(top_n)

    plt.figure(figsize=(12, 6))
    plt.bar(
        pos["KPI Segment Label"],
        pos["Contrib_for_chart"],
        color="green",
    )
    plt.xticks(rotation=45, ha="right")
    plt.xlabel("KPI Segment")
    plt.ylabel("Contribution to Absolute Change (%)")
    plt.title("Top Positive Revenue Drivers (business view)")
    plt.tight_layout()
    chart_path_pos = os.path.join(
        output_folder, "rca_top_positive_drivers.png"
    )
    plt.savefig(chart_path_pos, dpi=150)
    plt.close()

    # --------------- TOP NEGATIVE DRIVERS (chart view) ---------------
    neg = df_plot[df_plot["Contrib_for_chart"] < 0]
    neg = neg.sort_values(
        "Contrib_for_chart"
    ).head(top_n)

    plt.figure(figsize=(12, 6))
    plt.bar(
        neg["KPI Segment Label"],
        neg["Contrib_for_chart"],
        color="red",
    )
    plt.xticks(rotation=45, ha="right")
    plt.xlabel("KPI Segment")
    plt.ylabel("Contribution to Absolute Change (%)")
    plt.title("Top Negative Revenue Drivers (business view)")
    plt.tight_layout()
    chart_path_neg = os.path.join(
        output_folder, "rca_top_negative_drivers.png"
    )
    plt.savefig(chart_path_neg, dpi=150)
    plt.close()

    return chart_path_pos, chart_path_neg


# -------------- NARRATIVE (business view for GP/BL Multisim) --------------


def _abs_change_business_view(row, section_col):
    """
    For narrative only:
    - GP_MULTISIM / BL_MULTISIM: invert sign (increase = negative, decrease = positive).
    - Others: keep original Absolute Change.
    """
    label = str(row[section_col])
    ac = row["Absolute Change"]

    if "GP_MULTISIM" in label or "BL_MULTISIM" in label:
        return -ac
    return ac


def format_driver_row(row, section_col):
    # Use business-view Absolute Change for sign in text
    ac_business = _abs_change_business_view(row, section_col)
    pct_change = row["% Change"] * 100
    label = row[section_col]
    return f"{label} ({ac_business:+,.2f} / {pct_change:+.2f}%)"


def get_top_drivers_by_section(rca_df, section, top_n_pos=2, top_n_neg=2):
    df_sec = rca_df[rca_df["Section"] == section].copy()
    if df_sec.empty:
        return [], []

    # Add a business-view column just for narrative ranking
    df_sec["AbsChange_business"] = df_sec.apply(
        lambda r: _abs_change_business_view(r, section), axis=1
    )

    # Positive: largest business-view Absolute Change
    pos = df_sec.sort_values("AbsChange_business", ascending=False).head(top_n_pos)

    # Negative: smallest business-view Absolute Change
    neg = df_sec.sort_values("AbsChange_business", ascending=True).head(top_n_neg)

    pos_list = [format_driver_row(r, section) for _, r in pos.iterrows()]
    neg_list = [format_driver_row(r, section) for _, r in neg.iterrows()]
    return pos_list, neg_list


def generate_structured_rca_text(rca_df, brand_name="Brand"):
    sections_to_use = [
        "Handset Type",
        "Arpu Segment",
        "Usage Category",
        "Gb Slab",
        "Base Type",
        "Multisimmer",
        "Clustername",
        "Mou Slab",
        "Aon Bucket",
        "Vc User Category",
    ]

    lines = []
    lines.append(f"{brand_name}: Key change drivers")
    lines.append("")
    lines.append("Biggest positive impact:")
    for sec in sections_to_use:
        pos, _ = get_top_drivers_by_section(rca_df, sec)
        if pos:
            lines.append(f"- {sec}:")
            for item in pos:
                lines.append(f"  - {item}")

    lines.append("")
    lines.append("Negative impacts / areas to watch:")
    for sec in sections_to_use:
        _, neg = get_top_drivers_by_section(rca_df, sec)
        if neg:
            lines.append(f"- {sec}:")
            for item in neg:
                lines.append(f"  - {item}")

    return "\n".join(lines)


def save_rca_text(text, output_folder="output_new", filename="rca_insights.txt"):
    os.makedirs(output_folder, exist_ok=True)
    path = os.path.join(output_folder, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


# -------------- MAIN (for local testing) --------------


def main():
    file_path = "sample.xlsx"  # adjust if needed
    tables = read_multiple_tables(file_path, sheet_name=0)
    rca_results = process_rca(tables)

    if rca_results.empty:
        print("No valid RCA analysis found.")
        return

    rca_results = add_kpi_label_column(rca_results)

    os.makedirs("output_new", exist_ok=True)
    output_path = os.path.join("output_new", "rca_results.xlsx")
    rca_results.to_excel(output_path, index=False)
    print("RCA results written to", output_path)

    chart_pos, chart_neg = plot_rca_drivers(
        rca_results, top_n=10, output_folder="output_new"
    )
    print(f"Visualization saved: {chart_pos} and {chart_neg}")

    raw_text = generate_structured_rca_text(rca_results, brand_name="Robi")
    print("\n===== RCA NARRATIVE (KEY FACTORS) =====\n")
    print(raw_text)

    txt_path = save_rca_text(
        raw_text, output_folder="output_new", filename="rca_insights.txt"
    )
    print(f"\nRCA insights text file saved to: {txt_path}")


if __name__ == "__main__":
    main()

