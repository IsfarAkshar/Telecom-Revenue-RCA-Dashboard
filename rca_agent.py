import os
import pandas as pd
import matplotlib.pyplot as plt
from transformers import pipeline


def read_multiple_tables(file_path, sheet_name=0):
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
    table.columns = table.iloc[0].astype(str).str.strip()
    table = table.drop(table.index[0]).reset_index(drop=True)
    table.columns = table.columns.str.strip()
    return table


def compute_rca_for_table(table):
    for col in ['Pre', 'Post', 'Absolute Change', '% Change']:
        if col in table.columns:
            table[col] = pd.to_numeric(
                table[col].astype(str).str.replace(',', '').str.replace('%', ''),
                errors='coerce'
            )

    section_col = table.columns[0]
    totals_df = table[table[section_col] == 'X']

    if totals_df.empty:
        total_abs_change = table['Absolute Change'].dropna().sum()
        total_post = table['Post'].dropna().sum()
    else:
        totals = totals_df.iloc[0]
        total_abs_change = totals['Absolute Change']
        total_post = totals['Post']

    if not totals_df.empty:
        valid_rows = table[
            (table[section_col] != 'X')
            & table['Absolute Change'].notna()
            & table['Post'].notna()
        ]
    else:
        valid_rows = table[
            table['Absolute Change'].notna()
            & table['Post'].notna()
        ]

    valid_rows['Contribution to Absolute Change (%)'] = (
        valid_rows['Absolute Change'] / total_abs_change
    ) * 100

    valid_rows['Contribution to Post (%)'] = (
        valid_rows['Post'] / total_post
    ) * 100

    valid_rows['Combined Impact Score'] = (
        valid_rows['Contribution to Absolute Change (%)'].abs()
        + valid_rows['Contribution to Post (%)'].abs()
    )

    valid_rows['RCA Priority'] = valid_rows['Combined Impact Score'].rank(
        ascending=False
    )

    valid_rows['Section'] = section_col

    return valid_rows


def process_rca(tables):
    processed_tables = []

    for table in tables:
        table = clean_and_prepare_table(table)
        expected_cols = {'Pre', 'Post', 'Absolute Change'}
        if not expected_cols.issubset(set(table.columns)):
            continue

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


def add_kpi_label_column(rca_df):
    rca_df = rca_df.copy()

    def get_label(row):
        section = row['Section']
        value = row.get(section, None)
        return f"{section}: {value}"

    rca_df['KPI Segment Label'] = rca_df.apply(get_label, axis=1)
    return rca_df


def plot_rca_drivers(rca_df, top_n=10, output_folder="output"):
    os.makedirs(output_folder, exist_ok=True)
    rca_df = rca_df.copy()

    # Top positive contributors
    pos = rca_df[rca_df['Contribution to Absolute Change (%)'] > 0]
    pos = pos.sort_values(
        'Contribution to Absolute Change (%)',
        ascending=False
    ).head(top_n)

    plt.figure(figsize=(12, 6))
    plt.bar(
        pos['KPI Segment Label'],
        pos['Contribution to Absolute Change (%)'],
        color='green'
    )
    plt.xticks(rotation=45, ha='right')
    plt.xlabel('KPI Segment')
    plt.ylabel('Contribution to Absolute Change (%)')
    plt.title('Top Positive Revenue Drivers')
    plt.tight_layout()
    chart_path_pos = os.path.join(output_folder, "rca_top_positive_drivers.png")
    plt.savefig(chart_path_pos)
    plt.close()

    # Top negative contributors
    neg = rca_df[rca_df['Contribution to Absolute Change (%)'] < 0]
    neg = neg.sort_values(
        'Contribution to Absolute Change (%)'
    ).head(top_n)

    plt.figure(figsize=(12, 6))
    plt.bar(
        neg['KPI Segment Label'],
        neg['Contribution to Absolute Change (%)'],
        color='red'
    )
    plt.xticks(rotation=45, ha='right')
    plt.xlabel('KPI Segment')
    plt.ylabel('Contribution to Absolute Change (%)')
    plt.title('Top Negative Revenue Drivers')
    plt.tight_layout()
    chart_path_neg = os.path.join(output_folder, "rca_top_negative_drivers.png")
    plt.savefig(chart_path_neg)
    plt.close()

    return chart_path_pos, chart_path_neg


def summarize_rca(rca_df):
    summarizer = pipeline(
        'summarization',
        model='sshleifer/distilbart-cnn-12-6'
    )

    top_rows = rca_df.sort_values('RCA Priority').head(10)
    lines = []
    for _, row in top_rows.iterrows():
        label = row['KPI Segment Label']
        delta = row['Absolute Change']
        contrib = row['Contribution to Absolute Change (%)']
        lines.append(
            f"{label} | Change {delta:.1f} | Contribution {contrib:.2f}%"
        )

    text = "Root Cause Analysis Drivers:\n" + "\n".join(lines)

    summary = summarizer(
        text,
        max_length=80,
        min_length=25,
        do_sample=False
    )[0]['summary_text']

    return summary


def main():
    file_path = "sample.xlsx"  # put your Excel in the same folder
    tables = read_multiple_tables(file_path)
    rca_results = process_rca(tables)

    if rca_results.empty:
        print("No valid RCA analysis found.")
        return

    rca_results = add_kpi_label_column(rca_results)

    os.makedirs("output", exist_ok=True)
    output_path = os.path.join("output", "rca_results.xlsx")
    rca_results.to_excel(output_path, index=False)
    print("RCA results written to", output_path)

    chart_pos, chart_neg = plot_rca_drivers(
        rca_results,
        top_n=10,
        output_folder="output"
    )
    print(f"Visualization saved: {chart_pos} and {chart_neg}")

    summary = summarize_rca(rca_results)
    print("Self-hosted RCA Summary:\n", summary)


if __name__ == "__main__":
    main()
