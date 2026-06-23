"""
Prepaid Expense Workpaper Automation – Business Logic & Entry Point
"""

import pandas as pd
from openpyxl import load_workbook
import re
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from openpyxl.styles import Font, PatternFill

from gui import PrepaidApp


# ──────────────────────────────────────────────────────────
#  Core processing function (runs in a background thread)
# ──────────────────────────────────────────────────────────
def process_workpaper(basefile, client_name, fy_start, fy_end,
                      method, sample_n, progress):
    """Generate the prepaid-expense workpaper.

    Parameters
    ----------
    basefile     : str   – path to the input .xlsx file
    client_name  : str   – client / firm name
    fy_start     : str   – FY start date  (DD-MM-YYYY)
    fy_end       : str   – FY end date    (DD-MM-YYYY)
    method       : str   – "top" or "random"
    sample_n     : int   – number of samples to pick
    progress     : callable(float, str) – callback to report progress

    Returns
    -------
    tuple : (total_rows, sample_total, addition_count,
             addition_sum, addition_pct, output_path)
    """

    output_path = basefile.replace(".xlsx", "_output.xlsx")

    fy_start_date = datetime.strptime(fy_start, "%d-%m-%Y")
    fy_end_date   = datetime.strptime(fy_end,   "%d-%m-%Y")
    start_year    = fy_start_date.year
    end_year      = fy_end_date.year

    # ── Read Template ──
    progress(0.10, "Reading template...")
    template_df = pd.read_excel(basefile, sheet_name="Template", header=1)
    template_df = template_df.iloc[1:]
    template_df.columns = (
        template_df.columns.astype(str)
        .str.replace("\n", " ")
        .str.strip()
    )
    if "Expense" in template_df.columns:
        template_df = template_df.rename(columns={"Expense": "Particulars"})
    template_df = template_df.dropna(how="all").reset_index(drop=True)

    # Drop pre-numbered template rows that carry only a serial number ("S.no")
    # and no real data.  These otherwise get written out and leave a long
    # empty tail below the TOTAL row.
    meaningful_cols = [
        c for c in template_df.columns
        if str(c).strip().lower() != "s.no" and not str(c).startswith("Unnamed")
    ]
    if meaningful_cols:
        template_df = (
            template_df.dropna(how="all", subset=meaningful_cols)
            .reset_index(drop=True)
        )

    # ── Load Workbook ──
    progress(0.20, "Loading workbook...")
    wb  = load_workbook(basefile)
    ws  = wb["Prepaid expenses"]
    ws2 = wb["Lead"]

    # ── Header cells ──
    ws["C2"]  = client_name
    ws["C3"]  = f"{fy_start} to {fy_end}"
    ws["Q38"] = fy_start_date
    ws["R38"] = fy_end_date
    ws["Q38"].number_format = "d-mmm-yy"
    ws["R38"].number_format = "d-mmm-yy"
    
    ws2["C2"] = client_name
    ws2["C3"] = f"{fy_start} to {fy_end}"
    
    # Previous FY date and Current FY date for Leads sheet
    ws2["C20"] = fy_end_date 
    ws2["D20"] = fy_end_date - relativedelta(years=1)
    ws2["C20"].number_format = "d-mmm-yy"
    ws2["D20"].number_format = "d-mmm-yy"
    
    def update_dynamic_dates(value):
        if value is None:
            return value
        if not isinstance(value, str):
            return value

        text = value
        start_dm = fy_start_date.strftime("%d %B").lstrip("0")
        end_dm = fy_end_date.strftime("%d %B").lstrip("0")
        start_month = fy_start_date.strftime("%B")
        end_month = fy_end_date.strftime("%B")
        months = r"(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        
        # Replace Full Date + Year
        text = re.sub(rf"0?1(?:st)?[\s\-]+{months}[\s\-]+(20XX|\d{{4}})", f"{start_dm} {start_year}", text, flags=re.IGNORECASE)
        text = re.sub(rf"(?:28|29|30|31)(?:st|nd|rd|th)?[\s\-]+{months}[\s\-]+(20XX|\d{{4}})", f"{end_dm} {end_year}", text, flags=re.IGNORECASE)
        
        # Replace Date + Month (without year)
        text = re.sub(rf"0?1(?:st)?[\s\-]+{months}", start_dm, text, flags=re.IGNORECASE)
        text = re.sub(rf"(?:28|29|30|31)(?:st|nd|rd|th)?[\s\-]+{months}", end_dm, text, flags=re.IGNORECASE)
        
        # Fallback year replacement for standalone 20XX or 202X
        if end_month in text:
            text = text.replace("20XX", str(end_year))
        elif start_month in text:
            text = text.replace("20XX", str(start_year))
        else:
            text = text.replace("20XX", str(end_year))
            
        return text

    # ── Map columns ──
    progress(0.30, "Mapping columns...")
    output_headers = [cell.value for cell in ws[40]]
    output_headers = [
        str(c).replace("\n", " ").strip() if c else None
        for c in output_headers
    ]
    output_col_map = {
        col: idx + 1 for idx, col in enumerate(output_headers) if col
    }
    special_col_mapping = {
        "Document No/ Invoice No.": "Document No",
        "Party Name": "Vendor Name",
    }

    # ── Prepare data ──
    progress(0.40, "Preparing data...")
    start_row = 42

    write_ops = []
    for i, row in template_df.iterrows():
        excel_row = start_row + i
        ssd = pd.to_datetime(row.get("Service start date"), errors="coerce")
        valid_addition = pd.notna(ssd) and ssd >= fy_start_date

        for col_name in template_df.columns:
            if col_name not in output_col_map:
                continue
            col_idx = output_col_map[col_name]

            if col_name == "Addition during the year":
                value = row.get("Amount", None) if valid_addition else None
            elif col_name in special_col_mapping:
                value = row.get(special_col_mapping[col_name])
            else:
                value = row.get(col_name)

            if "date" in str(col_name).lower() and pd.notna(value):
                if isinstance(value, (int, float)):
                    try:
                        value = pd.to_datetime(value, origin="1899-12-30", unit="D")
                    except Exception:
                        pass
                elif isinstance(value, str):
                    try:
                        value = pd.to_datetime(value)
                    except Exception:
                        pass
                if pd.isna(value):
                    value = None

            value = update_dynamic_dates(value)
            template_df.at[i, col_name] = value
            write_ops.append((excel_row, col_idx, value))

    # ── Write data ──
    progress(0.50, "Writing data to Excel...")
    for r, c, v in write_ops:
        cell = ws.cell(row=r, column=c, value=v)
        if isinstance(v, (datetime, pd.Timestamp)):
            cell.number_format = "d-mmm-yy"


    # ── Determine the last real data row dynamically from the
    #    Service Start / End Date columns (TOTAL goes immediately after it) ──
    start_col_idx = next(
        (i for c, i in output_col_map.items() if "service start date" in str(c).lower()), 17
    )
    end_col_idx = next(
        (i for c, i in output_col_map.items() if "service end date" in str(c).lower()), 18
    )

    def is_empty(val):
        """A cell is empty when it has no real value: None, NaN/NaT/NA,
        an empty string, or whitespace only.  Formatting is ignored."""
        if val is None:
            return True
        try:
            if pd.isna(val):          # float NaN, NaT, pandas NA
                return True
        except (TypeError, ValueError):
            pass
        return isinstance(val, str) and val.strip() == ""

    last_start_row = last_end_row = start_row - 1
    for r in range(start_row, ws.max_row + 1):
        if not is_empty(ws.cell(row=r, column=start_col_idx).value):
            last_start_row = max(last_start_row, r)
        if not is_empty(ws.cell(row=r, column=end_col_idx).value):
            last_end_row = max(last_end_row, r)

    last_data_row = max(last_start_row, last_end_row)
    last_row = last_data_row if last_data_row >= start_row else start_row - 1

    print(f"Last Service Start Date Row = {last_start_row}")
    print(f"Last Service End Date Row = {last_end_row}")
    print(f"Last Data Row = {last_row}")

    # ── Formulas ──
    progress(0.60, "Applying formulas...")
    for row in range(start_row, last_row + 1):
        ws[f"N{row}"]  = f"=I{row}-M{row}"
        ws[f"O{row}"]  = f"=R{row}-Q{row}+1"
        ws[f"P{row}"]  = (
            f"=IF(R{row}<$Q$38,0,"
            f"IF(Q{row}<$Q$38,R{row}-$Q$38+1,O{row}))"
        )
        ws[f"S{row}"]  = (
            f"=IF(Q{row}>$Q$38,"
            f"IF(R{row}>$R$38,$R$38-Q{row}+1,R{row}-Q{row}+1),"
            f"IF(R{row}>$R$38,$R$38-$Q$38+1,R{row}-$Q$38+1))"
        )
        ws[f"T{row}"]  = f"=P{row}-S{row}"
        ws[f"U{row}"]  = f"=I{row}*(S{row}/O{row})"
        ws[f"V{row}"]  = f"=ROUND((T{row}/O{row})*I{row},2)"
        ws[f"W{row}"]  = f"=L{row}-U{row}"
        ws[f"X{row}"]  = f"=N{row}-V{row}"
        ws[f"AA{row}"] = f"=IF(T{row}>365,(V{row}*(365/T{row})),V{row})"
        ws[f"AB{row}"] = f"=V{row}-AA{row}"

    # ── Formatting ──
    progress(0.70, "Applying formatting...")
    col_styles = {}
    for col in range(1, ws.max_column + 1):
        src = ws.cell(row=42, column=col)
        if src.has_style:
            col_styles[col] = src._style
    if col_styles:
        for row in range(start_row, last_row + 2):
            for col, style in col_styles.items():
                ws.cell(row=row, column=col)._style = style

    # ── Update dynamic dates across all sheets ──
    progress(0.80, "Updating header dates...")
    for sheet in wb.worksheets:
        for row in sheet.iter_rows():
            if not any(c.value is not None for c in row):
                continue
            for cell in row:
                if isinstance(cell.value, str):
                    val = cell.value
                    cell.value = update_dynamic_dates(val)

    # ── Totals ──
    # TOTAL row sits immediately after the last real data row.
    total_row = last_row + 1
    print(f"Total Row = {total_row}")

    coord27_s = ws.cell(row=start_row, column=27).coordinate
    coord27_e = ws.cell(row=last_row,  column=27).coordinate
    coord28_s = ws.cell(row=start_row, column=28).coordinate
    coord28_e = ws.cell(row=last_row,  column=28).coordinate

    # Wipe any leftover template formulas/values sitting on the TOTAL row so it
    # shows only the "Total" label and the Current / Non-Current sums.
    for col in range(1, ws.max_column + 1):
        ws.cell(row=total_row, column=col).value = None

    ws["E32"] = f"=SUM({coord27_s}:{coord27_e})"
    ws["E33"] = f"=SUM({coord28_s}:{coord28_e})"
    ws.cell(row=total_row, column=1).value = "Total"           # column A label
    ws.cell(row=total_row, column=27).value = f"=SUM({coord27_s}:{coord27_e})"
    ws.cell(row=total_row, column=28).value = f"=SUM({coord28_s}:{coord28_e})"

    # ── Sampling ──
    progress(0.90, "Performing sampling...")
    data  = ws.iter_rows(min_row=start_row, max_row=last_row, values_only=True)
    df_ws = pd.DataFrame(data, columns=output_headers)

    df_ws["_sampling_amount_"] = None
    ssd_col = pd.to_datetime(df_ws["Service start date"], errors="coerce")
    mask = ssd_col >= fy_start_date
    df_ws.loc[mask, "_sampling_amount_"] = df_ws.loc[mask, "Amount"]

    if method == "top":
        sampled_indices = (
            df_ws[df_ws["_sampling_amount_"].notna()]
            .sort_values(by="_sampling_amount_", ascending=False)
            .head(sample_n)
            .index
        )
    elif method == "random":
        eligible_df = df_ws[df_ws["_sampling_amount_"].notna()]
        if sample_n > len(eligible_df):
            raise ValueError("Sample size exceeds eligible records")
        sampled_indices = eligible_df.sample(n=sample_n, random_state=42).index
    else:
        sampled_indices = []

    doc_col_idx   = output_col_map.get("Document No/ Invoice No.")
    party_col_idx = output_col_map.get("Party Name")
    doc_src_idx   = (df_ws.columns.get_loc("Document No")
                     if "Document No" in df_ws.columns else None)
    party_src_idx = (df_ws.columns.get_loc("Vendor Name")
                     if "Vendor Name" in df_ws.columns else None)

    # Clear all first
    for i in range(len(df_ws)):
        er = start_row + i
        if doc_col_idx:
            ws.cell(row=er, column=doc_col_idx, value=None)
        if party_col_idx:
            ws.cell(row=er, column=party_col_idx, value=None)

    # Write only sampled rows
    for i in sampled_indices:
        er = start_row + i
        if doc_col_idx and doc_src_idx is not None:
            ws.cell(row=er, column=doc_col_idx,
                    value=df_ws.iloc[i, doc_src_idx])
        if party_col_idx and party_src_idx is not None:
            ws.cell(row=er, column=party_col_idx,
                    value=df_ws.iloc[i, party_src_idx])

    # Statistics
    sample_total = df_ws.loc[sampled_indices, "_sampling_amount_"].sum()
    add_series = pd.to_numeric(
        template_df["Addition during the year"].dropna(), errors="coerce"
    ).dropna()
    addition_sum   = add_series.sum()
    addition_count = len(add_series)
    addition_pct   = (sample_total / addition_sum * 100) if addition_sum else 0

    # Total-row formatting: bold text, NO grey background fill.
    bottom_row = total_row
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=bottom_row, column=col)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type=None)   # remove the grey background
        if col >= 2:
            cell.number_format = "#,##0"

    # ── Force the Service Start / End Date columns to a real date format
    #    (match the Template's "d-mmm-yy") so they don't show as serial numbers ──
    DATE_FMT = "d-mmm-yy"
    for r in range(start_row, last_row + 1):
        ws.cell(row=r, column=start_col_idx).number_format = DATE_FMT
        ws.cell(row=r, column=end_col_idx).number_format = DATE_FMT

    # ── End the sheet right after the TOTAL row: remove every trailing
    #    template row below it so nothing prints after the totals ──
    if ws.max_row > total_row:
        ws.delete_rows(total_row + 1, ws.max_row - total_row)

    # Remove Template sheet & unhide all
    if "Template" in wb.sheetnames:
        del wb["Template"]
    for sheet in wb.worksheets:
        sheet.sheet_state = "visible"

    # Save
    progress(0.95, "Saving output file...")
    wb.save(output_path)

    return (
        len(template_df),
        sample_total,
        addition_count,
        addition_sum,
        addition_pct,
        output_path,
    )


# ──────────────────────────────────────────────────────────
#  Entry Point
# ──────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = PrepaidApp(process_fn=process_workpaper)
    app.mainloop()