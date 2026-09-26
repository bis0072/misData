import io
from pathlib import Path
from typing import Optional
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def decode_html_bytes(data: bytes) -> str:
    """Decode HTML bytes to string safely trying UTF-8 first, falling back to Latin-1."""
    try:
        return data.decode('utf-8')
    except UnicodeDecodeError:
        return data.decode('latin-1', errors='replace')


def parse_table_from_html(html_text: str, table_index: int = 4) -> pd.DataFrame:
    """Return a DataFrame from the nth <table> in an HTML string."""
    soup = BeautifulSoup(html_text, 'html.parser')
    tables = soup.find_all('table')
    if len(tables) <= table_index:
        raise ValueError(
            f"Table index {table_index} not found in HTML. Available tables: {len(tables)}"
        )
    table = tables[table_index]
    rows = table.find_all('tr')
    if not rows:
        raise ValueError(f"Table at index {table_index} has no rows.")

    headers = [cell.get_text(strip=True) for cell in rows[0].find_all(['th', 'td'])]
    data = []
    for row in rows[1:]:
        cells = [cell.get_text(strip=True) for cell in row.find_all(['th', 'td'])]
        if cells and any(c != '' for c in cells):
            data.append(cells)

    # Pad / trim rows to header length
    col_count = len(headers)
    padded = [r[:col_count] + [''] * max(0, col_count - len(r)) for r in data]
    return pd.DataFrame(padded, columns=headers)


def get_page_title_from_html(html_text: str, fallback: str = "Report", tag: str = 'p') -> str:
    """Extract report title from the first <p> or <title> tag, skipping loading messages."""
    soup = BeautifulSoup(html_text, 'html.parser')
    for p in soup.find_all(tag):
        text = p.get_text(' ', strip=True)
        if len(text) > 10 and 'loading' not in text.lower():
            return text[:120]
    title_tag = soup.find('title')
    if title_tag and title_tag.get_text(strip=True):
        t = title_tag.get_text(' ', strip=True)
        if 'loading' not in t.lower():
            return t[:120]
    return fallback


def to_num(series: pd.Series) -> pd.Series:
    """Convert comma-formatted string numbers to numeric floats/ints."""
    return pd.to_numeric(series.astype(str).str.replace(',', ''), errors='coerce').fillna(0)


def sanitize_sheet_name(name: str, max_length: int = 31) -> str:
    """Clean sheet names according to Excel specifications (<= 31 chars, no :\\/?*[])."""
    for ch in [':', '\\', '/', '?', '*', '[', ']']:
        name = name.replace(ch, '_')
    name = name.strip("' ").strip()
    return name[:max_length] if name else "Sheet"


def find_column(df: pd.DataFrame, target: str, alternatives: list[str]) -> Optional[str]:
    """Find a column by exact or case-insensitive/partial match."""
    if target in df.columns:
        return target
    lower_map = {c.strip().lower(): c for c in df.columns}
    for alt in [target] + alternatives:
        if alt.strip().lower() in lower_map:
            return lower_map[alt.strip().lower()]
    for alt in [target] + alternatives:
        alt_l = alt.strip().lower()
        for col_l, original_col in lower_map.items():
            if alt_l in col_l:
                return original_col
    return None


def style_sheet(ws, header_row: int = 1) -> None:
    """Apply professional styling to an openpyxl worksheet."""
    hdr_fill = PatternFill("solid", fgColor="1F4E79")
    hdr_font = Font(color="FFFFFF", bold=True, size=10)
    alt_fill = PatternFill("solid", fgColor="D6E4F0")
    lkp_fill = PatternFill("solid", fgColor="E8F5E9")  # light green for looked-up columns
    lkp_font = Font(color="1A5276", bold=False, size=10)
    thin = Side(style='thin', color='B0BEC5')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for row_idx, row in enumerate(ws.iter_rows(), start=1):
        if row_idx < header_row:
            continue
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False)
            if row_idx == header_row:
                cell.fill = hdr_fill
                cell.font = hdr_font
            elif row_idx > header_row:
                col_header = str(ws.cell(row=header_row, column=cell.column).value or '').lower()
                if any(k in col_header for k in ['total pol', 'total prem', 'prem (', 'pol (', 'nop (', 'non sing prem']):
                    cell.fill = lkp_fill
                    cell.font = lkp_font
                elif (row_idx - header_row) % 2 == 0:
                    cell.fill = alt_fill

    # Auto-adjust column widths based on table content
    for col in ws.columns:
        max_len = max((len(str(cell.value or '')) for cell in col if cell.row >= header_row), default=8)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 4, 10), 40)

    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)


def generate_excel_in_memory(
    html1_bytes: bytes,
    html2_bytes: Optional[bytes] = None,
    filename1: str = "Ranking List",
    filename2: str = "Period Stats",
    table_index: int = 4,
    comparison_files: Optional[list[tuple[bytes, str]]] = None,
    nop_col_index: Optional[int] = None,
    prem_col_index: Optional[int] = None,
    index_base: int = 0,
    report_heading: Optional[str] = "RESULTS FOR THE DATA"
) -> tuple[io.BytesIO, dict]:
    """
    Parses Base HTML bytes and dynamic comparison HTML files entirely in RAM.
    Performs VLOOKUPs matching by 'Agent' code using specified or auto-detected column indices:
      - nop_col_index: Column index for Total POL / NOP (e.g. 7 in 0-based or 8 in 1-based)
      - prem_col_index: Column index for Total Premium (e.g. 10 or 9 in 0-based or 11 in 1-based)
      - index_base: 0 for Python 0-indexed (default), 1 for Excel 1-indexed
    Zero disk storage is used.
    """
    # Convert 1-based (Excel) to 0-based (Python) if index_base == 1
    eff_nop_idx = (nop_col_index - 1) if (index_base == 1 and nop_col_index is not None) else nop_col_index
    eff_prem_idx = (prem_col_index - 1) if (index_base == 1 and prem_col_index is not None) else prem_col_index

    # Build list of comparison files
    comp_list: list[tuple[bytes, str]] = []
    if comparison_files:
        comp_list.extend(comparison_files)
    elif html2_bytes:
        comp_list.append((html2_bytes, filename2))

    if not comp_list:
        raise ValueError("At least one comparison HTML file must be provided.")

    # Parse Base File (html1)
    html1_text = decode_html_bytes(html1_bytes)
    df1 = parse_table_from_html(html1_text, table_index=table_index)

    agent_col1 = find_column(df1, 'Agent', ['agent', 'agent code', 'agent_code'])
    if not agent_col1:
        raise ValueError(f"Base file '{filename1}' missing required column 'Agent'")
    if agent_col1 != 'Agent':
        df1['Agent'] = df1[agent_col1]

    df1['Agent'] = df1['Agent'].astype(str).str.strip()
    title1 = get_page_title_from_html(html1_text, fallback=filename1)

    df_merged = df1.copy()
    raw_comp_sheets: list[tuple[str, pd.DataFrame]] = []
    comparison_meta: list[dict] = []
    seen_stems: dict[str, int] = {}
    all_pol_cols: list[str] = []

    # Process each comparison file
    for idx, (comp_bytes, comp_filename) in enumerate(comp_list):
        stem = Path(comp_filename).stem
        if not stem:
            stem = f"file{idx + 2}"

        count = seen_stems.get(stem, 0)
        seen_stems[stem] = count + 1
        label = stem if count == 0 else f"{stem}_{count}"

        comp_text = decode_html_bytes(comp_bytes)
        df_comp = parse_table_from_html(comp_text, table_index=table_index)

        # 1. Resolve Agent column
        comp_agent = find_column(df_comp, 'Agent', ['agent', 'agent code'])
        if not comp_agent:
            # Fallback to column index 3 if available
            if len(df_comp.columns) > 3 and 'agent' in str(df_comp.columns[3]).lower():
                comp_agent = df_comp.columns[3]
            else:
                raise ValueError(f"Comparison file '{comp_filename}' missing required column 'Agent'")
        df_comp[comp_agent] = df_comp[comp_agent].astype(str).str.strip()

        # 2. Resolve Total POL / NOP column (by user-specified index or header name match)
        comp_pol = None
        if eff_nop_idx is not None and 0 <= eff_nop_idx < len(df_comp.columns):
            comp_pol = df_comp.columns[eff_nop_idx]
        else:
            comp_pol = find_column(df_comp, 'Total POL', ['tot pol', 'total policies', 'total_pol', 'tot policies', 'nop', 'tot no'])

        # 3. Resolve Total Premium column (by user-specified index or header name match)
        comp_prem = None
        if eff_prem_idx is not None and 0 <= eff_prem_idx < len(df_comp.columns):
            comp_prem = df_comp.columns[eff_prem_idx]
        else:
            comp_prem = find_column(df_comp, 'TOTAL PREM(SP+NSP)', ['total prem', 'total premium', 'tot prem', 'prem(sp+nsp)', 'non sing prem.'])

        if not comp_pol or not comp_prem:
            available_cols = [f"[Col {i+1} | Idx {i}]: {col}" for i, col in enumerate(df_comp.columns)]
            missing = []
            if not comp_pol:
                missing.append(f"Total POL / NOP (idx {eff_nop_idx})")
            if not comp_prem:
                missing.append(f"Total Premium (idx {eff_prem_idx})")
            raise ValueError(
                f"Comparison file '{comp_filename}' could not resolve column(s): {', '.join(missing)}. "
                f"Available table columns: {available_cols}"
            )

        # Column names for merged sheet
        if eff_nop_idx is not None and 0 <= eff_nop_idx < len(df_comp.columns):
            col_pol_label = comp_pol.strip()
        else:
            col_pol_label = "Total POL"

        if eff_prem_idx is not None and 0 <= eff_prem_idx < len(df_comp.columns):
            col_prem_label = comp_prem.strip()
        else:
            col_prem_label = "Total Prem"

        col_pol_name = f"{col_pol_label} ({label})"
        col_prem_name = f"{col_prem_label} ({label})"

        all_pol_cols.append(col_pol_name)

        # Lookup table
        lookup_df = df_comp[[comp_agent, comp_pol, comp_prem]].copy()
        lookup_df.columns = ['Agent', col_pol_name, col_prem_name]
        lookup_df = lookup_df.drop_duplicates(subset='Agent')

        # Merge (VLOOKUP)
        df_merged = df_merged.merge(lookup_df, on='Agent', how='left')
        df_merged[col_pol_name] = df_merged[col_pol_name].fillna('0')
        df_merged[col_prem_name] = df_merged[col_prem_name].fillna('0')

        comp_title = get_page_title_from_html(comp_text, fallback=comp_filename)

        # Sheet name
        if len(comp_list) == 1 and 'html2' in comp_filename.lower():
            sheet_name = 'HTML2 - Period Stats'
        else:
            sheet_name = sanitize_sheet_name(f"{label} - Raw")

        raw_comp_sheets.append((sheet_name, df_comp))

        # Metrics for this file
        matched_this = int(df_merged[col_pol_name].ne('0').sum())
        unmatched_this = int(df_merged[col_pol_name].eq('0').sum())
        pol_sum = float(to_num(df_merged[col_pol_name]).sum())
        prem_sum = float(to_num(df_merged[col_prem_name]).sum())

        comparison_meta.append({
            "filename": comp_filename,
            "label": label,
            "title": comp_title,
            "pol_col": comp_pol,
            "prem_col": comp_prem,
            "matched": matched_this,
            "unmatched": unmatched_this,
            "total_policies": pol_sum,
            "total_premium": prem_sum
        })

    # Overall matched across files
    matched_any_mask = pd.Series(False, index=df_merged.index)
    for pol_col in all_pol_cols:
        matched_any_mask |= df_merged[pol_col].ne('0')

    overall_matched = int(matched_any_mask.sum())
    overall_unmatched = int((~matched_any_mask).sum())

    tot_no_num = to_num(df_merged['Tot No']) if 'Tot No' in df_merged.columns else pd.Series(0, index=df_merged.index)
    tot_prem1_num = to_num(df_merged['Total Prem']) if 'Total Prem' in df_merged.columns else pd.Series(0, index=df_merged.index)

    # In-memory Excel Generation: ZERO disk footprint
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        # Sheet 1: Base File Raw
        base_sheet_name = 'HTML1 - Ranking'
        df1.to_excel(writer, sheet_name=base_sheet_name, index=False)
        ws1 = writer.book[base_sheet_name]
        style_sheet(ws1)

        # Sheet 2..N: Comparison Files Raw
        for sheet_name, df_c in raw_comp_sheets:
            df_c.to_excel(writer, sheet_name=sheet_name, index=False)
            ws_c = writer.book[sheet_name]
            style_sheet(ws_c)

        # Unified Merged Sheet (VLOOKUP)
        export_cols = [c for c in df_merged.columns if not c.startswith('_')]
        df_export = df_merged[export_cols].copy()
        max_col = len(export_cols)

        # Structure of Merged Sheet:
        # Row 1: Big stylish heading banner merged dynamically from Col 1 to max_col
        # Row 2: Source 1 (Base file) metadata banner merged across all cols
        # Rows 3 .. 2 + num_comp - 1: Comparison file(s) metadata banner(s)
        # Openpyxl header_row = num_comp + 3
        # Pandas startrow = num_comp + 2 (so headers land on row num_comp + 3)

        num_comp = len(comp_list)
        startrow = num_comp + 2
        header_row = startrow + 1

        df_export.to_excel(writer, sheet_name='Merged (VLOOKUP)', index=False, startrow=startrow)
        ws3 = writer.book['Merged (VLOOKUP)']

        # Style table headers and data rows
        style_sheet(ws3, header_row=header_row)

        # 1. Big Stylish Heading Banner on Row 1 (Merged dynamically from Col 1 to max_col)
        heading_text = (report_heading or "RESULTS FOR THE DATA").strip()
        ws3.cell(row=1, column=1, value=heading_text)
        ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
        ws3.row_dimensions[1].height = 42

        title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
        title_fill = PatternFill("solid", fgColor="1F4E79")  # Executive deep navy
        title_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        title_border = Border(
            left=Side(style='thin', color='B0BEC5'),
            right=Side(style='thin', color='B0BEC5'),
            top=Side(style='thin', color='B0BEC5'),
            bottom=Side(style='medium', color='0F294A')
        )

        for c_idx in range(1, max_col + 1):
            cell = ws3.cell(row=1, column=c_idx)
            cell.fill = title_fill
            cell.font = title_font
            cell.alignment = title_align
            cell.border = title_border

        # 2. Source Metadata Rows (Merged across all columns for a clean corporate subtitle bar)
        meta_fill = PatternFill("solid", fgColor="EBF5FB")
        meta_font = Font(name="Calibri", size=9, bold=True, color="1F4E79")
        meta_align = Alignment(horizontal="left", vertical="center")
        meta_border = Border(
            left=Side(style='thin', color='CFD8DC'),
            right=Side(style='thin', color='CFD8DC'),
            top=Side(style='thin', color='CFD8DC'),
            bottom=Side(style='thin', color='CFD8DC')
        )

        # Row 2: Base File Info
        ws3.cell(row=2, column=1, value=f"Source 1 (Base): {title1} [{filename1}]")
        ws3.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
        ws3.row_dimensions[2].height = 20
        for c_idx in range(1, max_col + 1):
            cell = ws3.cell(row=2, column=c_idx)
            cell.fill = meta_fill
            cell.font = meta_font
            cell.alignment = meta_align
            cell.border = meta_border

        # Rows 3..: Comparison Files Info
        for i, meta in enumerate(comparison_meta):
            r = i + 3
            ws3.cell(row=r, column=1, value=f"Source {i + 2} (Compare): {meta['title']} [{meta['filename']}] (Cols: {meta['pol_col']}, {meta['prem_col']})")
            ws3.merge_cells(start_row=r, start_column=1, end_row=r, end_column=max_col)
            ws3.row_dimensions[r].height = 20
            for c_idx in range(1, max_col + 1):
                cell = ws3.cell(row=r, column=c_idx)
                cell.fill = meta_fill
                cell.font = meta_font
                cell.alignment = meta_align
                cell.border = meta_border

        # Summary Sheet
        wb = writer.book
        if len(comp_list) == 1:
            meta = comparison_meta[0]
            summary_df = pd.DataFrame({
                'Metric': [
                    f'Total Agents ({filename1})',
                    f'Matched with {meta["filename"]}',
                    'Not matched',
                    '',
                    f'Total Policies ({filename1} Tot No)',
                    f'Total Policies ({meta["filename"]} {meta["pol_col"]} - matched)',
                    '',
                    f'Total Premium {filename1}',
                    f'Total Premium {meta["filename"]} (matched)',
                ],
                'Value': [
                    len(df1),
                    meta["matched"],
                    meta["unmatched"],
                    '',
                    float(tot_no_num.sum()),
                    meta["total_policies"],
                    '',
                    float(tot_prem1_num.sum()),
                    meta["total_premium"],
                ]
            })
        else:
            metrics = [
                f'Total Agents in Base File ({filename1})',
                'Matched in at least one comparison file',
                'Not matched in any comparison file',
                '',
                f'Total Policies in Base File ({filename1})',
                f'Total Premium in Base File ({filename1})',
                '',
                '--- Breakdown Per Comparison File ---'
            ]
            values = [
                len(df1),
                overall_matched,
                overall_unmatched,
                '',
                float(tot_no_num.sum()),
                float(tot_prem1_num.sum()),
                '',
                ''
            ]

            for meta in comparison_meta:
                metrics.extend([
                    f'Matched with {meta["filename"]}',
                    f'Total POL matched ({meta["filename"]} [{meta["pol_col"]}])',
                    f'Total Prem matched ({meta["filename"]} [{meta["prem_col"]}])',
                    ''
                ])
                values.extend([
                    meta["matched"],
                    meta["total_policies"],
                    meta["total_premium"],
                    ''
                ])

            summary_df = pd.DataFrame({'Metric': metrics, 'Value': values})

        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        ws4 = wb['Summary']
        style_sheet(ws4)

    buffer.seek(0)

    stats = {
        "title1": title1,
        "filename1": filename1,
        "total_agents": len(df1),
        "matched": overall_matched,
        "unmatched": overall_unmatched,
        "total_policies_html1": float(tot_no_num.sum()),
        "total_premium_html1": float(tot_prem1_num.sum()),
        "comparison_count": len(comp_list),
        "files": comparison_meta,
        "filesize_bytes": buffer.getbuffer().nbytes
    }

    if comparison_meta:
        stats["title2"] = comparison_meta[0]["title"]
        stats["total_policies_html2"] = comparison_meta[0]["total_policies"]
        stats["total_premium_html2"] = comparison_meta[0]["total_premium"]

    return buffer, stats
