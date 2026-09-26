import io
import openpyxl
from fastapi.testclient import TestClient
from main import app

client = TestClient(app)

def test_health():
    res = client.get("/health")
    assert res.status_code == 200
    assert res.json()["status"] == "healthy"
    print("Health check passed.")

def test_index_page():
    res = client.get("/")
    assert res.status_code == 200
    assert "MIS HTML to Excel Studio" in res.text
    print("Web UI page passed.")

def test_summary_endpoint():
    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        res = client.post(
            "/api/summary",
            files={
                "file1": ("html1.htm", f1, "text/html"),
                "file2": ("html2.htm", f2, "text/html")
            },
            data={"table_index": 4}
        )
    assert res.status_code == 200
    data = res.json()
    assert data["status"] == "success"
    assert data["summary"]["total_agents"] == 126
    assert data["summary"]["matched"] == 30
    assert data["summary"]["unmatched"] == 96
    print("Summary API endpoint passed:", data["summary"])

def test_convert_endpoint():
    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        res = client.post(
            "/api/convert",
            files={
                "file1": ("html1.htm", f1, "text/html"),
                "file2": ("html2.htm", f2, "text/html")
            },
            data={"table_index": 4}
        )
    assert res.status_code == 200
    assert res.headers["content-type"] == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    assert res.headers["x-stats-matched"] == "30"
    assert res.headers["x-stats-unmatched"] == "96"
    assert res.headers["x-stats-total-agents"] == "126"

    # Verify that the returned bytes are a valid Excel workbook in memory
    content = res.content
    assert len(content) > 1000
    wb = openpyxl.load_workbook(io.BytesIO(content))
    sheets = wb.sheetnames
    assert "HTML1 - Ranking" in sheets
    assert "HTML2 - Period Stats" in sheets
    assert "Merged (VLOOKUP)" in sheets
    assert "Summary" in sheets
    print("Convert endpoint passed! Generated valid Excel workbook with sheets:", sheets)

def test_dynamic_multi_files():
    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        f1_data = f1.read()
        f2_data = f2.read()

    res = client.post(
        "/api/convert",
        files=[
            ("file1", ("baseReport.html", f1_data, "text/html")),
            ("comparison_files", ("temadata.htm", f2_data, "text/html")),
            ("comparison_files", ("dvnDor.html", f2_data, "text/html")),
        ],
        data={"table_index": 4}
    )
    assert res.status_code == 200
    assert res.headers["x-stats-matched"] == "30"
    assert res.headers["x-stats-files-count"] == "2"

    wb = openpyxl.load_workbook(io.BytesIO(res.content))
    sheets = wb.sheetnames
    assert "HTML1 - Ranking" in sheets
    assert "temadata - Raw" in sheets
    assert "dvnDor - Raw" in sheets
    assert "Merged (VLOOKUP)" in sheets
    assert "Summary" in sheets

    ws = wb["Merged (VLOOKUP)"]
    headers = [str(cell.value) for cell in ws[5]]
    assert "Total POL (temadata)" in headers
    assert "Total Prem (temadata)" in headers
    assert "Total POL (dvnDor)" in headers
    assert "Total Prem (dvnDor)" in headers
    print("Dynamic multi-file test passed with custom filenames:", sheets)

def test_custom_column_indices():
    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        f1_data = f1.read()
        f2_data = f2.read()

    res = client.post(
        "/api/convert",
        files={
            "file1": ("html1.htm", f1_data, "text/html"),
            "file2": ("html2.htm", f2_data, "text/html")
        },
        data={"table_index": 4, "nop_col_index": 7, "prem_col_index": 9}
    )
    assert res.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(res.content))
    ws = wb["Merged (VLOOKUP)"]
    headers = [str(cell.value) for cell in ws[4]]
    assert "Total POL (html2)" in headers
    assert "NON SING PREM. (html2)" in headers
    print("Custom column indices test (Python 0-based 7 & 9) passed:", headers)

def test_excel_1based_indices():
    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        f1_data = f1.read()
        f2_data = f2.read()

    res = client.post(
        "/api/convert",
        files={
            "file1": ("html1.htm", f1_data, "text/html"),
            "file2": ("html2.htm", f2_data, "text/html")
        },
        data={"table_index": 4, "nop_col_index": 8, "prem_col_index": 11, "index_base": 1}
    )
    assert res.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(res.content))
    ws = wb["Merged (VLOOKUP)"]
    headers = [str(cell.value) for cell in ws[4]]
    assert "Total POL (html2)" in headers
    assert "TOTAL PREM(SP+NSP) (html2)" in headers
    print("Excel 1-based indices test (Col 8 & 11) passed:", headers)

def test_dynamic_output_filename_with_date():
    from datetime import datetime
    today_str = datetime.now().strftime("%Y-%m-%d")

    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        f1_data = f1.read()
        f2_data = f2.read()

    # 1. Default filename: report_with_lookup_YYYY-MM-DD.xlsx
    res_default = client.post(
        "/api/convert",
        files={
            "file1": ("html1.htm", f1_data, "text/html"),
            "file2": ("html2.htm", f2_data, "text/html")
        },
        data={"table_index": 4}
    )
    assert res_default.status_code == 200
    expected_default_name = f"report_with_lookup_{today_str}.xlsx"
    assert res_default.headers["x-output-filename"] == expected_default_name
    assert f'filename="{expected_default_name}"' in res_default.headers["content-disposition"]

    # 2. Custom prefix filename: branch_mis_summary_YYYY-MM-DD.xlsx
    res_custom = client.post(
        "/api/convert",
        files={
            "file1": ("html1.htm", f1_data, "text/html"),
            "file2": ("html2.htm", f2_data, "text/html")
        },
        data={"table_index": 4, "output_filename": "branch_mis_summary"}
    )
    assert res_custom.status_code == 200
    expected_custom_name = f"branch_mis_summary_{today_str}.xlsx"
    assert res_custom.headers["x-output-filename"] == expected_custom_name
    assert f'filename="{expected_custom_name}"' in res_custom.headers["content-disposition"]
    print(f"Dynamic date-stamped output filename test passed! Generated: '{expected_default_name}' and '{expected_custom_name}'")

def test_dynamic_merged_sheet_heading():
    with open("html1.htm", "rb") as f1, open("html2.htm", "rb") as f2:
        f1_data = f1.read()
        f2_data = f2.read()

    # Test with custom heading matching user prompt: "RELSUTS FOR THE DATA"
    custom_heading = "RELSUTS FOR THE DATA"
    res = client.post(
        "/api/convert",
        files={
            "file1": ("html1.htm", f1_data, "text/html"),
            "file2": ("html2.htm", f2_data, "text/html")
        },
        data={"table_index": 4, "report_heading": custom_heading}
    )
    assert res.status_code == 200
    wb = openpyxl.load_workbook(io.BytesIO(res.content))
    ws = wb["Merged (VLOOKUP)"]

    # 1. Heading text and location
    assert ws["A1"].value == custom_heading

    # 2. Styling: Big size font (16pt bold white)
    cell_a1 = ws["A1"]
    assert cell_a1.font.size == 16
    assert cell_a1.font.bold is True
    assert cell_a1.font.color.rgb == "00FFFFFF" or cell_a1.font.color.rgb == "FFFFFF"

    # 3. Styling: Executive navy fill and row height
    assert cell_a1.fill.fgColor.rgb == "001F4E79" or cell_a1.fill.fgColor.rgb == "1F4E79"
    assert ws.row_dimensions[1].height == 42

    # 4. Merged dynamically across all data columns
    # There are 16 columns in df_merged, so merged range should start at A1 and end at P1
    merged_ranges = [str(r) for r in ws.merged_cells.ranges]
    assert any(r.startswith("A1:") for r in merged_ranges), f"Expected A1 merged range, got: {merged_ranges}"
    row1_merge = [r for r in merged_ranges if r.startswith("A1:")][0]
    print(f"Dynamic stylish heading test passed! Range: {row1_merge}, Heading: '{custom_heading}', Height: 42, Font: 16pt Bold")

if __name__ == "__main__":
    test_health()
    test_index_page()
    test_summary_endpoint()
    test_convert_endpoint()
    test_dynamic_multi_files()
    test_custom_column_indices()
    test_excel_1based_indices()
    test_dynamic_output_filename_with_date()
    test_dynamic_merged_sheet_heading()
    print("\nALL IN-MEMORY, DYNAMIC MULTI-FILE, COLUMN INDEX, DATE FILENAME, AND STYLISH MERGED HEADING TESTS PASSED PERFECTLY!")


