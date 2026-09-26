from typing import Optional
from fastapi import APIRouter, File, UploadFile, Form, HTTPException, status
from fastapi.responses import StreamingResponse
from converter import generate_excel_in_memory

router = APIRouter()


async def read_uploaded_file(file: UploadFile) -> bytes:
    """Read and validate that an uploaded file is not empty."""
    if not file.filename:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="File must be provided with a valid filename."
        )
    content = await file.read()
    if len(content) == 0:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"File '{file.filename}' is empty."
        )
    return content


async def extract_comparison_files(
    file2: Optional[UploadFile],
    comparison_files: list[UploadFile]
) -> list[tuple[bytes, str]]:
    """Gather and read bytes from all provided comparison files (supports single or dynamic multi-file)."""
    candidates: list[UploadFile] = []
    if file2 and file2.filename:
        candidates.append(file2)
    for f in comparison_files:
        if f and f.filename:
            candidates.append(f)

    if not candidates:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="At least one comparison HTML file must be uploaded."
        )

    collected: list[tuple[bytes, str]] = []
    for f in candidates:
        content = await read_uploaded_file(f)
        collected.append((content, f.filename))
    return collected


@router.get("/health", tags=["Health"])
def health_check():
    """Health check endpoint for container and cloud platform monitoring."""
    return {
        "status": "healthy",
        "service": "mis-html-to-excel",
        "storage": "100% in-memory / zero-disk"
    }


@router.post("/api/convert", tags=["Converter"])
async def convert_html_files(
    file1: UploadFile = File(..., description="Base HTML file (e.g. Ranking list)"),
    file2: Optional[UploadFile] = File(None, description="Single comparison file (for backward compatibility)"),
    comparison_files: list[UploadFile] = File(default=[], description="Dynamic comparison HTML files"),
    table_index: int = Form(4, description="Table position in HTML (default 4)"),
    nop_col_index: Optional[int] = Form(None, description="Column position for Total POL/NOP in lookup files"),
    prem_col_index: Optional[int] = Form(None, description="Column position for Total Prem in lookup files"),
    index_base: int = Form(0, description="0 for 0-indexed (Python), 1 for 1-indexed (Excel)"),
    output_filename: Optional[str] = Form(None, description="Optional custom prefix for the output filename"),
    report_heading: Optional[str] = Form("RESULTS FOR THE DATA", description="Heading banner text for the Merged sheet")
):
    """
    Accepts a constant base HTML file and one or more comparison HTML files (any filename).
    Allows configuring table index, custom lookup column indices, dynamic date-stamped filename,
    and a custom stylish heading banner merged dynamically across all columns in RAM.
    """
    content1 = await read_uploaded_file(file1)
    comp_files = await extract_comparison_files(file2, comparison_files)

    try:
        excel_buffer, stats = generate_excel_in_memory(
            html1_bytes=content1,
            filename1=file1.filename,
            comparison_files=comp_files,
            table_index=table_index,
            nop_col_index=nop_col_index,
            prem_col_index=prem_col_index,
            index_base=index_base,
            report_heading=report_heading
        )

        from datetime import datetime
        date_str = datetime.now().strftime("%Y-%m-%d")

        if output_filename and output_filename.strip():
            clean_name = output_filename.strip()
            if clean_name.lower().endswith(".xlsx"):
                clean_name = clean_name[:-5]
            final_filename = f"{clean_name}_{date_str}.xlsx" if date_str not in clean_name else f"{clean_name}.xlsx"
        else:
            final_filename = f"report_with_lookup_{date_str}.xlsx"

        headers = {
            "Content-Disposition": f'attachment; filename="{final_filename}"',
            "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "X-Stats-Matched": str(stats["matched"]),
            "X-Stats-Unmatched": str(stats["unmatched"]),
            "X-Stats-Total-Agents": str(stats["total_agents"]),
            "X-Stats-Files-Count": str(stats["comparison_count"]),
            "X-Output-Filename": final_filename,
            "Access-Control-Expose-Headers": (
                "Content-Disposition, X-Stats-Matched, X-Stats-Unmatched, "
                "X-Stats-Total-Agents, X-Stats-Files-Count, X-Output-Filename"
            )
        }

        return StreamingResponse(
            excel_buffer,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers
        )

    except ValueError as ve:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(ve))
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=f"Processing error: {str(exc)}")


@router.post("/api/summary", tags=["Converter"])
async def preview_summary(
    file1: UploadFile = File(..., description="Base HTML file (e.g. Ranking list)"),
    file2: Optional[UploadFile] = File(None, description="Single comparison file (for backward compatibility)"),
    comparison_files: list[UploadFile] = File(default=[], description="Dynamic comparison HTML files"),
    table_index: int = Form(4, description="Table position in HTML (default 4)"),
    nop_col_index: Optional[int] = Form(None, description="Column position for Total POL/NOP in lookup files"),
    prem_col_index: Optional[int] = Form(None, description="Column position for Total Prem in lookup files"),
    index_base: int = Form(0, description="0 for 0-indexed (Python), 1 for 1-indexed (Excel)"),
    report_heading: Optional[str] = Form("RESULTS FOR THE DATA", description="Heading banner text for the Merged sheet")
):
    """
    Returns processed summary metrics for all uploaded files with customizable column indices without downloading full Excel workbook.
    """
    content1 = await read_uploaded_file(file1)
    comp_files = await extract_comparison_files(file2, comparison_files)

    try:
        _, stats = generate_excel_in_memory(
            html1_bytes=content1,
            filename1=file1.filename,
            comparison_files=comp_files,
            table_index=table_index,
            nop_col_index=nop_col_index,
            prem_col_index=prem_col_index,
            index_base=index_base,
            report_heading=report_heading
        )
        return {"status": "success", "summary": stats}
    except ValueError as ve:
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=str(ve))
    except Exception as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=f"Processing error: {str(exc)}")

