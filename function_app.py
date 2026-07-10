"""
Azure Function: PDF (Uptick/Anglicare SoW) -> Excel

Three HTTP-triggered endpoints:

1. POST /api/pdf-to-excel
   Body:    raw PDF bytes (Content-Type: application/pdf)
   Headers: X-File-Name (optional)
   Returns: raw .xlsx bytes, unpriced (LineItems table with blank
            Unit Price / Total Price columns). One-shot, no pricing lookup.

2. POST /api/extract-items
   Body:    raw PDF bytes (Content-Type: application/pdf)
   Headers: X-File-Name (optional)
   Returns: JSON array of extracted line items (Unit Price / Total Price
            keys present but blank - for Power Automate to fill in).

3. POST /api/build-excel
   Body:    JSON array of items (same shape as extract-items returns, with
            Unit Price / Total Price now populated)
   Headers: X-File-Name (optional)
   Returns: raw .xlsx bytes, containing a real Excel Table named "LineItems".

Typical Power Automate flow using #2 + #3 (no temp file needed):
    HTTP (extract-items) -> Apply to each item -> match against a pricing
    array already loaded from SharePoint ("List rows present in a table")
    -> Append to array variable with Unit Price/Total Price filled in ->
    HTTP (build-excel) with that array as the JSON body -> attach the
    response bytes directly to an email.

#1 (pdf-to-excel) remains available for direct testing or any flow that
doesn't need pricing.
"""

import logging
import json
import tempfile
from pathlib import Path

import azure.functions as func

from pdf_extractor import extract_items, write_to_excel

app = func.FunctionApp()


@app.route(route="pdf-to-excel", auth_level=func.AuthLevel.FUNCTION, methods=["POST"])
def pdf_to_excel(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("pdf-to-excel function triggered.")

    pdf_bytes = req.get_body()
    if not pdf_bytes:
        return _json_error("No PDF content received in request body.", 400)

    filename = req.headers.get("X-File-Name", "input.pdf")
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / filename
            pdf_path.write_bytes(pdf_bytes)

            logging.info(f"Processing {filename} ({len(pdf_bytes)} bytes)")
            items = extract_items(str(pdf_path))

            if not items:
                return _json_error(
                    f"No line items could be extracted from '{filename}'. "
                    "Check that this is a valid Uptick-generated Scope of Works PDF.",
                    422,
                )

            output_path = Path(tmpdir) / (pdf_path.stem + ".xlsx")
            write_to_excel(items, str(output_path))
            excel_bytes = output_path.read_bytes()

    except Exception as e:
        logging.exception(f"Error processing {filename}")
        return _json_error(f"Error processing PDF: {e}", 500)

    logging.info(f"Extracted {len(items)} line items from {filename}. Returning .xlsx.")

    return func.HttpResponse(
        body=excel_bytes,
        status_code=200,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{pdf_path.stem}.xlsx"'
        },
    )


@app.route(route="extract-items", auth_level=func.AuthLevel.FUNCTION, methods=["POST"])
def extract_items_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("extract-items function triggered.")

    pdf_bytes = req.get_body()
    if not pdf_bytes:
        return _json_error("No PDF content received in request body.", 400)

    filename = req.headers.get("X-File-Name", "input.pdf")
    if not filename.lower().endswith(".pdf"):
        filename += ".pdf"

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = Path(tmpdir) / filename
            pdf_path.write_bytes(pdf_bytes)

            logging.info(f"Processing {filename} ({len(pdf_bytes)} bytes)")
            items = extract_items(str(pdf_path))

    except Exception as e:
        logging.exception(f"Error extracting {filename}")
        return _json_error(f"Error extracting PDF: {e}", 500)

    if not items:
        return _json_error(
            f"No line items could be extracted from '{filename}'. "
            "Check that this is a valid Uptick-generated Scope of Works PDF.",
            422,
        )

    for item in items:
        item.setdefault("Sub-contractor COST ($)", "")
        item.setdefault("Sub-contractor Total ($)", "")
        item.setdefault("Anglicare Cost", "")
        item.setdefault("Anglicare Total ($)", "")
        item.setdefault("CBC Cost", "")
        item.setdefault("CBC Total ($)", "")
        item.setdefault("FINAL", "")
        item.setdefault("FINAL Total ($)", "")

    logging.info(f"Extracted {len(items)} line items from {filename}.")

    return func.HttpResponse(
        body=json.dumps(items),
        status_code=200,
        mimetype="application/json",
    )


@app.route(route="build-excel", auth_level=func.AuthLevel.FUNCTION, methods=["POST"])
def build_excel_endpoint(req: func.HttpRequest) -> func.HttpResponse:
    logging.info("build-excel function triggered.")

    try:
        items = req.get_json()
    except ValueError:
        return _json_error("Request body must be valid JSON (an array of items).", 400)

    if not isinstance(items, list) or not items:
        return _json_error("Request body must be a non-empty JSON array of items.", 400)

    filename = req.headers.get("X-File-Name", "output.xlsx")
    stem = filename
    if filename.lower().endswith(".xlsx") or filename.lower().endswith(".pdf"):
        stem = filename[:-4]

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = Path(tmpdir) / f"{stem}.xlsx"
            write_to_excel(items, str(output_path))
            excel_bytes = output_path.read_bytes()

    except Exception as e:
        logging.exception("Error building Excel file")
        return _json_error(f"Error building Excel file: {e}", 500)

    logging.info(f"Built .xlsx from {len(items)} priced items.")

    return func.HttpResponse(
        body=excel_bytes,
        status_code=200,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{stem}.xlsx"'
        },
    )


def _json_error(message: str, status_code: int) -> func.HttpResponse:
    return func.HttpResponse(
        body=json.dumps({"error": message}),
        status_code=status_code,
        mimetype="application/json",
    )
