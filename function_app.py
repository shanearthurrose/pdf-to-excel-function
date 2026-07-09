"""
Azure Function: PDF (Uptick/Anglicare SoW) -> Excel

HTTP-triggered function intended to be called from Power Automate.

Contract:
    POST /api/pdf-to-excel
    Headers:
        x-functions-key: <function key>              (auth)
        Content-Type:    application/pdf
        X-File-Name:     original filename (optional, used to name output)
    Body:
        Raw PDF bytes (NOT base64 — send as binary content)

    Response (200):
        Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
        Body: raw .xlsx bytes

    Response (4xx/5xx):
        Content-Type: application/json
        Body: {"error": "..."}

Power Automate should take the HTTP action's response body directly and
pass it into a "Create file" action (SharePoint/OneDrive) as the file content.
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

    # Power Automate can pass the original filename in a header so the
    # output file / logs are traceable back to the source email attachment.
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


def _json_error(message: str, status_code: int) -> func.HttpResponse:
    return func.HttpResponse(
        body=json.dumps({"error": message}),
        status_code=status_code,
        mimetype="application/json",
    )
