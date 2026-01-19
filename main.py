from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO

app = FastAPI()

@app.post("/excel")
def crear_excel(payload: dict):
    data = payload.get("data", [])
    if not data:
        return {"error": "No data provided"}

    filename = payload.get("filename", "archivo.xlsx")
    sheet = payload.get("sheet", "Datos")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet

    headers = list(data[0].keys())

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="4F81BD",
        end_color="4F81BD",
        fill_type="solid"
    )

    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill

    for row_idx, row in enumerate(data, start=2):
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row.get(header))

    # Guardamos en memoria
    output = BytesIO()
    wb.save(output)
    wb.close()
    output.seek(0)

    # Convertimos a bytes
    excel_bytes = output.getvalue()

    # Aquí va el headers en StreamingResponse
    return StreamingResponse(
        BytesIO(excel_bytes),  # flujo de bytes
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "Content-Length": str(len(excel_bytes)),  # muy importante
            "Cache-Control": "no-store"
        }
    )
