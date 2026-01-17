from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO

app = FastAPI()


@app.post("/excel")
def crear_excel(payload: dict):

    data = payload["data"]          # filas
    filename = payload.get("filename", "archivo.xlsx")
    sheet = payload.get("sheet", "Datos")

    wb = Workbook()
    ws = wb.active
    ws.title = sheet

    headers = list(data[0].keys())

    # Estilo encabezado
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(
        start_color="4F81BD",
        end_color="4F81BD",
        fill_type="solid"
    )

    # Escribir encabezados
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill

    # Escribir filas
    for row_idx, row in enumerate(data, start=2):
        for col_idx, header in enumerate(headers, start=1):
            ws.cell(row=row_idx, column=col_idx, value=row.get(header))

    # Guardar en memoria
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"'
        }
    )
