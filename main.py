from typing import List
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

app = FastAPI()

@app.post("/excel")
def crear_excel(payload: List[dict]):
    if not payload:
        raise HTTPException(status_code=400, detail="No payload received")
    
    # n8n envuelve todo en un array con "json", tomamos el primer elemento
    payload_dict = payload[0].get("json", {})

    data = payload_dict.get("data", [])
    if not data:
        raise HTTPException(status_code=400, detail="No data provided")

    filename = payload_dict.get("filename", "archivo.xlsx")
    sheet = payload_dict.get("sheet", "Datos")

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

    temp_file = f"/tmp/{filename}"
    wb.save(temp_file)

    return FileResponse(
        path=temp_file,
        filename=filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
