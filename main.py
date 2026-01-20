from fastapi import FastAPI, File, UploadFile
from fastapi.responses import HTMLResponse, StreamingResponse
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill
from io import BytesIO
from jinja2 import Template

app = FastAPI()

@app.post("/excel")
def crear_excel(payload: dict):
    print("Payload recibido:", payload)
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

html_template = """
<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
  <thead>
    <tr>
      {% for header in headers %}
      <th style="background-color: {{ header.bg_color }}; color: {{ header.font_color }}; font-weight: {{ header.font_weight }};">
        {{ header.value }}
      </th>
      {% endfor %}
    </tr>
  </thead>
  <tbody>
    {% for row in rows %}
    <tr>
      {% for cell in row %}
      <td style="background-color: {{ cell.bg_color }}; color: {{ cell.font_color }}; font-weight: {{ cell.font_weight }};">
        {{ cell.value }}
      </td>
      {% endfor %}
    </tr>
    {% endfor %}
  </tbody>
</table>
"""

def get_cell_styles(cell):
    bg_color = "transparent"

    if cell.fill and cell.fill.fill_type == "solid":
        color = cell.fill.start_color
        if color and color.rgb:
            bg_color = f"#{str(color.rgb)[-6:]}"

    font_color = "black"
    if cell.font and cell.font.color:
        color = cell.font.color
        if color and color.rgb:
            font_color = f"#{str(color.rgb)[-6:]}"

    font_weight = "bold" if cell.font and cell.font.bold else "normal"

    return {
        "bg_color": bg_color,
        "font_color": font_color,
        "font_weight": font_weight,
    }

COLUMNAS_PERMITIDAS = {"Nombre", "Correo", "Telefono"}

@app.post("/procesar_excel", response_class=HTMLResponse)
async def procesar_excel(file: UploadFile = File(...)):
    contents = await file.read()
    wb = load_workbook(BytesIO(contents), data_only=True)
    ws = wb.active

    # Leer headers y mapear índice
    header_cells = list(ws[1])
    headers = []
    columnas_idx = []

    for idx, cell in enumerate(header_cells):
        if cell.value in COLUMNAS_PERMITIDAS:
            columnas_idx.append(idx)
            headers.append({
                "value": cell.value or "",
                **get_cell_styles(cell)
            })

    # Leer solo 6 filas y solo columnas permitidas
    rows = []
    for i, row in enumerate(ws.iter_rows(min_row=2), start=1):
        if i > 6:
            break

        fila = []
        for idx in columnas_idx:
            cell = row[idx]
            fila.append({
                "value": cell.value or "",
                **get_cell_styles(cell)
            })
        rows.append(fila)

    html = Template(html_template).render(headers=headers, rows=rows)
    return html




