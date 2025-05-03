import io
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment
from django.shortcuts import render
from django.http import HttpResponse

def upload_pdf(request):
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_bytes = request.FILES['pdf_file'].read()

        # 1) Жёсткие настройки для чётких таблиц
        table_settings = {
            "vertical_strategy":   "lines",    # ищем явные вертикальные линии
            "horizontal_strategy": "lines",    # и горизонтальные
            "intersection_tolerance": 3,       # допустимый разброс пересечений
        }

        # 2) Извлекаем данные
        all_data = []
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                tbl = page.extract_table(table_settings=table_settings)
                if tbl:
                    all_data.extend(tbl)

        # 3) Формируем DataFrame (или сразу пишем в Workbook)
        headers, *rows = all_data if all_data else ([], [])
        wb = Workbook()
        ws = wb.active
        if headers: ws.append(headers)
        for r in rows: ws.append(r)

        # 4) Настраиваем фильтры, закрепление и выравнивание
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            max_len = max(len(str(c.value)) for c in col if c.value)
            ws.column_dimensions[col[0].column_letter].width = max_len + 2
            for c in col:
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # 5) Отдаём пользователю
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)
        resp = HttpResponse(
            buf.read(),
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        resp['Content-Disposition'] = 'attachment; filename="converted.xlsx"'
        return resp

    return render(request, 'upload.html')