import os
import pdfplumber
import pandas as pd
from django.shortcuts import render
from django.http import HttpResponse
from django.conf import settings
from openpyxl.styles import Alignment


def upload_pdf(request):
    if request.method == 'POST' and request.FILES.get('pdf_file'):
        pdf_file = request.FILES['pdf_file']
        file_path = os.path.join(settings.MEDIA_ROOT, 'temp', pdf_file.name)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        with open(file_path, 'wb+') as dest:
            for chunk in pdf_file.chunks():
                dest.write(chunk)

        # Извлекаем таблицу
        all_data = []
        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                table = page.extract_table()
                if table:
                    all_data.extend(table)

        # Формируем DataFrame с заголовками из первой строки
        if all_data:
            headers = all_data[0]
            data_rows = all_data[1:]
            df = pd.DataFrame(data_rows, columns=headers)
        else:
            df = pd.DataFrame()

        # Сохраняем в Excel с фильтрами, закреплённой шапкой, авто-шириной и выравниванием
        excel_path = os.path.join(settings.MEDIA_ROOT, 'converted.xlsx')
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1')
            ws = writer.sheets['Sheet1']

            # Закрепляем первую строку и включаем фильтры
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

            # Авто-ширина и выравнивание по центру
            for column_cells in ws.columns:
                # вычисляем максимальную длину в столбце
                max_length = max(
                    len(str(cell.value)) if cell.value is not None else 0
                    for cell in column_cells
                )
                col_letter = column_cells[0].column_letter
                # ставим ширину чуть больше максимальной длины
                ws.column_dimensions[col_letter].width = max_length + 2

                # выравниваем все ячейки столбца по центру
                for cell in column_cells:
                    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

        # Отдаём файл пользователю
        with open(excel_path, 'rb') as f:
            resp = HttpResponse(
                f.read(),
                content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
            resp['Content-Disposition'] = 'attachment; filename="converted.xlsx"'
            return resp

    return render(request, 'upload.html')