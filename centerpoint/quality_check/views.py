import os
from django.http import JsonResponse, HttpResponse, HttpResponseNotFound, HttpResponseServerError
from django.shortcuts import render
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def index(request):
    return render(request, 'quality_check/qc_interface.html')


def process_file(request):
    if request.method == 'POST' and request.FILES.get('excelFile'):
        excel_file = request.FILES['excelFile']
        df = pd.read_excel(excel_file, sheet_name='Listing')
        df['Remarks'] = ''
        if 'Barcode' in df.columns:
            duplicate_mask = df.duplicated(subset='Barcode', keep=False)
            duplicate_rows = df[duplicate_mask].index
            for row_num in duplicate_rows:
                df.at[row_num, 'Barcode'] = ''
                df.at[row_num, 'Remarks'] = f'Duplicate Barcode detected in row {row_num + 2}'

        processed_file_name = 'processed_file.xls'
        df.to_excel(f'quality_check/checking_data/{processed_file_name}', index=False)
        output_folder = os.path.join("quality_check", "checking_data")
        output_path = os.path.join(output_folder, processed_file_name)

        return render(request, 'quality_check/qc_interface.html', {
                'output_path': output_path.replace("\\", "/")
            })

    return JsonResponse({'error': 'Invalid request'})

def download_template(request, file_path):
    file_path = file_path.replace("/", "\\")
    file_name = os.path.basename(file_path)
    if not file_path:
        return HttpResponseNotFound("File path is missing")
    if not os.path.isfile(file_path):
        return HttpResponseNotFound("File not found")
    try:
        with open(file_path, 'rb') as file:
            response = HttpResponse(file.read(),
                                    content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            response['Content-Disposition'] = f'attachment; filename="{file_name}"'
            return response
    except Exception as e:
        return HttpResponseServerError("An error occurred while processing the request")