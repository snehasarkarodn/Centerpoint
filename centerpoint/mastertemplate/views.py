from django.shortcuts import render
from django.http import HttpResponse, FileResponse
from django.views.decorators.csrf import csrf_exempt
from IPython.display import display
import pandas as pd
import ipywidgets as widgets
import shutil
import os
import uuid
import openpyxl
import json
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from .models import ProcessedData
import datetime

def process_english(english_string):
    unique_values = []
    seen_values = set()
    for value in english_string.split("|"):
        if value not in seen_values:
            seen_values.add(value)
            unique_values.append(value)
    return unique_values

def color_code_columns(sheet, selected_df, mandatory_fixed):
    for idx, header in enumerate(sheet[1], start=1):
        if header.value in mandatory_fixed.columns:
            col_letter = get_column_letter(idx)
            is_mandatory = mandatory_fixed.at[
                               0, header.value] == 'Yes' if header.value in mandatory_fixed.columns else False

            cell = sheet.cell(row=1, column=idx)
            if is_mandatory:
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00",
                                                        fill_type="solid")  # Green
            else:
                cell.fill = openpyxl.styles.PatternFill(start_color="FFFF00", end_color="FFFF00",
                                                        fill_type="solid")  # Yellow

        elif header.value in selected_df["Field Name"].values:
            row_index = selected_df.index[selected_df["Field Name"] == header.value].tolist()[0]
            is_mandatory = selected_df.at[row_index, "Mandatory"] == 'yes' if header.value in selected_df[
                "Field Name"].values else False

            is_Field_Type = selected_df.at[row_index, "Field Type"] == 'select' or selected_df.at[
                row_index, "Field Type"] == 'multi-select' if header.value in selected_df[
                "Field Name"].values else False

            cell = sheet.cell(row=1, column=idx)
            if is_mandatory:
                cell.fill = openpyxl.styles.PatternFill(start_color="00FF00", end_color="00FF00",
                                                        fill_type="solid")  # Green
            else:
                cell.fill = openpyxl.styles.PatternFill(start_color="FFFF00", end_color="FFFF00",
                                                        fill_type="solid")  # Yellow

            if is_Field_Type:
                cell.font = openpyxl.styles.Font(color="FF0000")

def add_hidden_data_validation(sheet, col_idx, max_length, hidden_col_idx):
    col_letter = get_column_letter(col_idx)
    dv = DataValidation(type="list", formula1=f'hidden_sheet!${get_column_letter(hidden_col_idx)}$2:${get_column_letter(hidden_col_idx)}${max_length + 1}')

    # Apply data validation to the entire column in the second row
    for r in range(2, 100):
        sheet.add_data_validation(dv)
        dv.add(sheet.cell(row=r, column=col_idx))

def add_temp_data_validation(sheet, idx, max_temp_len, hidden_col_idx):
    col_letter = get_column_letter(idx)
    dv = DataValidation(type="list", formula1=f'hidden_sheet1!${get_column_letter(hidden_col_idx)}$2:${get_column_letter(hidden_col_idx)}${max_temp_len + 1}')

    # Apply data validation to the entire column in the second row
    for r in range(2, 100):
        sheet.add_data_validation(dv)
        dv.add(sheet.cell(row=r, column=idx))


def copy_and_modify_master_temp(selected_df, mandatory_fixed, selected_values, default_template, template_dropdown):
    master_temp_path = os.path.join("mastertemplate","Centerpoint_master_template", "CP Content_Template.xlsx")

    if not default_template.equals(pd.read_excel(master_temp_path)):
        default_template.to_excel(os.path.join("mastertemplate","Centerpoint_master_template", "CP Content_Template.xlsx"), index=False)
    original_filename, extension = os.path.splitext(os.path.basename(master_temp_path))
    unique_id = str(uuid.uuid4())[:8]
    output_folder = os.path.join("mastertemplate","stored_data")
    output_path = os.path.join(output_folder, f"{original_filename}_{unique_id}{extension}")
    shutil.copyfile(master_temp_path, output_path)

    wb = openpyxl.load_workbook(output_path)
    sheet = wb.active

    hidden_sheet = wb.create_sheet("hidden_sheet")
    hidden_sheet.sheet_state = "hidden"

    hidden_sheet1 = wb.create_sheet("hidden_sheet1")
    hidden_sheet1.sheet_state = "hidden"
    max_temp_len=0

    processed_english_df = pd.DataFrame()
    max_length = 0

    # add columns as per selected category
    for index, row in selected_df.iterrows():
        field_type_value = row["Field Name"]
        if field_type_value not in sheet[1]:
            sheet.cell(row=1, column=sheet.max_column + 1, value=field_type_value)

        if row.Processed_English != ['nan'] and len(row.Processed_English) > 0:
            max_length = max(max_length, len(row.Processed_English))

    for index, row in selected_df.iterrows():
        field_type_value = row["Field Name"]

        if row.Processed_English != ['nan'] and len(row.Processed_English) > 0:
            padded_values = row.Processed_English + [None] * (max_length - len(row.Processed_English))
            processed_english_df[field_type_value] = padded_values

    # Putting data in the hidden sheet
    headers = list(processed_english_df.columns)
    for c_idx, header in enumerate(headers, start=1):
        hidden_sheet.cell(row=1, column=c_idx, value=header)

    for r_idx, row in enumerate(processed_english_df.itertuples(index=False, name=None), start=2):
        for c_idx, value in enumerate(row, start=1):
            hidden_sheet.cell(row=r_idx, column=c_idx, value=value)

    # Putting data in the hidden sheet1
    headers1 = list(template_dropdown.columns)
    for c_idx, col_name in enumerate(template_dropdown.columns, start=1):
        hidden_sheet1.cell(row=1, column=c_idx, value=col_name)

    for r_idx, row in template_dropdown.iterrows():
        for c_idx, value in enumerate(row, start=1):
            hidden_sheet1.cell(row=r_idx + 2, column=c_idx, value=value)

    max_temp_len = template_dropdown.shape[0]

    # Adding Data Validation
    for idx, header in enumerate(sheet[1], start=1):
        if header.value in headers:
            col_letter = get_column_letter(idx)
            hidden_col_idx = headers.index(header.value) + 1 if header.value in headers else None

            if hidden_col_idx is not None:
                for r in range(2, hidden_sheet.max_row + 1):
                    add_hidden_data_validation(sheet, idx, max_length, hidden_col_idx)
            else:
                for r in range(2, sheet.max_row + 1):
                    sheet.cell(row=r, column=idx, value=None)

        elif header.value in headers1:
            col_letter = get_column_letter(idx)
            hidden_col_idx = headers1.index(header.value) + 1 if header.value in headers1 else None

            if hidden_col_idx is not None:
                for r in range(2, hidden_sheet1.max_row + 1):
                    add_temp_data_validation(sheet, idx, max_temp_len, hidden_col_idx)
            else:
                for r in range(2, sheet.max_row + 1):
                    sheet.cell(row=r, column=idx, value=None)

        elif header.value == "Base Product ID":
            col_letter = get_column_letter(idx)
            for r in range(2, 100):
                style_no_cell = sheet.cell(row=r, column=1)
                sheet[f"{col_letter}{r}"].value = f'=D{r}& "CP" & TEXT(TODAY(), "DD-MM-YYYY")'

    color_code_columns(sheet, selected_df, mandatory_fixed)

    wb.save(output_path)

    processed_data_instance = ProcessedData.objects.create(
        unique_id=unique_id,
        output_path=output_path,
        selected_values=selected_values
    )

    return output_path


def index(request):
    Sheet_ID_1 = '1hGJINSRtlDs9yEXH5wI9eGazN9WUgLvS'
    sheet = pd.ExcelFile(f"https://docs.google.com/spreadsheets/d/{Sheet_ID_1}/export?format=xlsx")
    merged_temp = pd.read_excel(sheet)
    dropdown_values = [i.strip() for i in merged_temp["Template Name"].unique()]
    return render(request, 'mastertemplate/user_interface1.html', {'dropdown_values': dropdown_values})

def main_func(request):
    #master_temp = pd.read_excel(r"C:\Users\Lenovo\Desktop\centerpoint\CP Content_Template.xlsx")
    Sheet_ID_1 = '1hGJINSRtlDs9yEXH5wI9eGazN9WUgLvS'
    sheet = pd.ExcelFile(f"https://docs.google.com/spreadsheets/d/{Sheet_ID_1}/export?format=xlsx")
    merged_temp = pd.read_excel(sheet, sheet_name="Attribute and Values")
    default_template = pd.read_excel(sheet, sheet_name="CP Content_Template")
    template_dropdown = pd.read_excel(sheet, sheet_name="Template_dropdown_values")
    mandatory_fixed = pd.read_excel(sheet, sheet_name="Template_Mandatory")

    dropdown_values = [i.strip() for i in merged_temp["Template Name"].unique()]
    selected_df = pd.DataFrame()

    if request.method == 'POST':
        selected_values_str = request.POST.get('selected_values', '[]')
        selected_values = json.loads(selected_values_str)

        for selected_template in selected_values:
            selected_rows = merged_temp[merged_temp["Template Name"].str.contains(selected_template)]
            if not selected_rows.empty:
                selected_df = pd.concat([selected_df, selected_rows], ignore_index=True)

        selected_df.drop(columns=["Template Name"], inplace=True)
        selected_df["English"] = selected_df["English"].astype(str)
        selected_df["English"] = (selected_df.groupby("Field Name")["English"].transform(lambda x: "|".join(x.unique())))
        selected_df = selected_df.drop_duplicates(subset="Field Name")
        selected_df["Processed_English"] = selected_df["English"].apply(lambda x: process_english(x))

        output_path = copy_and_modify_master_temp(selected_df, mandatory_fixed, selected_values, default_template, template_dropdown)

        if os.path.exists(output_path):
            return render(request, 'mastertemplate/user_interface1.html', {
                'dropdown_values': dropdown_values,
                'output_path': output_path.replace("\\", "/")
            })
        else:
            return HttpResponse("File not found.")

    # Render the initial form page
    return render(request, 'your_app_name/your_template_name.html', {
        'dropdown': dropdown,
        'button': button,
        'output': output
    })

def download_template(request,file_path):
    file_path=file_path.replace("/", "\\")
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
        logger.error(str(e))
        return HttpResponseServerError("An error occurred while processing the request")