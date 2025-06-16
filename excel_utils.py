import openpyxl
from helpers import format_headers, format_values

def read_excel(file_path: str) -> dict[str, dict[str, list[str]]]:
    """
    Reads an Excel file and extracts headers from the first column and values from the second column
    for all sheets. Returns a dictionary with sheet names as keys and extracted data as values.
    :param file_path: Absolute path to file
    """
    workbook = openpyxl.load_workbook(file_path)
    data = {}

    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        headers = [row[0] for row in sheet.iter_rows(
            min_row=1, max_row=sheet.max_row, min_col=1, max_col=1, values_only=True
        )]
        values = [row[0] for row in sheet.iter_rows(
            min_row=1, max_row=sheet.max_row, min_col=2, max_col=2, values_only=True
        )]

        try:
            start_index = values.index(1) + 1
            filtered_headers = headers[start_index:]
            filtered_headers = format_headers(filtered_headers)
            filtered_values = values[start_index:]
            filtered_values = format_values(filtered_values)
        except ValueError:
            filtered_headers = []
            filtered_values = []

        data[sheet_name] = {"headers": filtered_headers, "values": filtered_values}

    return data


def get_questions(file_path: str) -> list[str]:
    """
    Reads an Excel file and extracts the questions on each sheet
    :param file_path: Absolute path to file
    """
    workbook = openpyxl.load_workbook(file_path)
    questions = []
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]

        headers = [row[0] for row in sheet.iter_rows(
            min_row=1, max_row=sheet.max_row, min_col=1, max_col=1, values_only=True
        )]
        questions.append(headers[2])
    return questions



