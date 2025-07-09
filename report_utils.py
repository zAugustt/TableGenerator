from docx_utils import write_doc
from excel_utils import read_excel, get_question_data

def run_report(file_path: str, args: dict[str, str]) -> None:
    """
    Runs the report generation process by reading the Excel file, extracting data,
    and writing it to a Word document.
    :param file_path: Absolute path to the Excel file.
    :param args: Dictionary containing report options such as total position, font type, font size, and custom title.
    """
    suffix = "_v" if args.get("ordering") == "Vertical" else "_h" if args.get("ordering") == "Horizontal" else "_b"
    excel_data = read_excel(file_path)
    if file_path.lower().endswith('.xlsx'):
        output_file_path = file_path[:-5] + suffix + ".docx"
    elif file_path.lower().endswith('.xls'):
        output_file_path = file_path[:-4] + suffix + ".docx"
    else:
        exit("Invalid input file")
    pre_data = get_question_data(file_path)
    write_doc(excel_data, pre_data, output_file_path, args)
    print(f"Report written to {output_file_path}")
