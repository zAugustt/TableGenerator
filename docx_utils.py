from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_UNDERLINE
from helpers import add_percentages_to_values

def style_table(table, args: dict[str, str]) -> None:
    """
    Styles the table based on the provided arguments.
    :param table: The table to be styled.
    :param args: Dictionary containing report options such as total position, font type, font size, ordering, etc.
    """

    # for row in table.rows:
    #     for i, cell in enumerate(row.cells):
    #         if i == 0:
    #             cell.width = Inches(1)
    #         else:
    #             cell.width = Inches(1.5)

    # Remove gridlines
    # tbl = table._tbl
    # for tblBorders in tbl.xpath(".//w:tblBorders"):
    #     tblBorders.getparent().remove(tblBorders)

    if args.get("font_type"):
        for row in table.rows:
            for cell in row.cells:
                cell.paragraphs[0].runs[0].font.name = args.get("font_type")
                cell.paragraphs[0].runs[0].font.size = Pt(int(args.get("font_size")))
    if args.get("text_type") == "All Caps":
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.text = paragraph.text.upper()
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            for paragraph in cell.paragraphs:
                if "total" in paragraph.text.lower():
                    for run in paragraph.runs:
                        if args.get("total_position") == "inline":
                            run.bold = True
                            run.underline = WD_UNDERLINE.SINGLE
                        else:
                            run.bold = True
                            run.italic = True
                            run.text = run.text.upper()

                    connected_cell = row.cells[i - 1] if args.get("header_side") == "Right" else row.cells[i + 1]
                    for connected_paragraph in connected_cell.paragraphs:
                        for connected_run in connected_paragraph.runs:
                            if args.get("total_position") == "inline":
                                connected_run.bold = True
                                connected_run.underline = WD_UNDERLINE.SINGLE
                            else:
                                connected_run.bold = True
                                connected_run.italic = True
                                connected_run.text = connected_run.text.upper()

    # TODO: Formatting totals on top, bottom or inline (selector)
    return table


def write_doc(data: dict[str, dict[str, list[str]]], questions: list[str], output_path: str,
              args: dict[str, str]) -> None:
    """
    Writes dictionaries of headers and values into tables in a .docx file.
    :param data: Dictionary with sheet names as keys and dictionaries of headers and values as values.
    :param questions: List of question associated with each sheet.
    :param output_path: Path to save the generated Word document.
    :param args: Dictionary containing report options such as total position, font type, font size, ordering, etc.
    """

    document = Document()
    i = 0
    other_dict = ["other", "unsure", "refused", "no opinion"]
    for sheet_name, content in data.items():
        document.add_heading(sheet_name, level=1)
        document.add_paragraph(questions[i])
        if args.get("ordering") == "Vertical":
            table = document.add_table(rows=0, cols=2)
            table.style = 'Table Grid'
            for header, value in zip(content["headers"], content["values"]):
                if header and (header.lower() in other_dict or "total" in header.lower()):
                    blank_row = table.add_row()
                    blank_row.cells[0].text = ""
                    blank_row.cells[1].text = ""
                row = table.add_row()
                if args.get("header_side") == "Right":
                    value_cell = row.cells[0]
                    value_cell.text = add_percentages_to_values(value)
                    value_cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                    row.cells[1].text = str(header) if header is not None else ""
                else:
                    value_cell = row.cells[1]
                    value_cell.text = add_percentages_to_values(value)
                    value_cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    row.cells[0].text = str(header) if header is not None else ""

        elif args.get("ordering") == "Horizontal":
            num_cols = len(content["headers"])
            table = document.add_table(rows=2, cols=num_cols)
            table.style = 'Table Grid'
            for col, header in enumerate(content["headers"]):
                table.cell(0, col).text = str(header) if header is not None else ""
            for col, value in enumerate(content["values"]):
                value_cell = table.cell(1, col)
                value_cell.text = str(value) + "%" if value is not None else ""
                value_cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        style_table(table, args)
        i += 1

    document.save(output_path)