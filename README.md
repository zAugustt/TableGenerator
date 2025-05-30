# TableGenerator

TableGenerator is a Python application that reads data from Excel files and generates formatted tables in Microsoft Word documents. It provides a user-friendly GUI for selecting Excel files and customizing report options such as font, text style, and table layout.

## Features

- Read data from all sheets in an Excel file (`.xlsx` or `.xls`)
- Extracts headers and values, formats them, and writes them into Word tables
- Supports both vertical and horizontal table layouts
- Customizable font type, font size, text style (title or all caps), and header side
- Simple GUI built with Tkinter for easy configuration and file selection

## Requirements

- Python 3.11+
- [openpyxl](https://pypi.org/project/openpyxl/)
- [python-docx](https://pypi.org/project/python-docx/)
- Tkinter (usually included with Python)

## Installation

1. Clone this repository or download the source code.
2. Install the required Python packages:
   ```sh
   pip install openpyxl python-docx
   ```

## Usage

1. Run the application:
   ```sh
   python main.py
   ```
2. Use the GUI to:
   - Select an Excel file
   - Choose your preferred font, font size, text type, header side, and table ordering
   - Click "Generate Report" to create a Word document with formatted tables

The generated `.docx` file will be saved in the same directory as your Excel file, with a suffix indicating the table orientation.

## Project Structure

- `main.py` - Main application file containing all logic and GUI code

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.

## Acknowledgments

- [openpyxl](https://openpyxl.readthedocs.io/)
- [python-docx](https://python-docx.readthedocs.io/)
- Tkinter (Python standard library)