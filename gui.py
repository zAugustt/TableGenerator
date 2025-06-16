from tkinter import Tk, filedialog, Label, Button, Entry, StringVar, OptionMenu, Frame, Spinbox
import os
from main import run_report

def open_gui() -> None:
    """
    Opens a GUI for selecting an Excel file and configuring report options.
    The GUI allows the user to select a file, choose the position of totals, font type, font size,
    text type, and run the report generation process.
    The GUI uses Tkinter for the interface.
    """
    root = Tk()
    root.title("Report Writer")
    root.geometry("350x550")

    file_frame = Frame(root)
    file_frame.pack(pady=10)
    file_label = Label(file_frame, text="No file selected")
    file_label.pack(side="left")

    def select_file():
        path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        file_var.set(path)
        file_label.config(text=os.path.basename(path) if path else "No file selected")

    file_var = StringVar()
    Button(
        file_frame,
        text="Browse",
        command=select_file,
        activebackground="#cccccc",
        activeforeground="#000000"
    ).pack(side="left", padx=5)

    common_fonts = ["Arial", "Calibri", "Times New Roman", "Verdana", "Courier New", "Georgia", "Tahoma", "Helvetica"]
    Label(root, text="Font Type:").pack()
    font_type_var = StringVar(value="Tahoma")
    font_dropdown = OptionMenu(root, font_type_var, *common_fonts)
    font_dropdown.pack()

    Label(root, text="\nOr enter a custom font:").pack()
    custom_font_var = StringVar()
    Entry(root, textvariable=custom_font_var).pack()

    font_warning = Label(root, text="", fg="red")
    font_warning.pack()

    def on_font_entry(*args):
        if custom_font_var.get():
            font_warning.config(text="Warning: Custom font may not be available in Word.")
            font_type_var.set(custom_font_var.get())
        else:
            font_warning.config(text="")

    custom_font_var.trace_add("write", on_font_entry)

    Label(root, text="Font Size:").pack()
    font_size_var = StringVar(value="9")
    Spinbox(
        root,
        values=tuple(str(i) for i in range(6, 25)),
        textvariable=font_size_var,
        width=5,
        state="readonly"
    ).pack()
    font_size_var.set("9")

    total_position_var = StringVar(value="Top")
    Label(root, text="\nTotals Position: (Not implemented)").pack()
    OptionMenu(root, total_position_var, "Top", "Bottom", "Inline").pack()

    Label(root, text="\nText Type:").pack()
    text_type_var = StringVar(value="Title")
    OptionMenu(root, text_type_var, "Title", "All Caps").pack()

    Label(root, text="\nHeader Side:").pack()
    header_side_var = StringVar(value="Right")
    OptionMenu(root, header_side_var, "Right", "Left").pack()

    Label(root, text="\nOrdering:").pack()
    ordering_var = StringVar(value="Vertical")
    OptionMenu(root, ordering_var, "Vertical", "Horizontal").pack()

    def on_run():
        if not file_var.get():
            file_label.config(text="Please select a file!")
            return
        args = {
            "total_position": total_position_var.get(),
            "font_type": font_type_var.get(),
            "font_size": font_size_var.get(),
            "text_type": text_type_var.get(),
            "header_side": header_side_var.get(),
            "ordering": ordering_var.get()
        }
        run_report(file_var.get(), args)
        root.quit()

    Button(
        root,
        text="Generate Report",
        command=on_run,
        activebackground="#cccccc",
        activeforeground="#000000"
    ).pack(pady=10)
    root.mainloop()