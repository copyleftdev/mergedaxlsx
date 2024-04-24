import pandas as pd
from tkinter import Tk, Label, Entry, Button, Frame, Text, Scrollbar, VERTICAL, HORIZONTAL, END, N, S, E, W
from tkinter import messagebox
from tkinter.ttk import Style

def load_and_process_data(file_path, sheet1, sheet2, id_column, quantity_column):
    try:
        # Load data from specified sheets and columns
        df1 = pd.read_excel(file_path, sheet_name=sheet1, usecols=[id_column, quantity_column])
        df2 = pd.read_excel(file_path, sheet_name=sheet2, usecols=[id_column, quantity_column])

        # Merge the DataFrames on the ID column
        merged_df = pd.merge(df1, df2, on=id_column, how='left', suffixes=('_list1', '_list2'))

        # Handle missing values
        merged_df[f'{quantity_column}_list2'].fillna(0, inplace=True)

        # Calculate the total quantities
        merged_df['total_quantities'] = merged_df[f'{quantity_column}_list1'] + merged_df[f'{quantity_column}_list2']

        # Group by ID and sum the total quantities
        result_df = merged_df.groupby(id_column).agg({'total_quantities': 'sum'}).reset_index()
        return result_df
    except Exception as e:
        messagebox.showerror("Error", f"Failed to process the Excel file: {str(e)}")
        return None

def display_results(results, text_widget):
    text_widget.delete('1.0', END)
    if results is not None:
        for index, row in results.iterrows():
            line = f"ID: {row['ID']} - Total Quantities: {row['total_quantities']}\n"
            text_widget.insert(END, line)

def main():
    root = Tk()
    root.title("Excel Merger Tool")
    root.geometry("600x400")  # Initial size of the window

    # Applying a theme
    style = Style()
    style.theme_use('clam')

    # Main frame
    main_frame = Frame(root)
    main_frame.grid(sticky=N+S+E+W)

    # Make the main frame expandable
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=1)
    main_frame.rowconfigure(4, weight=1)

    # Input fields
    Label(main_frame, text="Excel File Path:").grid(row=0, column=0, sticky=W, padx=10, pady=5)
    file_path_entry = Entry(main_frame, width=50)
    file_path_entry.grid(row=0, column=1, sticky=E+W, padx=10)

    Label(main_frame, text="Sheet 1:").grid(row=1, column=0, sticky=W, padx=10, pady=5)
    sheet1_entry = Entry(main_frame, width=20)
    sheet1_entry.grid(row=1, column=1, sticky=W, padx=10)

    Label(main_frame, text="Sheet 2:").grid(row=2, column=0, sticky=W, padx=10, pady=5)
    sheet2_entry = Entry(main_frame, width=20)
    sheet2_entry.grid(row=2, column=1, sticky=W, padx=10)

    # Process button
    process_btn = Button(main_frame, text="Merge and Sum Quantities", command=lambda: display_results(
        load_and_process_data(
            file_path_entry.get(),
            sheet1_entry.get() or "Sheet1",
            sheet2_entry.get() or "Sheet2",
            "ID", "Quantities"),
        output_text))
    process_btn.grid(row=3, column=1, sticky=W, padx=10, pady=10)

    # Text widget with scroll bar for results
    text_frame = Frame(main_frame)
    text_frame.grid(row=4, column=0, columnspan=2, sticky=N+S+E+W, padx=10, pady=5)
    text_frame.columnconfigure(0, weight=1)
    text_frame.rowconfigure(0, weight=1)
    
    output_text = Text(text_frame, wrap="word", height=10)
    output_text.grid(row=0, column=0, sticky=N+S+E+W)
    
    scroll_y = Scrollbar(text_frame, orient=VERTICAL, command=output_text.yview)
    scroll_y.grid(row=0, column=1, sticky=N+S)
    output_text['yscrollcommand'] = scroll_y.set

    scroll_x = Scrollbar(text_frame, orient=HORIZONTAL, command=output_text.xview)
    scroll_x.grid(row=1, column=0, sticky=E+W)
    output_text['xscrollcommand'] = scroll_x.set

    root.mainloop()

if __name__ == "__main__":
    main()
