# Excel Merger Tool

The Excel Merger Tool is a simple, graphical application designed to merge data from two Excel sheets based on a unique ID and sum specified quantities for each ID. It is built with Python and Tkinter, providing a user-friendly interface for non-technical users.

## Features

- **Merge Data**: Combines data from two specified Excel sheets based on unique IDs.
- **Sum Quantities**: Calculates the total quantities for each unique ID after merging.
- **Graphical Interface**: Simple and intuitive graphical user interface that requires no coding knowledge.

## Prerequisites

Before you start using the Excel Merger Tool, ensure you have the following installed:
- Python (version 3.6 or later)
- pip (Python package installer)

## Installation

1. **Clone or download this repository**:
   ```bash
   git clone https://github.com/copyleftdev/mergedaxlsx.git
   ```
   or download and extract the ZIP file.

2. **Install the required Python packages**:
   Navigate to the directory where you have saved the tool and run:
   ```bash
   pip install pandas
   pip install openpyxl
   ```

## Usage

To use the Excel Merger Tool, follow these steps:

1. **Launch the application**:
   Open your command line interface (CLI), navigate to the tool's directory, and run:
   ```bash
   python excel_gui.py
   ```

2. **Enter the Excel file details**:
   - **Excel File Path**: Enter the full path to your Excel file.
   - **Sheet 1**: Name of the first Excel sheet to merge.
   - **Sheet 2**: Name of the second Excel sheet to merge.

3. **Click 'Merge and Sum Quantities'**:
   After entering the file and sheet details, click the button to process the data. The results will be displayed in the application window.

4. **View Results**:
   The merged results and summed quantities will appear in the main text area of the application.

