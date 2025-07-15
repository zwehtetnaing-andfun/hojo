# ***Hojo*** Excel Comparison Program

This Python program compares Excel files in paired folders (V1 and V2), highlights differences in yellow, and generates detailed Markdown and Excel reports. It processes subfolders within a `recompare` directory, logs operations for debugging, and recalculates formulas in Excel files using Microsoft Excel.

## Prerequisites

Before running the program, ensure the following are installed:

- **Python 3.8 or higher**: Required to run the script.
- **Windows Operating System**: The program uses `pywin32` for Excel formula recalculation, which is Windows-specific.
- **Microsoft Excel**: Required for recalculating formulas in Excel files.
- **Required Python Libraries**: Listed in the `requirements.txt` file (see Installation).

## Installation

1. **Install Python**:

   - Download and install Python 3.8+ from python.org.
   - Ensure `pip` is available and added to your system PATH.

2. **Install Dependencies**:

   - Save the `requirements.txt` file in the same directory as `excel_comparison.py`.
   - Open a terminal or command prompt in the directory and run:

     ```bash
     pip install -r requirements.txt
     ```
   - This installs `openpyxl` and `pywin32`. Note: `tkinter` is typically included with Python.

3. **Verify Microsoft Excel**:

   - Ensure Microsoft Excel is installed, as the program uses it to recalculate formulas.

4. **Download the Program**:

   - Save `excel_comparison.py` and `requirements.txt` to your working directory.

## requirements.txt

Ensure your `requirements.txt` contains the following:

```
openpyxl>=3.0.10
pywin32>=306
```

## Usage

1. **Prepare Your Files**:

   - Create a `recompare` folder with the following structure:

     ```
     recompare/
     ├── 00000000/
     │   ├── V1/
     │   │   ├── file1.xlsx
     │   │   ├── file2.xlsx
     │   ├── V2/
     │   │   ├── file1.xlsx
     │   │   ├── file2.xlsx
     │   ├── result/
     ├── 00000001/
     │   ├── V1/
     │   ├── V2/
     │   ├── result/
     ```
   - Place Excel files to compare in `V1` and `V2` subfolders. Matching files must have the same base name (e.g., `file1.xlsx` in both `V1` and `V2`).

2. **Run the Program**:

   - Navigate to the directory containing `excel_comparison.py` in a terminal or command prompt.
   - Execute the script:

     ```bash
     python excel_comparison.py
     ```
   - A dialog will prompt you to select the `recompare` folder.

3. **Program Workflow**:

   - The program will:
     - Recalculate formulas in V2 Excel files using Microsoft Excel.
     - Compare files in `V1` and `V2`, highlighting differences in yellow in the V2 files.
     - Save modified V2 files with a prefix (`O_` for no differences, `X_` for differences) in the `result` subfolder.
     - Generate a Markdown report (`<timestamp>_comparison_report.md`) and an Excel report (`<timestamp>_comparison_report.xlsx`) summarizing results.
   - Logs are saved in the `logs` folder with timestamps for debugging.

4. **View Results**:

   - Check the `result` subfolder in each subfolder for modified Excel files.
   - Review the Markdown and Excel reports in the working directory for a summary of differences.
   - Check the `logs` folder for detailed execution logs.

## Notes

- **Error Handling**: If errors occur (e.g., missing files or folders), they are logged, and a message box will display the issue. Check the logs for details.
- **File Formats**: The program supports `.xlsx` and `.xls` files.
- **Time Range Handling**: Time ranges (e.g., `9:00~17:00`) are normalized for consistent comparison.
- **Dependencies**: Ensure `requirements.txt` is up-to-date. If you add new libraries, update the file and reinstall dependencies.
- **Windows Only**: The `pywin32` library and Excel formula recalculation require a Windows environment.

## Troubleshooting

- **Excel Not Found**: Ensure Microsoft Excel is installed and accessible.
- **Permission Issues**: Run the terminal as an administrator if you encounter file access errors.
- **Missing Dependencies**: Verify `requirements.txt` installation with `pip list`.
- **Log Files**: Check the `logs` folder for detailed error messages if the program fails.