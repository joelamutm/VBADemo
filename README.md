# Excel VBA Automation Demo: `breakDown` & `exportPDF`

A beginner-friendly Excel VBA exercise designed to demonstrate how simple automation can eliminate repetitive manual work.

## Files in This Repository
* **`sample_data.xlsx`** — A plain sample workbook with 100 rows of practice data.
* **`SplitAndExportPDF.bas`** — The VBA module containing two macros:
  1. **`breakDown()`** — Copies 100 rows from the main sheet into 100 separate worksheets.
  2. **`exportPDF()`** — Saves every open worksheet in the workbook as its own individual PDF file.

## How to Test & Run

### 1. Download the Files
1. Download both `sample_data.xlsx` and `SplitAndExportPDF.bas` from this repository.
2. Place both files into a **new, empty folder** on your computer so the 100 generated PDFs stay organized in one place.

### 2. Import the VBA Module into Excel
1. Open `sample_data.xlsx` in Excel.
2. Press **`Alt + F11`** (Windows) or **`Option + F11`** (Mac) to open the VBA Editor.
3. Click **File** > **Import File...** and select `SplitAndExportPDF.bas`.
4. Close the VBA Editor window to return to Excel.

### 3. Run the Macros
1. Press **`Alt + F8`** (Windows) or **`Option + F8`** (Mac) to open the Macro list.
2. Select **`breakDown`** and click **Run** to split the 100 rows into 100 individual tabs.
3. Press **`Alt + F8`** again, select **`exportPDF`**, and click **Run** to export every tab as a separate PDF.
4. Check the folder where you saved your Excel file to view the 100 generated PDFs.

## License
This project is licensed under the [MIT License](LICENSE) — feel free to use, modify, and share it for your own learning and projects.
