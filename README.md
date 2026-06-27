# VBA Macro-Enabled Balance Sheet Builder (AS)

This project generates a fully automated, macro-enabled Schedule III Division I (GAAP) Excel Workbook (`Balance_Sheet_Builder_Macro.xlsm`) completely identical to GrowthIQ's Balance Sheet Builder tool.

It integrates both the formula-based rollups and VBA automation macros for:
- **Import Trial Balance**: A native file dialog macro that imports a Trial Balance CSV/Excel sheet and rebuilds adjusted formulas.
- **Ledger Adjustments**: Automatically links adjustments posted on the `AE` sheet to the Trial Balance and updates Balance Sheet and P&L.
- **Export to PDF**: Compiles all statements and notes into a print-ready combined PDF.
- **Application Window UI Toggle**: Hides Excel gridlines, headings, and formula bar for a custom client dashboard look.

---

## 1. Project Directory Contents

- `build_vba_workbook.py`: Python automation script that constructs the sheets and injects VBA code.
- `vba_code/`:
  - `Navigation.bas`: Toggles Excel toolbars and runs README shapes navigation.
  - `ImportTB.bas`: Prompts file dialog and handles trial balance copying.
  - `ExportPDF.bas`: Exports reports array to PDF format.

---

## 2. Setup and Compilation Instructions

### Prerequisites
1. **Enable VBA Project Trust Settings** (Required for Python to write macros):
   - Open Microsoft Excel.
   - Go to **File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings**.
   - Check **"Trust access to the VBA project object model"** and click **OK**.
2. **Install Python Packages**:
   - Run in your terminal:
     ```bash
     pip install openpyxl pywin32
     ```

### Compile the Workbook
To compile the macro workbook, open a terminal in this directory and run:
```bash
python build_vba_workbook.py
```

- **Success**: Creates `Balance_Sheet_Builder_Macro.xlsm` containing the automated sheets, default sample data, and macros linked to native shapes on the `README` dashboard.
- **Fallback**: If Excel trust settings prevent script access, it saves the styled sheets as `Balance_Sheet_Builder_Template.xlsx` and gives instructions on manually importing the `.bas` files.
