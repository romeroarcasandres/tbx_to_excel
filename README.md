# TBX to Excel Converter

The **TBX to Excel Converter** is a Python-based tool that converts TermBase eXchange (`.tbx`) files into Excel (`.xlsx`) format. The converter provides both **interactive** and **automatic** modes to configure which fields to extract, rename them if desired, and generate a structured, column-based Excel file for further use.

---

## Overview

This tool allows users to:
- **Parse TBX files:** Extract terms, languages, and associated metadata from TBX files.
- **Interactive Field Selection:** Choose which TBX fields to include in the Excel output.
- **Optional Field Renaming:** Keep original TBX field names or rename them for clarity.
- **Automatic Mode:** Skip interaction and export all available fields as-is.
- **Export to Excel:** Create a well-structured `.xlsx` file with auto-adjusted column widths.
- **Summary Statistics:** View a summary of extracted entries, languages, and columns.

---

## Requirements

- **Python 3**
- **Libraries:**
  - `pandas`
  - `xlsxwriter` (preferred for Excel writing)
  - `openpyxl` (fallback if needed)

Install dependencies via:

```bash
pip install pandas xlsxwriter openpyxl
```

---

## Files

- `tbx_to_excel.py` – The main script file that implements TBX parsing and Excel conversion.

---

## Usage

1. **Run the Script:**
   ```bash
   python tbx_to_excel.py input_file.tbx
   ```

2. **Interactive Configuration (default mode):**
   - The script scans the TBX file and displays all available fields.
   - You select which fields to include by typing their numbers (e.g., `1,3,5`) or `all`.
   - You can optionally rename the selected fields.

3. **Automatic Mode:**
   - Skip interaction by using the `--auto` flag.  
   - All available fields are included with their original names.
   ```bash
   python tbx_to_excel.py input_file.tbx --auto
   ```

4. **Specify Output File (optional):**
   ```bash
   python tbx_to_excel.py input_file.tbx -o output.xlsx
   ```

5. **View Summary (optional):**
   - Add `-s` or `--summary` to print details after conversion.
   ```bash
   python tbx_to_excel.py input_file.tbx -s
   ```

---

## Example Workflow

- Convert with full interactivity:
  ```bash
  python tbx_to_excel.py terminology.tbx
  ```
- Convert automatically, saving to `terminology.xlsx`:
  ```bash
  python tbx_to_excel.py terminology.tbx --auto
  ```
- Convert automatically and specify output file:
  ```bash
  python tbx_to_excel.py terminology.tbx -o export.xlsx --auto
  ```

---

## Important Notes

- **Namespaces Handling:**  
  The script automatically detects and registers XML namespaces common in TBX files.
  
- **Field Flexibility:**  
  Both entry-level (`entry_id`, `entry_descrip_*`) and term-level (`term`, `termNote_*`, `descrip_*`) fields can be included.
  
- **Excel Output:**  
  - Column widths are auto-adjusted for readability.  
  - If multiple terms exist in the same language, they are suffixed (e.g., `en_term`, `en_term_2`).  

- **Performance:**  
  For very large TBX files, interactive scanning may take some time. Use `--auto` for faster batch conversion.

---

## License

This project is licensed under the [CC4 License](LICENSE). Please refer to the LICENSE file for additional details.
