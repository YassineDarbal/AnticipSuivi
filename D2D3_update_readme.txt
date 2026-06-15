D2D3 macro update export
========================

Purpose
-------
Use `D2D3_updated_macro_code.txt` as a paste-ready text export of the updated `D2D3.vb` module.

How to apply manually in Excel VBA
----------------------------------
1. Open your Excel workbook.
2. Press Alt+F11 to open the VBA editor.
3. Open the existing D2D3 module.
4. Replace the module contents with the contents of `D2D3_updated_macro_code.txt`.
5. Ensure the workbook contains a worksheet named exactly `D2D3_source`.
6. Load the Power Query output into `D2D3_source` using the M query you already have in `PowerQuery File M Code - Updated.txt`.
7. Compile the VBA project from Debug > Compile VBAProject.
8. Run `Extract_D2D3_MAJ`.

Important notes
---------------
- The hardcoded/commented password logic was intentionally left as-is.
- The macro no longer asks the user to select an external D2D3 workbook.
- The macro refreshes and reads the internal `D2D3_source` sheet.
- The macro normalizes D2D3 values in memory before writing to the suivi sheets, so it does not mutate the Power Query output columns directly.
