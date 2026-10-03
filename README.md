# VBA

A collection of VBA modules for Microsoft Excel.

## Modules

### `HighlightDuplicates.bas`

Highlights duplicate values in the selected column with a light red fill.

| Macro | What it does |
| --- | --- |
| `HighlightDuplicatesInSelectedColumn` | Highlights every cell in the column whose value appears more than once. |
| `ClearDuplicateHighlights` | Removes the highlight added by the macro above from the selected column. |

How the column is chosen:

- Select a single cell or an entire column header to check the whole used part of that column.
- Select a multi-cell range within a column to check only that range.
- If the selection spans several columns, only the first one is checked.

Comparison is case-insensitive and ignores leading/trailing spaces. Blank cells and error values are skipped.

## Importing a module into Excel

1. Download the `.bas` file from this repository (e.g. `HighlightDuplicates.bas`).
2. Open your workbook in Excel.
3. Press **Alt + F11** (Windows) or **Option + F11** / **Fn + Option + F11** (Mac) to open the Visual Basic Editor.
4. In the Project pane, select the workbook you want to add the module to.
5. Choose **File → Import File…** (or right-click the project and choose **Import File…**), then pick the `.bas` file.
6. The module appears under **Modules** in that project. Close the Visual Basic Editor.
7. Save the workbook as a macro-enabled workbook (`.xlsm`), or as `PERSONAL.XLSB` to make the macros available in every workbook.

## Running a macro

1. Select a cell or range in the column you want to check.
2. Press **Alt + F8** (Windows) or **Option + F8** (Mac), choose `HighlightDuplicatesInSelectedColumn`, and click **Run**.

Optionally assign the macro to a keyboard shortcut via **Options…** in the Macro dialog, or to a button on the Quick Access Toolbar.

> If macros are blocked, enable them under **File → Options → Trust Center → Trust Center Settings → Macro Settings**, or unblock the downloaded file (right-click → **Properties** → **Unblock**) before importing.
