# VBA Modules — Import Guide

## How to Import VBA Modules into Excel

### Method 1: Manual Import (Any Excel version)

1. **Open your Excel workbook** (`.xlsx` → save as `.xlsm` to enable macros)
2. Press **`Alt + F11`** to open the VBA Editor
3. In the Project Explorer (left panel), right-click your workbook name
4. Select **Insert → Module**
5. Open the `.bas` file from this folder in a text editor
6. **Copy everything** (Ctrl+A, Ctrl+C)
7. **Paste** into the new module in the VBA Editor (Ctrl+V)
8. Repeat for each `.bas` file (one module per file)

### Method 2: File Import

1. Press **`Alt + F11`** to open the VBA Editor
2. Right-click your workbook → **Import File...**
3. Select the `.bas` file → click Open
4. The module appears in the Project Explorer

---

## Modules in This Folder

| File | Purpose | Key Macros |
|---|---|---|
| `Module_StudentCRUD.bas` | Add/Update/Delete/Search students | `AddStudent`, `UpdateStudent`, `DeleteStudent`, `SearchStudent`, `ListAllStudents` |
| `Module_GradeCalculator.bas` | Auto-calculate grades + charts | `CalculateAllGrades`, `GetLetterGrade`, `GenerateGradeChart` |
| `Module_ReportGenerator.bas` | Generate report cards + PDF export | `GenerateStudentReport`, `GenerateClassReport`, `ExportToPDF` |

---

## How to Run Macros

### From the Developer Tab
1. Enable Developer tab: File → Options → Customize Ribbon → check "Developer"
2. Go to **Developer tab → Macros**
3. Select the macro name → click **Run**

### Assign to a Button
1. Developer tab → **Insert → Button (Form Control)**
2. Draw a button on the sheet
3. Assign a macro from the list → click OK
4. Click the button to run

### Keyboard Shortcut
1. Developer tab → Macros → select macro → **Options**
2. Set a shortcut key (e.g., `Ctrl+Shift+G` for grade calculation)

---

## Enabling Macros

Excel may block macros by default. To enable:

1. **File → Options → Trust Center → Trust Center Settings**
2. **Macro Settings → Enable all macros** (or "Disable with notification")
3. Save workbook as **`.xlsm`** (Excel Macro-Enabled Workbook)

--- 

*Atlas — Learn VBA by understanding what the code does, not just running it.* 📚
