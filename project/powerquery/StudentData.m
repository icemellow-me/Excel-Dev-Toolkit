let
    // ═══════════════════════════════════════════════════════════════
    // Power Query M: Student Data Import & Cleaning
    // ═══════════════════════════════════════════════════════════════
    // HOW TO USE:
    // 1. Open Excel → Data tab → Get Data → From File → From Excel Workbook
    // 2. Select your student data file
    // 3. In Power Query Editor → Advanced Editor
    // 4. Paste this script
    // 5. Close & Load
    // 

    // ── Source: Read from Excel file ──
    Source = Excel.Workbook(File.Contents("C:\Data\students_raw.xlsx"), null, true),
    
    // ── Select the Students sheet ──
    StudentSheet = Source{[Item="Students", Kind="Sheet"]}[Data],
    
    // ── Promote first row as headers ──
    PromoteHeaders = Table.PromoteHeaders(StudentSheet, [PromoteAllScalars=true]),
    
    // ── Define column types ──
    TypedColumns = Table.TransformColumnTypes(PromoteHeaders, {
        {"StudentID", type text},
        {"First Name", type text},
        {"Last Name", type text},
        {"Gender", type text},
        {"Date of Birth", type date},
        {"Class", type text},
        {"House", type text},
        {"Parent Name", type text},
        {"Parent Phone", Int64.Type},
        {"Address", type text},
        {"Admission Date", type date},
        {"Attendance %", type number},
        {"English", type number},
        {"Mathematics", type number},
        {"Science", type number},
        {"ICT", type number},
        {"Social Studies", type number}
    }),
    
    // ── Clean: Trim text columns, remove nulls ──
    CleanNames = Table.TransformColumns(TypedColumns, {
        {"First Name", Text.Trim, type text},
        {"Last Name", Text.Trim, type text},
        {"Parent Name", Text.Trim, type text},
        {"Address", Text.Trim, type text}
    }),
    
    // ── Remove rows with empty StudentIDs ──
    RemoveBlanks = Table.SelectRows(CleanNames, each [StudentID] <> null and [StudentID] <> ""),
    
    // ── Add Full Name column ──
    AddFullName = Table.AddColumn(RemoveBlanks, "Full Name", 
        each [First Name] & " " & [Last Name], type text),
    
    // ── Add Total Score column ──
    AddTotal = Table.AddColumn(AddFullName, "Total",
        each [English] + [Mathematics] + [Science] + [ICT] + [Social Studies],
        type number),
    
    // ── Add Average column ──
    AddAverage = Table.AddColumn(AddTotal, "Average",
        each Number.Round([Total] / 5, 2),
        type number),
    
    // ── Add Letter Grade column ──
    AddGrade = Table.AddColumn(AddAverage, "Grade",
        each
            if [Average] >= 80 then "A"
            else if [Average] >= 70 then "B"
            else if [Average] >= 60 then "C"
            else if [Average] >= 50 then "D"
            else if [Average] >= 40 then "E"
            else "F",
        type text),
    
    // ── Add Remarks column ──
    AddRemarks = Table.AddColumn(AddGrade, "Remarks",
        each
            if [Grade] = "A" then "Excellent"
            else if [Grade] = "B" then "Very Good"
            else if [Grade] = "C" then "Good"
            else if [Grade] = "D" then "Pass"
            else if [Grade] = "E" then "Weak"
            else "Fail",
        type text),
    
    // ── Reorder columns ──
    ReorderColumns = Table.ReorderColumns(AddRemarks, {
        "StudentID", "First Name", "Last Name", "Full Name", "Gender",
        "Date of Birth", "Class", "House", "Parent Name", "Parent Phone",
        "Address", "Admission Date", "Attendance %",
        "English", "Mathematics", "Science", "ICT", "Social Studies",
        "Total", "Average", "Grade", "Remarks"
    })
in
    ReorderColumns
