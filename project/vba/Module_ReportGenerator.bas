' ═══════════════════════════════════════════════════════════════
' Module: Module_ReportGenerator
' Purpose: Generate individual and summary reports
' Author: Atlas Student Management System
' ═══════════════════════════════════════════════════════════════

Option Explicit

' ═══════════════════════════════════════════════════════════════
' Generate a report card for a single student
' ═══════════════════════════════════════════════════════════════
Sub GenerateStudentReport()
    Dim studentID As String
    studentID = InputBox("Enter Student ID for report (e.g. ST001):", "Report Card Generator")
    If studentID = "" Then Exit Sub
    
    Dim wsGrades As Worksheet
    Set wsGrades = ThisWorkbook.Sheets("Grades")
    
    ' Find student
    Dim foundRow As Range
    Set foundRow = wsGrades.Columns(1).Find(What:=studentID, LookAt:=xlWhole)
    
    If foundRow Is Nothing Then
        MsgBox "Student " & studentID & " not found!", vbExclamation
        Exit Sub
    End If
    
    Dim rowNum As Long
    rowNum = foundRow.Row
    
    ' Create new report sheet
    Dim reportName As String
    reportName = "Report_" & studentID
    
    ' Delete if exists
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(reportName)
    ws.Delete
    On Error GoTo 0
    
    ' Create new sheet
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = reportName
    
    ' ── Report Header ──
    ws.Range("B2").Value = "STUDENT REPORT CARD"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True
    ws.Range("B2:G2").Merge
    
    ws.Range("B4").Value = "Student ID:"
    ws.Range("C4").Value = studentID
    ws.Range("B5").Value = "Name:"
    ws.Range("C5").Value = wsGrades.Cells(rowNum, 2).Value
    ws.Range("B6").Value = "Class:"
    ws.Range("C6").Value = "Grade 10"
    
    For Each cell In ws.Range("B4:B6")
        cell.Font.Bold = True
    Next cell
    
    ' ── Subject Scores Table ──
    ws.Range("B8").Value = "SUBJECT"
    ws.Range("C8").Value = "SCORE"
    ws.Range("D8").Value = "GRADE"
    
    ws.Range("B8:D8").Font.Bold = True
    ws.Range("B8:D8").Interior.Color = RGB(68, 114, 196)
    ws.Range("B8:D8").Font.Color = RGB(255, 255, 255)
    
    Dim subjects As Variant, grades As Variant
    subjects = Array("English", "Mathematics", "Science", "ICT", "Social Studies")
    grades = Array("C", "D", "E", "B", "A") ' Just placeholder; real code computes these
    
    Dim i As Integer
    For i = 0 To 4
        ws.Cells(9 + i, 2).Value = subjects(i)
        ws.Cells(9 + i, 3).Value = wsGrades.Cells(rowNum, 3 + i).Value
        
        ' Compute grade
        Dim score As Double
        score = wsGrades.Cells(rowNum, 3 + i).Value
        Dim g As String
        If score >= 80 Then g = "A"
        ElseIf score >= 70 Then g = "B"
        ElseIf score >= 60 Then g = "C"
        ElseIf score >= 50 Then g = "D"
        Else: g = "F"
        End If
        ws.Cells(9 + i, 4).Value = g
    Next i
    
    ' ── Summary ──
    ws.Range("B15").Value = "Total:"
    ws.Range("C15").Value = wsGrades.Cells(rowNum, 8).Value
    ws.Range("B16").Value = "Average:"
    ws.Range("C16").Value = wsGrades.Cells(rowNum, 9).Value
    ws.Range("B17").Value = "Overall Grade:"
    ws.Range("C17").Value = wsGrades.Cells(rowNum, 10).Value
    ws.Range("B18").Value = "Remarks:"
    ws.Range("C18").Value = wsGrades.Cells(rowNum, 11).Value
    
    For Each cell In ws.Range("B15:B18")
        cell.Font.Bold = True
    Next cell
    
    ' ── Formatting ──
    ws.Range("C4:C6").ColumnWidth = 25
    ws.Range("B2:B18").ColumnWidth = 18
    ws.Range("C8:D8").ColumnWidth = 15
    
    ' Borders
    ws.Range("B8:D13").Borders.LineStyle = xlContinuous
    
    ' MsgBox
    MsgBox "Report generated on sheet: " & reportName, vbInformation, "Report Complete"
    
    ' Print preview
    ws.PageSetup.Orientation = xlPortrait
    ws.PrintPreview
End Sub

' ═══════════════════════════════════════════════════════════════
' Generate a summary report for entire class
' ═══════════════════════════════════════════════════════════════
Sub GenerateClassReport()
    Dim wsGrades As Worksheet
    Set wsGrades = ThisWorkbook.Sheets("Grades")
    
    Dim lastRow As Long
    lastRow = wsGrades.Cells(wsGrades.Rows.Count, "A").End(xlUp).Row
    
    ' Create sorted copy
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ClassSummary")
    ws.Delete
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ws.Name = "ClassSummary"
    
    ' Header
    ws.Range("A1").Value = "CLASS SUMMARY REPORT"
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").Font.Bold = True
    
    ws.Range("A3").Value = "Rank"
    ws.Range("B3").Value = "Student ID"
    ws.Range("C3").Value = "Name"
    ws.Range("D3").Value = "Average"
    ws.Range("E3").Value = "Grade"
    
    ws.Range("A3:E3").Font.Bold = True
    ws.Range("A3:E3").Interior.Color = RGB(68, 114, 196)
    ws.Range("A3:E3").Font.Color = RGB(255, 255, 255)
    
    ' Copy and sort data by average (descending)
    Dim students() As Variant
    ReDim students(lastRow - 2, 4)
    
    Dim i As Long
    For i = 2 To lastRow
        students(i - 2, 0) = i         ' original row
        students(i - 2, 1) = wsGrades.Cells(i, 1).Value  ' ID
        students(i - 2, 2) = wsGrades.Cells(i, 2).Value  ' Name
        students(i - 2, 3) = wsGrades.Cells(i, 9).Value  ' Average
        students(i - 2, 4) = wsGrades.Cells(i, 10).Value ' Grade
    Next i
    
    ' Simple bubble sort by average descending
    Dim j As Long, temp As Variant
    For i = LBound(students, 1) To UBound(students, 1) - 1
        For j = i + 1 To UBound(students, 1)
            If students(j, 3) > students(i, 3) Then
                ' Swap
                For k = 0 To 4
                    temp = students(i, k)
                    students(i, k) = students(j, k)
                    students(j, k) = temp
                Next k
            End If
        Next j
    Next i
    
    ' Write sorted data
    For i = 0 To UBound(students, 1)
        ws.Cells(i + 4, 1).Value = i + 1  ' Rank
        ws.Cells(i + 4, 2).Value = students(i, 1)  ' ID
        ws.Cells(i + 4, 3).Value = students(i, 2)  ' Name
        ws.Cells(i + 4, 4).Value = students(i, 3)  ' Average
        ws.Cells(i + 4, 5).Value = students(i, 4)  ' Grade
    Next i
    
    ' Format
    ws.Range("A3:E3").Borders.LineStyle = xlContinuous
    ws.Columns("A:E").AutoFit
    
    ' Footer with averages
    Dim summaryRow As Long
    summaryRow = lastRow - 2 + 5
    
    ws.Cells(summaryRow, 2).Value = "Class Average:"
    ws.Cells(summaryRow, 3).Value = _
        Application.WorksheetFunction.average(ws.Range(ws.Cells(4, 4), ws.Cells(summaryRow - 1, 4)))
    ws.Cells(summaryRow, 2).Font.Bold = True
    
    MsgBox "Class summary report generated on: ClassSummary sheet", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════
' Export reports to PDF
' ═══════════════════════════════════════════════════════════════
Sub ExportToPDF()
    Dim sheetName As String
    sheetName = InputBox("Enter sheet name to export to PDF:", "Export PDF", "ClassSummary")
    
    On Error GoTo NotFound
    ThisWorkbook.Sheets(sheetName).ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=ThisWorkbook.Path & "\" & sheetName & ".pdf", _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        OpenAfterPublish:=True
    Exit Sub
    
NotFound:
    MsgBox "Sheet '" & sheetName & "' not found!", vbExclamation
End Sub
