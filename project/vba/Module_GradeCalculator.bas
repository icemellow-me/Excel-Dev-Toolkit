' ═══════════════════════════════════════════════════════════════
' Module: Module_GradeCalculator
' Purpose: Automated grade calculation using VBA
' Author: Atlas Student Management System
' ═══════════════════════════════════════════════════════════════

Option Explicit

' ── Constants ──
Const GRADES_SHEET As String = "Grades"
Const FIRST_DATA_ROW As Long = 2
' Columns: A=ID, B=Name, C=English, D=Math, E=Science, F=ICT, G=Social
'          H=Total, I=Average, J=Grade, K=Remarks

' ═══════════════════════════════════════════════════════════════
' Main: Calculate all grades
' ═══════════════════════════════════════════════════════════════
Sub CalculateAllGrades()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GRADES_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < FIRST_DATA_ROW Then
        MsgBox "No student data found!", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim totalScore As Double, avgScore As Double
    Dim letterGrade As String, remarks As String
    
    For i = FIRST_DATA_ROW To lastRow
        ' Calculate Total (columns C through G)
        totalScore = 0
        Dim j As Integer
        For j = 3 To 7
            If IsNumeric(ws.Cells(i, j).Value) Then
                totalScore = totalScore + ws.Cells(i, j).Value
            End If
        Next j
        
        ' Calculate Average
        avgScore = totalScore / 5
        
        ' Write Total and Average
        ws.Cells(i, 8).Value = totalScore
        ws.Cells(i, 9).Value = Round(avgScore, 2)
        
        ' Determine Letter Grade
        letterGrade = GetLetterGrade(avgScore)
        ws.Cells(i, 10).Value = letterGrade
        
        ' Determine Remarks
        remarks = GetRemarks(letterGrade)
        ws.Cells(i, 11).Value = remarks
        
        ' Color-code the grade cell
        Call ColorGradeCell(ws.Cells(i, 10), letterGrade)
    Next i
    
    ' Add class statistics
    Call AddClassStatistics(ws, lastRow)
    
    Application.ScreenUpdating = True
    MsgBox "Grades calculated for " & (lastRow - 1) & " students.", vbInformation, "Complete"
End Sub

' ═══════════════════════════════════════════════════════════════
' Get letter grade from numeric score
' ═══════════════════════════════════════════════════════════════
Function GetLetterGrade(avgScore As Double) As String
    Select Case avgScore
        Case Is >= 80: GetLetterGrade = "A"
        Case Is >= 70: GetLetterGrade = "B"
        Case Is >= 60: GetLetterGrade = "C"
        Case Is >= 50: GetLetterGrade = "D"
        Case Is >= 40: GetLetterGrade = "E"
        Case Else:     GetLetterGrade = "F"
    End Select
End Function

' ═══════════════════════════════════════════════════════════════
' Get remarks based on grade
' ═══════════════════════════════════════════════════════════════
Function GetRemarks(grade As String) As String
    Select Case grade
        Case "A": GetRemarks = "Excellent — Outstanding performance"
        Case "B": GetRemarks = "Very Good — Above average performance"
        Case "C": GetRemarks = "Good — Satisfactory performance"
        Case "D": GetRemarks = "Pass — Needs improvement"
        Case "E": GetRemarks = "Weak — Requires significant improvement"
        Case "F": GetRemarks = "Fail — Must repeat examination"
        Case Else: GetRemarks = "N/A"
    End Select
End Function

' ═══════════════════════════════════════════════════════════════
' Color-code grade cells (green=best, red=fail)
' ═══════════════════════════════════════════════════════════════
Sub ColorGradeCell(cell As Range, grade As String)
    Select Case grade
        Case "A"
            cell.Interior.Color = RGB(0, 176, 80)    ' Green
            cell.Font.Color = RGB(255, 255, 255)
        Case "B"
            cell.Interior.Color = RGB(146, 208, 80)  ' Light green
            cell.Font.Color = RGB(0, 0, 0)
        Case "C"
            cell.Interior.Color = RGB(255, 255, 0)   ' Yellow
            cell.Font.Color = RGB(0, 0, 0)
        Case "D"
            cell.Interior.Color = RGB(255, 192, 0)   ' Orange
            cell.Font.Color = RGB(0, 0, 0)
        Case "E"
            cell.Interior.Color = RGB(255, 128, 0)   ' Dark orange
            cell.Font.Color = RGB(255, 255, 255)
        Case "F"
            cell.Interior.Color = RGB(255, 0, 0)     ' Red
            cell.Font.Color = RGB(255, 255, 255)
    End Select
    cell.Font.Bold = True
    cell.HorizontalAlignment = xlCenter
End Sub

' ═══════════════════════════════════════════════════════════════
' Add class statistics at the bottom of the sheet
' ═══════════════════════════════════════════════════════════════
Sub AddClassStatistics(ws As Worksheet, lastRow As Long)
    Dim statsRow As Long
    statsRow = lastRow + 3
    
    ' Headers
    ws.Cells(statsRow, 1).Value = "CLASS STATISTICS"
    ws.Cells(statsRow, 1).Font.Bold = True
    ws.Cells(statsRow, 1).Font.Size = 13
    
    ' Subject averages
    ws.Cells(statsRow + 1, 2).Value = "Subject Averages:"
    ws.Cells(statsRow + 1, 2).Font.Bold = True
    
    Dim subjects As Variant
    subjects = Array("English", "Mathematics", "Science", "ICT", "Social Studies")
    
    Dim k As Integer
    For k = 0 To 4
        ws.Cells(statsRow + 2, k + 3).Value = subjects(k)
        ws.Cells(statsRow + 3, k + 3).Value = _
            Application.WorksheetFunction.Average( _
                ws.Range(ws.Cells(FIRST_DATA_ROW, k + 3), ws.Cells(lastRow, k + 3)))
    Next k
    
    ' Overall average
    ws.Cells(statsRow + 5, 2).Value = "Class Average:"
    ws.Cells(statsRow + 5, 2).Font.Bold = True
    ws.Cells(statsRow + 5, 3).Value = _
        Application.WorksheetFunction.Average( _
            ws.Range(ws.Cells(FIRST_DATA_ROW, 9), ws.Cells(lastRow, 9)))
    ws.Cells(statsRow + 5, 3).NumberFormat = "0.00"
    
    ' Top performer
    ws.Cells(statsRow + 6, 2).Value = "Top Performer:"
    ws.Cells(statsRow + 6, 2).Font.Bold = True
    
    Dim maxAvgRow As Long
    maxAvgRow = FIRST_DATA_ROW
    Dim i As Long
    For i = FIRST_DATA_ROW To lastRow
        If ws.Cells(i, 9).Value > ws.Cells(maxAvgRow, 9).Value Then
            maxAvgRow = i
        End If
    Next i
    ws.Cells(statsRow + 6, 3).Value = ws.Cells(maxAvgRow, 2).Value & _
        " (" & ws.Cells(maxAvgRow, 9).Value & ")"
    
    ' Pass rate
    Dim passCount As Long
    passCount = 0
    For i = FIRST_DATA_ROW To lastRow
        If ws.Cells(i, 9).Value >= 50 Then passCount = passCount + 1
    Next i
    
    ws.Cells(statsRow + 7, 2).Value = "Pass Rate:"
    ws.Cells(statsRow + 7, 2).Font.Bold = True
    ws.Cells(statsRow + 7, 3).Value = Format(passCount / (lastRow - 1), "0.0%")
End Sub

' ═══════════════════════════════════════════════════════════════
' Generate a chart from the grades data
' ═══════════════════════════════════════════════════════════════
Sub GenerateGradeChart()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(GRADES_SHEET)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Delete existing charts
    Dim obj As ChartObject
    For Each obj In ws.ChartObjects
        obj.Delete
    Next obj
    
    ' Create bar chart
    Dim chartObj As ChartObject
    Set chartObj = ws.ChartObjects.Add(Left:=300, Width:=400, Top:=10, Height:=300)
    
    With chartObj.Chart
        .ChartType = xlBarClustered
        .SetSourceData ws.Range(ws.Cells(FIRST_DATA_ROW, 2), ws.Cells(lastRow, 9))
        .HasTitle = True
        .ChartTitle.Text = "Student Average Scores"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Students"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Average Score (%)"
    End With
    
    MsgBox "Chart generated!", vbInformation
End Sub
