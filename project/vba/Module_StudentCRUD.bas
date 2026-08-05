' ═══════════════════════════════════════════════════════════════
' Module: Module_StudentCRUD
' Purpose: CRUD operations for student records in Excel
' Author: Atlas Student Management System
' Date: 2026
'
' HOW TO USE:
' 1. Open Excel, press Alt+F11 to open the VBA Editor
' 2. Insert > Module
' 3. Paste this entire file
' 4. Run macros from Developer tab > Macros, or assign to buttons
' ═══════════════════════════════════════════════════════════════

Option Explicit

' ── Constants ──
Const SHEET_NAME As String = "Students"
Const ID_COL As Integer = 1
Const NAME_COL As Integer = 2

' ═══════════════════════════════════════════════════════════════
' Add a new student record
' ═══════════════════════════════════════════════════════════════
Sub AddStudent()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    
    ' Find the next empty row
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, ID_COL).End(xlUp).Row + 1
    
    ' Generate Student ID
    Dim studentID As String
    studentID = "ST" & Format(nextRow - 1, "000")
    
    ' Get student details via input
    Dim firstName As String, lastName As String
    Dim gender As String, className As String
    Dim house As String, parentName As String
    Dim parentPhone As String, address As String
    
    firstName = InputBox("Enter first name:", "New Student")
    If firstName = "" Then Exit Sub
    
    lastName = InputBox("Enter last name:", "New Student")
    gender = InputBox("Enter gender (Male/Female):", "New Student")
    className = InputBox("Enter class (e.g. Grade 10):", "New Student")
    house = InputBox("Enter house (Blue/Red/Green/Yellow):", "New Student")
    parentName = InputBox("Enter parent/guardian name:", "New Student")
    parentPhone = InputBox("Enter parent phone number:", "New Student")
    address = InputBox("Enter address (city):", "New Student")
    
    ' Write to sheet
    ws.Cells(nextRow, 1).Value = studentID
    ws.Cells(nextRow, 2).Value = firstName
    ws.Cells(nextRow, 3).Value = lastName
    ws.Cells(nextRow, 4).Value = gender
    ws.Cells(nextRow, 6).Value = className
    ws.Cells(nextRow, 7).Value = house
    ws.Cells(nextRow, 8).Value = parentName
    ws.Cells(nextRow, 9).Value = Val(parentPhone)
    ws.Cells(nextRow, 10).Value = address
    ws.Cells(nextRow, 11).Value = "2022-09-05"  ' Admission date
    
    ' Format the new row
    ws.Rows(nextRow).Font.Name = "Calibri"
    ws.Rows(nextRow).Font.Size = 11
    
    MsgBox "Student " & firstName & " " & lastName & " added successfully!" & vbCrLf & _
           "Student ID: " & studentID, vbInformation, "Success"
           
    Call RefreshFormulas(ws)
End Sub

' ═══════════════════════════════════════════════════════════════
' Update an existing student record
' ═══════════════════════════════════════════════════════════════
Sub UpdateStudent()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    
    Dim searchID As String
    searchID = InputBox("Enter Student ID to update (e.g. ST001):", "Update Student")
    If searchID = "" Then Exit Sub
    
    ' Find the row
    Dim foundRow As Range
    Set foundRow = ws.Columns(ID_COL).Find(What:=searchID, LookAt:=xlWhole)
    
    If foundRow Is Nothing Then
        MsgBox "Student ID " & searchID & " not found!", vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim rowNum As Long
    rowNum = foundRow.Row
    
    ' Show current values and ask for updates
    Dim field As String, newVal As String
    field = InputBox("Which field to update?" & vbCrLf & _
                     "1=FirstName, 2=LastName, 3=Gender, 4=Class, 5=House" & vbCrLf & _
                     "6=ParentName, 7=Phone, 8=Address, 9=Attendance", "Update Student")
    
    Select Case field
        Case "1": ws.Cells(rowNum, 2).Value = InputBox("New first name:", "Update", ws.Cells(rowNum, 2).Value)
        Case "2": ws.Cells(rowNum, 3).Value = InputBox("New last name:", "Update", ws.Cells(rowNum, 3).Value)
        Case "3": ws.Cells(rowNum, 4).Value = InputBox("New gender:", "Update", ws.Cells(rowNum, 4).Value)
        Case "4": ws.Cells(rowNum, 6).Value = InputBox("New class:", "Update", ws.Cells(rowNum, 6).Value)
        Case "5": ws.Cells(rowNum, 7).Value = InputBox("New house:", "Update", ws.Cells(rowNum, 7).Value)
        Case "6": ws.Cells(rowNum, 8).Value = InputBox("New parent name:", "Update", ws.Cells(rowNum, 8).Value)
        Case "7": ws.Cells(rowNum, 9).Value = Val(InputBox("New phone:", "Update", ws.Cells(rowNum, 9).Value))
        Case "8": ws.Cells(rowNum, 10).Value = InputBox("New address:", "Update", ws.Cells(rowNum, 10).Value)
        Case "9": ws.Cells(rowNum, 12).Value = Val(InputBox("New attendance %:", "Update", ws.Cells(rowNum, 12).Value))
        Case Else: MsgBox "Invalid selection", vbExclamation
    End Select
    
    MsgBox "Record updated!", vbInformation
End Sub

' ═══════════════════════════════════════════════════════════════
' Delete a student record
' ═══════════════════════════════════════════════════════════════
Sub DeleteStudent()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    
    Dim searchID As String
    searchID = InputBox("Enter Student ID to delete:", "Delete Student")
    If searchID = "" Then Exit Sub
    
    Dim foundRow As Range
    Set foundRow = ws.Columns(ID_COL).Find(What:=searchID, LookAt:=xlWhole)
    
    If foundRow Is Nothing Then
        MsgBox "Student ID " & searchID & " not found!", vbExclamation
        Exit Sub
    End If
    
    Dim confirm As VbMsgBoxResult
    confirm = MsgBox("Delete " & ws.Cells(foundRow.Row, 2).Value & " " & _
                     ws.Cells(foundRow.Row, 3).Value & "?", vbYesNo + vbExclamation, "Confirm Delete")
    
    If confirm = vbYes Then
        ws.Rows(foundRow.Row).Delete
        MsgBox "Student record deleted.", vbInformation
    End If
End Sub

' ═══════════════════════════════════════════════════════════════
' Search for a student
' ═══════════════════════════════════════════════════════════════
Sub SearchStudent()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    
    Dim searchTerm As String
    searchTerm = InputBox("Enter Student ID or Name to search:", "Search")
    If searchTerm = "" Then Exit Sub
    
    ' Search by ID or Name
    Dim foundRow As Range
    Set foundRow = ws.Columns(ID_COL).Find(What:=searchTerm, LookAt:=xlWhole)
    
    If foundRow Is Nothing Then
        Set foundRow = ws.Columns(2).Find(What:=searchTerm, LookAt:=xlPart)
    End If
    
    If foundRow Is Nothing Then
        MsgBox "No student found matching '" & searchTerm & "'", vbExclamation
        Exit Sub
    End If
    
    ' Highlight and display
    ws.Cells(foundRow.Row, ID_COL).Select
    MsgBox "Found: " & ws.Cells(foundRow.Row, 2).Value & " " & _
           ws.Cells(foundRow.Row, 3).Value & vbCrLf & _
           "ID: " & ws.Cells(foundRow.Row, 1).Value & vbCrLf & _
           "Gender: " & ws.Cells(foundRow.Row, 4).Value & vbCrLf & _
           "Class: " & ws.Cells(foundRow.Row, 6).Value & vbCrLf & _
           "House: " & ws.Cells(foundRow.Row, 7).Value, vbInformation, "Student Found"
End Sub

' ═══════════════════════════════════════════════════════════════
' Helper: Refresh formulas after data changes
' ═══════════════════════════════════════════════════════════════
Sub RefreshFormulas(ws As Worksheet)
    Application.Calculate
    ws.Calculate
End Sub

' ═══════════════════════════════════════════════════════════════
' List all students in a message box
' ═══════════════════════════════════════════════════════════════
Sub ListAllStudents()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, ID_COL).End(xlUp).Row
    
    Dim summary As String
    summary = "Total Students: " & (lastRow - 1) & vbCrLf & vbCrLf
    
    Dim i As Long
    For i = 2 To lastRow
        summary = summary & ws.Cells(i, 1).Value & " - " & _
                  ws.Cells(i, 2).Value & " " & ws.Cells(i, 3).Value & vbCrLf
    Next i
    
    MsgBox summary, vbInformation, "Student List"
End Sub
