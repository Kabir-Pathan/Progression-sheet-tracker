VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public lastSelectedStudent As String

Private Sub ComboBox1_Change()

    Dim NamesRange As Range
    Dim Cell As Range
    Dim MatchedRow As Long

    Set NamesRange = ThisWorkbook.Sheets("Progression").Range("A1:A" & ThisWorkbook.Sheets("Progression").Cells(ThisWorkbook.Sheets("Progression").Rows.Count, "A").End(xlUp).row)

    ' Loop through names to find the match
    For Each Cell In NamesRange
        If Cell.Value = Me.ComboBox1.Value Then
            MatchedRow = Cell.row
            Exit For
        End If
    Next Cell

    ' If a match is found, select it
    If MatchedRow <> 0 Then
        ThisWorkbook.Sheets("Progression").Activate
        ThisWorkbook.Sheets("Progression").Cells(MatchedRow, 1).Select
    End If

End Sub


Private Sub CreateSheetAndOpen_Click()

    Dim templateSheet As Worksheet
    
    Dim StudentName As String
    StudentName = ActiveSheet.ComboBox1.Value
    lastSelectedStudent = StudentName

    ' Check if the sheet already exists
    For Each Sheet In Worksheets
        If Sheet.Name = StudentName Then
            Sheet.Activate
            Exit Sub
        End If
    Next Sheet

    ' If the sheet does not exist, create it
    Set templateSheet = ThisWorkbook.Sheets("Template")
    templateSheet.Copy After:=Sheets(Sheets.Count) 'creates a new worksheet
    ActiveSheet.Name = StudentName 'names the active worksheet as the student name

    ' Update Shape TextBox text to match student name
    ActiveSheet.Shapes("TB").TextFrame2.TextRange.Characters.Text = StudentName
    
    UpdateMasterSheet (StudentName)

End Sub
Private Sub Del1_Click()
    On Error Resume Next ' Enable error handling
    
    Dim StudentName As String
    StudentName = ActiveSheet.ComboBox1.Value
    
    ' Attempt to delete the specified sheet
    ThisWorkbook.Sheets(StudentName).Delete
    
    ' Check if an error occurred during the deletion
    If Err.Number <> 0 Then
        MsgBox "Report doesn't exist.", vbExclamation
        Err.Clear ' Clear the error object
    End If
    
    On Error GoTo 0 ' Disable error handling
End Sub
Private Sub Duplicate_Click()
    Dim NewWb As Workbook
    Dim Year As Variant  ' Changed type to Variant
    Dim importResponse As VbMsgBoxResult
    Dim currentYear As String
    Dim fileExists As Boolean

    ' Extract current year from the workbook name
    currentYear = Replace(ThisWorkbook.Name, "Progression Class of ", "")
    currentYear = Replace(currentYear, ".xlsm", "")

    ' Prompt the user to enter the year
    Year = Application.InputBox("Enter the graduation year for the new class:", "Year Input", Type:=2)
    
    ' Check if file already exists for the entered year
    fileExists = Dir(ThisWorkbook.Path & "\" & "Progression Class of " & Year & ".xlsm") <> ""
    
    ' Loop until a valid year is entered or Cancel is clicked
    While Year = False Or Year = "" Or Not IsNumeric(Year) Or Year = currentYear Or fileExists
        If Year = False Then
            ' User clicked "Cancel"
            Exit Sub
        ElseIf Year = currentYear Then
            ' User entered the same year as the current workbook
            MsgBox "Please enter a different year than the current workbook:", vbExclamation, "Invalid Input"
        ElseIf fileExists Then
            ' A workbook with the entered year already exists
            MsgBox "A workbook for the entered year already exists in the directory. Please enter a different year:", vbExclamation, "Invalid Input"
        Else
            ' User clicked "Ok" without entering a year
            MsgBox "Please type a year number:", vbExclamation, "Invalid Input"
        End If
        Year = Application.InputBox("Enter the graduation year for the new class:", "Year Input", Type:=2)
        fileExists = Dir(ThisWorkbook.Path & "\" & "Progression Class of " & Year & ".xlsm") <> ""
    Wend

    ' Copy current workbook and set it to newWb variable
    ThisWorkbook.SaveCopyAs ThisWorkbook.Path & "\" & "Progression Class of " & Year & ".xlsm"
    Set NewWb = Workbooks.Open(ThisWorkbook.Path & "\" & "Progression Class of " & Year & ".xlsm")
    
    ' Delete all sheets except the specified ones
    Dim ws As Worksheet
    Application.DisplayAlerts = False  ' Turn off warnings to delete sheets
    For Each ws In NewWb.Worksheets
        If ws.Name <> "Progression" And ws.Name <> "Template" And ws.Name <> "Emails" And ws.Name <> "Checkpoints" Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True   ' Turn warnings back on

    ' Activate the new workbook
    NewWb.Activate
    NewWb.Sheets("Progression").Shapes("TextBox 2").TextFrame2.TextRange.Characters.Text = "Progression Class of " & Year
    
    ' Clear data in B:K from row 6 onward
    NewWb.Sheets("Progression").Range("B6:K" & NewWb.Sheets("Progression").Cells(Rows.Count, "B").End(xlUp).row).ClearContents

    ' Ask the user if they want to import new names
    importResponse = MsgBox("Do you want to import new student names?", vbYesNo, "Import Names")

    ' If the user chose to import new names, call the import function
    If importResponse = vbYes Then
        ImportNewNames NewWb, Year
    End If
    
    ' Save and close the original workbook after you're done with it
    ThisWorkbook.Save
    ThisWorkbook.Close
End Sub
Private Sub ExtractAll_Click()
    Dim ws As Worksheet
    Dim studentCount As Integer
    Dim response As VbMsgBoxResult

    ' Prompt user for confirmation
    response = MsgBox("Are you sure you want to send a copy of each individual progression sheet to all students?", vbYesNo + vbQuestion, "Confirm Send")

    ' If the user selects "No", then exit the subroutine
    If response = vbNo Then Exit Sub

    studentCount = 0

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Assuming every sheet other than "Progression", "Template", "Emails", and "Checkpoints" is a student sheet.
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Progression" And ws.Name <> "Template" And ws.Name <> "Emails" And ws.Name <> "Checkpoints" Then
            ' Activate the student sheet
            ws.Activate
            ' Call your existing extraction and sending function here
            ExtractEach
            studentCount = studentCount + 1
        End If
    Next ws

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "A copy of each individual progression sheet has been sent to all " & studentCount & " students!"
    
    ThisWorkbook.Sheets("Progression").Activate
End Sub


Public Sub Worksheet_Activate()
    Dim comboBox As OLEObject
    Dim rangeToPopulate As Range
    Dim lastRow As Long
    Dim i As Integer
    Dim Cell As Range

    ' Find the last used row in column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).row

    ' Set the range from A1 to the last used row in column A
    Set rangeToPopulate = Range("A6:A" & lastRow)

    ' Find the ComboBox by its name
    Set comboBox = Me.OLEObjects("ComboBox1")

    ' Clear the existing items in the ComboBox
    comboBox.Object.Clear

    ' Add items to the ComboBox from the range
    i = 0
    For Each Cell In rangeToPopulate
        comboBox.Object.AddItem Cell.Value
        If Cell.Value = lastSelectedStudent Then
            comboBox.Object.ListIndex = i
        End If
        i = i + 1
    Next Cell

    ' If no previous selection or if the previous selection is not found, select the first item
    If comboBox.Object.ListIndex = -1 And comboBox.Object.ListCount > 0 Then
        comboBox.Object.ListIndex = 0
    End If
End Sub
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedRange As Range
    Set changedRange = Intersect(Target, Range("A6:A" & Rows.Count)) ' check if the change happened in column A from row 6 to the end

    If Not changedRange Is Nothing Then ' If change happened in column A
        Call Worksheet_Activate ' Refresh the ComboBox
    End If
End Sub
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim nameRange As Range
    Dim foundNameIndex As Integer
    Dim lastRow As Long
    
    ' Find the last used row in column A
    lastRow = Cells(Rows.Count, "A").End(xlUp).row

    ' Define the range of cells with names (adjust as necessary)
    Set nameRange = ThisWorkbook.Sheets("Progression").Range("A6:A" & lastRow)  ' Assuming names start at A6

    ' If the selected cell is within the names range
    If Not Intersect(Target, nameRange) Is Nothing Then
        ' Find the name in the ComboBox list
        foundNameIndex = FindInComboBox(ThisWorkbook.Sheets("Progression").ComboBox1, Target.Value)

        ' If the name is found in the ComboBox list
        If foundNameIndex <> -1 Then
            ' Set the ComboBox's value to the name
            ThisWorkbook.Sheets("Progression").ComboBox1.ListIndex = foundNameIndex
        End If
    End If
End Sub

Function FindInComboBox(cb As Object, searchItem As String) As Integer
    Dim i As Integer
    For i = 0 To cb.ListCount - 1
        If cb.List(i) = searchItem Then
            FindInComboBox = i
            Exit Function
        End If
    Next i
    FindInComboBox = -1
End Function







