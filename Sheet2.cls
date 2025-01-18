VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Back_Click()
Worksheets("Progression").Activate
End Sub
Private Sub Clear_Click()
    Dim rng1 As Range
    Dim rng2 As Range
    Dim oleObj As Object
    
    Set rng1 = ActiveSheet.Range("A6:R14")
    Set rng2 = ActiveSheet.Range("A16:R24")
    
    ' Clear the contents and comments of the ranges
    rng1.ClearContents
    rng1.ClearComments
    rng2.ClearContents
    rng2.ClearComments
    
    ' Loop through all OLEObjects (ActiveX Controls) on the worksheet
    For Each oleObj In ActiveSheet.OLEObjects
        ' Check if the OLEObject is a checkbox
        If oleObj.progID = "Forms.CheckBox.1" Then
            ' Delete the checkbox
            oleObj.Delete
        End If
    Next oleObj
End Sub


Private Sub Del2_Click()
    If ActiveSheet.Name = "Template" Or ActiveSheet.Name = "Progression" Then
        MsgBox "You are not allowed to delete this sheet.", vbExclamation, "Action Not Allowed"
    Else
        ActiveSheet.Delete
        ThisWorkbook.Sheets("Progression").Activate
    End If
End Sub
Private Sub UpdateSalud_Click()
UpdateSaludStatusForAllRows
UpdateNotes
End Sub
Private Sub Extract_Click()
    Dim NewWb As Workbook
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExtStr As String
    Dim OutApp As Object
    Dim OutMail As Object
    Dim EmailAddr As String
    Dim StudentName As String
    Dim SplitName() As String
    Dim FormattedName As String
    Dim response As VbMsgBoxResult
    

    ' Prompt user for confirmation
    response = MsgBox("Are you sure you want to send a copy of this sheet to the student?", vbYesNo + vbQuestion, "Confirm Send")

    ' If the user selects "No", then exit the subroutine
    If response = vbNo Then Exit Sub

    ' Get student name from the active sheet's name
    StudentName = ActiveSheet.Name
    SplitName = Split(StudentName, ", ")
    If UBound(SplitName) >= 1 Then
        FormattedName = Trim(SplitName(1)) & " " & Trim(SplitName(0))
    Else
        FormattedName = StudentName
    End If

    ' Fetch email address using the function
    EmailAddr = GetEmailFromName(StudentName, ThisWorkbook.Sheets("Emails"))

    ' Exit if no email found
    If EmailAddr = "" Then
        MsgBox "No email found for " & StudentName, vbExclamation, "Error"
        Exit Sub
    End If

    ' Create a new workbook with the active sheet
    ThisWorkbook.Sheets(StudentName).Copy
    Set NewWb = ActiveWorkbook

    ' Define file path, name, and extension
    TempFilePath = Environ$("temp") & "\"
    TempFileName = "Student Progression Sheet " & Format(Now, "dd-mmm-yy")
    FileExtStr = ".xlsx"

    ' Save the new workbook to temp folder
    Application.DisplayAlerts = False
    NewWb.SaveAs TempFilePath & TempFileName & FileExtStr, FileFormat:=51 ' 51 represents xlsx format
    Application.DisplayAlerts = True

    ' Create an email and send
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = EmailAddr
        .CC = ""
        .BCC = ""
        .Subject = "Your Progression Sheet"
        .Body = "Dear " & FormattedName & ", attached is a copy of your progression sheet."
        .Attachments.Add NewWb.FullName
        .Send
    End With
    On Error GoTo 0

    ' Close and don't save changes to the new workbook
    NewWb.Close SaveChanges:=False

    ' Delete the temporary file
    Kill TempFilePath & TempFileName & FileExtStr

    ' Notify completion
    MsgBox "Progression Sheet sent to " & StudentName

    ' Clean up
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub
Function GetEmailFromName(StudentName As String, ws As Worksheet) As String
    Dim rng As Range
    Dim Cell As Range

    Set rng = ws.Range("A:A")

    ' Loop through each cell in the range
    For Each Cell In rng
        If Cell.Value = StudentName Then
            GetEmailFromName = Cell.Offset(0, 1).Value ' returns the email from column B
            Exit Function
        End If
    Next Cell

    GetEmailFromName = "" ' return empty if no match found
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = ActiveSheet 'Set reference to the active sheet
    
    'Exit if the active sheet is the "Template" sheet
    If ws.Name = "Template" Then Exit Sub

    'If the changed cell is in the case type column
    If Target.Column = 3 And Target.Cells.Count = 1 Then
        FillStepsAndSalud ws, Target.row
    ElseIf Target.Column >= 4 And Target.Column <= 15 Then
        UpdateSalud ws, Target.row
    End If
End Sub


Sub FillStepsAndSalud(ws As Worksheet, targetRow As Long)
    Dim checkpoints() As Variant
    checkpoints = GetCheckpoints(ws.Cells(targetRow, 3).Value)
    
    Dim i As Long
    'Clear previous values/comments
    ws.Range(ws.Cells(targetRow, 4), ws.Cells(targetRow, 15)).ClearContents
    ws.Range(ws.Cells(targetRow, 4), ws.Cells(targetRow, 15)).ClearComments
    ws.Range(ws.Cells(targetRow, 4), ws.Cells(targetRow, 15)).Validation.Delete 'Remove previous validation
    
    'Fill in the new checkpoints (if any)
    For i = LBound(checkpoints) To UBound(checkpoints)
        With ws.Cells(targetRow, 3 + i)
            .Value = "N"
            .AddComment Text:=checkpoints(i)
            
            With .Comment.Shape
                .Width = 65  ' adjust as needed
                .Height = 35  ' adjust as needed
    
                ' Check if the comment contains "(opt)" and adjust the background color accordingly
                If InStr(1, checkpoints(i), "(opt)", vbTextCompare) > 0 Then
                    .Fill.ForeColor.RGB = RGB(173, 216, 230) ' Light blue color
                Else
                    .Fill.ForeColor.RGB = RGB(232, 232, 232) ' Original color you had
                End If
            End With
    
            'Add data validation dropdown
            With .Validation
                .Delete 'Clear any existing validation rules
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                    xlBetween, Formula1:="Y,N"
                .ShowInput = True
                .ShowError = True
            End With
        End With
    Next i

    
    'Fill N/A for the remaining Step columns
    Dim colHeader As String
    For j = i + 3 To 15
        colHeader = ws.Cells(5, j).Value 'Assuming headers are in row 1
        If InStr(1, colHeader, "Step", vbTextCompare) > 0 Then
            ws.Cells(targetRow, j).Value = "---"
        Else
            Exit For 'Exit the loop if no further "Step" columns are detected
        End If
    Next j
    
    UpdateSalud ws, targetRow
End Sub

Sub UpdateSalud(ws As Worksheet, targetRow As Long)
    Dim allYes As Boolean
    allYes = True 'Assume all steps are marked as Y to begin with
    
    Dim i As Integer
    For i = 4 To 15
        If ws.Cells(targetRow, i).Value = "---" Then
            Exit For 'Stop checking when encountering the first N/A
        ElseIf Not (InStr(1, ws.Cells(targetRow, i).NoteText, "(opt)", vbTextCompare) > 0) And ws.Cells(targetRow, i).Value <> "Y" Then
            'Check if the comment does NOT contain "(opt)" and the cell value is NOT "Y"
            allYes = False
            Exit For 'Stop checking if any non-optional step is not marked as Y
        End If
    Next i
    
    'Update the "Salud" column
    If allYes Then
        ws.Cells(targetRow, 16).Value = ChrW(&H2713) 'Assuming "Salud" is in column 16
    Else
        ws.Cells(targetRow, 16).Value = "X"
    End If
End Sub
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Target.Column >= 4 And Target.Column <= 15 Then
        Cancel = True 'Prevent going into edit mode
        If Target.Value = "Y" Then
            Target.Value = "N"
        Else
            Target.Value = "Y"
        End If
        ' Call your UpdateSalud function here if needed
        UpdateSalud ActiveSheet, Target.row
    End If
End Sub



