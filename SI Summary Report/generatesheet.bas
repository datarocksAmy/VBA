Attribute VB_Name = "generatesheet"
Public Sub newsheet()
' Generate new sheet for each course

'Public sheet_name As String
Dim new_sheet As Worksheet
Dim WS As Worksheet: Set WS = ThisWorkbook.Sheets("Session-Grade")
Dim WB As Workbook

'sheet_name = InputBox(Prompt:="Enter sheet name: ", Default:="CHEM212")
If sheetExists(sheetMenu.TextBox1.Value) = True Then
    MsgBox ("Sheet name already exists. Please enter another name.")
    GoTo end_program
End If


If sheetMenu.TextBox1.Value = vbNullString Then
    MsgBox ("You just canceled!")
    GoTo end_program
End If

WS.Copy after:=Sheets(Sheets.Count): Set new_sheet = ActiveSheet: new_sheet.Name = sheetMenu.TextBox1.Value
ThisWorkbook.Sheets("Session-Grade").Activate

end_program:

End Sub
' Check if Sheet name already exist or not

Function sheetExists(sheetToFind As String) As Boolean
    sheetExists = False
    For Each Sheet In Worksheets
        If sheetToFind = Sheet.Name Then
            sheetExists = True
            Exit Function
        End If
    Next Sheet
End Function
