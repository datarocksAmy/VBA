Attribute VB_Name = "InstitutionFormatModule"
Sub InstitutionFormat()
' Format Institution Type from numbers into 4PR/4PU/2PR/2PR

Dim l As Long
Dim s As String
Dim d As String
Dim lastrow As Integer
Dim m As Integer


'Start from the first row of the information
Range("C42").Select
lastrow = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row
l = Range("C42").Select

'Change the value until the last row
For m = 42 To lastrow

'Skip if the Cell value is empty or a string(already formatted)
If ActiveSheet.Cells(m, "C").Value = " " Or Not IsNumeric(ActiveSheet.Cells(m, "C")) Then GoTo skip
    l = ActiveCell.Value
    s = CStr(l)
    
    If (Len(s) > 0) Then

        If (l = 1) Then
         d = Range("B2").Value
        
        ElseIf (l = 2) Then
         d = Range("B3").Value
         
        ElseIf (l = 3) Then
         d = Range("B4").Value
        
        ElseIf (l = 4) Then
         d = Range("B5").Value
         
        Else
         d = l
     
End If
' Print the value back to the column with formatted value
ActiveCell.Value = d

Else
ActiveCell.Value = s

End If


skip:
ActiveCell.Offset(1, 0).Select
Next m


End Sub

