Attribute VB_Name = "AcademicFormatModule"
Sub AcademicFormat()
'Format Academic Discipline Key into Abbreviations

Dim l As Long
Dim s As String
Dim d As String
Dim lastrow As Integer
Dim m As Integer


'Start from the first row of the information
Range("D42").Select
lastrow = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row
l = Range("D42").Select

'Change the value until the last row
For m = 42 To lastrow
'Skip if the Cell value is empty or a string(already formatted)

If ActiveSheet.Cells(m, "D").Value = " " Or Not IsNumeric(ActiveSheet.Cells(m, "D")) Then GoTo skip
 l = ActiveCell.Value
 s = CStr(l)
    If (Len(s) > 0) Then
        If (l = 1) Then
         d = Range("C10").Value
        
        ElseIf (l = 2) Then
         d = Range("C11").Value
         
        ElseIf (l = 3) Then
         d = Range("C12").Value
        
        ElseIf (l = 4) Then
         d = Range("C13").Value
         
        ElseIf (l = 5) Then
         d = Range("C14").Value
         
        ElseIf (l = 6) Then
         d = Range("C15").Value
         
        ElseIf (l = 7) Then
         d = Range("C16").Value
         
        ElseIf (l = 8) Then
         d = Range("C17").Value

    End If
    
ActiveCell.Value = d

Else
ActiveCell.Value = s

End If

skip:


ActiveCell.Offset(1, 0).Select
Next m


End Sub

