Attribute VB_Name = "MarkZero"
Sub MarkZeroColor()
' Mark row with 0 SI GPA, N-SI GPA, SI-DFW, N-SI DFW

Dim colchar As String
Dim lastrow As Integer
Dim m As Integer
Dim col As Integer

Range("B2").Select
lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row

For m = 2 To lastrow
On Error Resume Next
    For col = 7 To 10
    colchar = ConvertToLetter(col)
    If ActiveSheet.Cells(m, col).Value = 0 Then
        Cells(m, col).EntireRow.Interior.ColorIndex = 36
    End If
         
    Next col
Next m

End Sub
Sub unmarkColor()
' Unmark Color area
Dim colchar As String
Dim lastrow As Integer
Dim m As Integer
Dim col As Integer

Range("B2").Select
lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row

For m = 2 To lastrow
On Error Resume Next
    For col = 7 To 10
    colchar = ConvertToLetter(col)
    If ActiveSheet.Cells(m, col).Value = 0 Then
        Cells(m, col).EntireRow.Interior.ColorIndex = 0
    End If
         
    Next col
Next m

End Sub
' Function to Convert numbers to letter(character)
Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

