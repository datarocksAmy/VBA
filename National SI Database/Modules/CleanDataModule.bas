Attribute VB_Name = "CleanDataModule"
Sub cleanMissingData()
' Clean out any missing data and delete the whole row

Dim colchar As String
Dim lastrow As Integer
Dim m As Integer
Dim col As Integer

Range("C2").Select
lastrow = ActiveSheet.Cells(Rows.Count, "C").End(xlUp).Row

For m = 2 To lastrow
On Error Resume Next
    If Not IsNumeric(Sheets("Data Clean").Cells(m, "B").Value) Or Len(Sheets("Data Clean").Cells(m, "B").Value) = 0 Then
        Columns("B").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        Columns("D").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
        Else
        For col = 5 To 11
        colchar = ConvertToLetter(col)
         If Len(Sheets("Data Clean").Cells(m, colchar).Value) = 0 Then
                Columns(colchar).SpecialCells(xlCellTypeBlanks).EntireRow.Delete
         ElseIf IsNumeric(Sheets("Data Clean").Cells(m, colchar)) = False Or Sheets("Data Clean").Cells(m, colchar) = "na" Or Sheets("Data Clean").Cells(m, colchar) = "NA" Then
                Rows(m).EntireRow.Delete
         Else
         On Error Resume Next
         End If
         
        Next col

    End If
    
Next m


End Sub

Sub deleteBlankRows()
On Error Resume Next
Columns("D").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
End Sub
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
