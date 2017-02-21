Attribute VB_Name = "CopyAndPaste_Graph"
Sub copyandpasteGPA()

Dim lastrow As Integer
Dim num As Integer

'Obtain data from Sheet "Data Clean"
lastrow = Sheets("Data Clean").Cells(Rows.Count, "B").End(xlUp).Row

'Clear the data on the sheet before pasting the updated version
Sheets("GPA Graph").Range("A1:I100000").ClearContents
Sheets("GPA Graph").Range("A1:I100000").Clear


For num = 1 To lastrow
Sheets("GPA Graph").Range("A" & num).Value = Sheets("Data Clean").Range("A" & num).Value
Sheets("GPA Graph").Range("B" & num).Value = Sheets("Data Clean").Range("B" & num).Value
Sheets("GPA Graph").Range("C" & num).Value = Sheets("Data Clean").Range("C" & num).Value
Sheets("GPA Graph").Range("D" & num).Value = Sheets("Data Clean").Range("D" & num).Value
Sheets("GPA Graph").Range("E" & num).Value = Sheets("Data Clean").Range("E" & num).Value
Sheets("GPA Graph").Range("F" & num).Value = Sheets("Data Clean").Range("F" & num).Value
Sheets("GPA Graph").Range("G" & num).Value = Sheets("Data Clean").Range("G" & num).Value
Sheets("GPA Graph").Range("H" & num).Value = Sheets("Data Clean").Range("H" & num).Value
Sheets("GPA Graph").Range("I" & num).Value = Sheets("Data Clean").Range("I" & num).Value

Next num

Columns("A:I").HorizontalAlignment = xlCenter


End Sub

Sub copyandpasteDFW()

Dim lastrow As Integer
Dim num As Integer

'Obtain data from Sheet "Data Clean"
lastrow = Sheets("Data Clean").Cells(Rows.Count, "B").End(xlUp).Row

'Clear the data on the sheet before pasting the updated version
Sheets("DFW Graph").Range("A1:I100000").ClearContents
Sheets("DFW Graph").Range("A1:I100000").Clear


For num = 1 To lastrow
Sheets("DFW Graph").Range("A" & num).Value = Sheets("Data Clean").Range("A" & num).Value
Sheets("DFW Graph").Range("B" & num).Value = Sheets("Data Clean").Range("B" & num).Value
Sheets("DFW Graph").Range("C" & num).Value = Sheets("Data Clean").Range("C" & num).Value
Sheets("DFW Graph").Range("D" & num).Value = Sheets("Data Clean").Range("D" & num).Value
Sheets("DFW Graph").Range("E" & num).Value = Sheets("Data Clean").Range("E" & num).Value
Sheets("DFW Graph").Range("F" & num).Value = Sheets("Data Clean").Range("F" & num).Value
Sheets("DFW Graph").Range("G" & num).Value = Sheets("Data Clean").Range("G" & num).Value
Sheets("DFW Graph").Range("H" & num).Value = Sheets("Data Clean").Range("J" & num).Value
Sheets("DFW Graph").Range("I" & num).Value = Sheets("Data Clean").Range("K" & num).Value
'Sheets("DFW Graph").Range("J" & num).Value = Sheets("Data Clean").Range("L" & num).Value

Next num

Columns("A:I").HorizontalAlignment = xlCenter


End Sub

'Sort by Institution Tpye: perform GPA and DFW for SI and non-SI group
'Sort by Discipline: perform GPA and DFW for SI and non-SI group

Sub copyandpastePie()

Dim lastrow As Integer
Dim num As Integer

'Obtain data from Sheet "Data Clean"
lastrow = Sheets("Data Clean").Cells(Rows.Count, "B").End(xlUp).Row

'Clear the data on the sheet before pasting the updated version
Sheets("Pie Graph").Range("A1:F100000").ClearContents
Sheets("Pie Graph").Range("A1:F100000").Clear


For num = 1 To lastrow
Sheets("Pie Graph").Range("A" & num).Value = Sheets("Data Clean").Range("A" & num).Value
Sheets("Pie Graph").Range("B" & num).Value = Sheets("Data Clean").Range("B" & num).Value
Sheets("Pie Graph").Range("C" & num).Value = Sheets("Data Clean").Range("C" & num).Value
Sheets("Pie Graph").Range("D" & num).Value = Sheets("Data Clean").Range("D" & num).Value
Sheets("Pie Graph").Range("E" & num).Value = Sheets("Data Clean").Range("E" & num).Value
'Sheets("Pie Graph").Range("F" & num).Value = Sheets("Data Clean").Range("F" & num).Value


Next num

Columns("A:E").HorizontalAlignment = xlCenter


End Sub
