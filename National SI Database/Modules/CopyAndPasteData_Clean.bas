Attribute VB_Name = "CopyAndPasteData_Clean"
Sub DataCleanCopyPaste()
'Copy the raw data and paste in sheet "Data Clean" for further modify

Dim lastrow As Integer
Dim num As Integer
Dim num_D As Integer

'Obtain data from Sheet "Data"
lastrow = Sheets("Raw Data").Cells(Rows.Count, "A").End(xlUp).Row
lastrow_clean = Sheets("Data Clean").Cells(Rows.Count, "A").End(xlUp).Row

'Clear the data on the sheet before pasting the updated version
Sheets("Data Clean").Range("A3724:H100000").ClearContents
Sheets("Data Clean").Range("A3724:H100000").Clear


For num = 3723 To lastrow
Sheets("Data Clean").Range("A" & num).Value = Sheets("Raw Data").Range("A" & (num + 2162)).Value
Sheets("Data Clean").Range("B" & num).Value = Sheets("Raw Data").Range("B" & (num + 2162)).Value
Sheets("Data Clean").Range("C" & num).Value = Sheets("Raw Data").Range("C" & (num + 2162)).Value
Sheets("Data Clean").Range("D" & num).Value = Sheets("Raw Data").Range("D" & (num + 2162)).Value
Sheets("Data Clean").Range("E" & num).Value = Sheets("Raw Data").Range("E" & (num + 2162)).Value
Sheets("Data Clean").Range("F" & num).Value = Sheets("Raw Data").Range("F" & (num + 2162)).Value
Sheets("Data Clean").Range("G" & num).Value = Sheets("Raw Data").Range("G" & (num + 2162)).Value
Sheets("Data Clean").Range("H" & num).Value = Sheets("Raw Data").Range("H" & (num + 2162)).Value
Sheets("Data Clean").Range("I" & num).Value = Sheets("Raw Data").Range("I" & (num + 2162)).Value
Sheets("Data Clean").Range("J" & num).Value = Sheets("Raw Data").Range("J" & (num + 2162)).Value
Sheets("Data Clean").Range("K" & num).Value = Sheets("Raw Data").Range("K" & (num + 2162)).Value
Sheets("Data Clean").Range("L" & num).Value = Sheets("Raw Data").Range("L" & (num + 2162)).Value
Sheets("Data Clean").Range("M" & num).Value = Sheets("Raw Data").Range("M" & (num + 2162)).Value

Next num



Columns("A:M").HorizontalAlignment = xlCenter


End Sub

'Sort by Institution Tpye: perform GPA and DFW for SI and non-SI group
'Sort by Discipline: perform GPA and DFW for SI and non-SI group

