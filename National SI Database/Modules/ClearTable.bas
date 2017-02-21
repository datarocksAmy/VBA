Attribute VB_Name = "ClearTable"
Sub clearTable_GPA()
' Clear the table content in Sheet("GPA Graph")
Sheets("GPA Graph").Range("P3:T6").ClearContents
Sheets("GPA Graph").Range("P55:T62").ClearContents

End Sub

Sub clearTable_DFW()
' Clear the table content in Sheet("DFW Graph")
Sheets("DFW Graph").Range("P3:T6").ClearContents
Sheets("DFW Graph").Range("P55:T62").ClearContents


End Sub

Sub clearTable_Pie()
' Clear the table content in Sheet("Pie Graph")

Sheets("Pie Graph").Range("L3:L10").ClearContents

End Sub
