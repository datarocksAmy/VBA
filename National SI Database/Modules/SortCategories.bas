Attribute VB_Name = "SortCategories"
Sub sortInstType()
'Sort the sheet by Institution Type

Dim lastrow As Integer

lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("B2:B" & lastrow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:I" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With


End Sub
Sub sortAcademicDis()
'Sort the sheet by Academic Discipline

Dim lastrow As Integer
Dim m As Integer

lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("C2:C" & lastrow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:I" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With


End Sub
Sub sortAcademicDis_Pie()
'Sort the sheet by Academic Discipline for Pie Graph

Dim lastrow As Integer
Dim m As Integer

lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("C2:C" & lastrow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:F" & lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With


End Sub

