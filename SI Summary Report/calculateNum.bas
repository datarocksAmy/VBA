Attribute VB_Name = "calculate"
Sub calcGrade()
' Calculate grade for each session
Dim lastrow As Integer, NSI_StartRow As Integer, NSI_EndRow As Integer
Dim all_data_lastrow As Integer
Dim SI_session As Integer
Dim session1to4 As Double, session5to9 As Double, session10plus As Double
Dim avg1to4 As Double, avg5to9 As Double, avg10plus As Double
Dim count1to4 As Integer, count5to9 As Integer, count10plus As Integer

all_data_lastrow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

' Sort Sessions
ActiveWorkbook.ActiveSheet.Sort.SortFields.Clear
ActiveWorkbook.ActiveSheet.Sort.SortFields.Add Key:=Range("B2:B" & all_data_lastrow) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

With ActiveWorkbook.ActiveSheet.Sort
        .SetRange Range("A1:C" & all_data_lastrow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With

' Calculate Grade
lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
NSI_StartRow = ActiveSheet.Range("B:B").Find(what:=0, after:=Range("B1")).Row
NSI_EndRow = ActiveSheet.Range("B:B").Find(what:=0, after:=Range("B1"), searchdirection:=xlPrevious).Row

For SI_session = (NSI_EndRow + 1) To lastrow

'SI Session 1-4
If ActiveSheet.Cells(SI_session, "B").Value >= 1 And ActiveSheet.Cells(SI_session, "B").Value < 5 Then
    If IsNumeric(ActiveSheet.Cells(SI_session, "C").Value) = True Then
        session1to4 = session1to4 + ActiveSheet.Cells(SI_session, "C").Value
        count1to4 = count1to4 + 1
        avg1to4 = session1to4 / count1to4
    End If
    
'SI Session 5-9
ElseIf ActiveSheet.Cells(SI_session, "B").Value >= 5 And ActiveSheet.Cells(SI_session, "B").Value < 10 Then
    If IsNumeric(ActiveSheet.Cells(SI_session, "C").Value) = True Then
    session5to9 = session5to9 + ActiveSheet.Cells(SI_session, "C").Value
    count5to9 = count5to9 + 1
    avg5to9 = session5to9 / count5to9
    End If
    
'SI Session 10+
ElseIf ActiveSheet.Cells(SI_session, "B").Value >= 10 Then
    If IsNumeric(ActiveSheet.Cells(SI_session, "C").Value) = True Then
    session10plus = session10plus + ActiveSheet.Cells(SI_session, "C").Value
    count10plus = count10plus + 1
    avg10plus = session10plus / count10plus
    End If
End If

Next SI_session


' Average grdade for each SI sessions
avg1to4 = session1to4 / count1to4
avg5to9 = session5to9 / count5to9
avg10plus = session10plus / count10plus

' Plug the numbers in
ActiveSheet.Range("G5").Value = Application.Average(Range("C" & NSI_StartRow & ":C" & NSI_EndRow))
ActiveSheet.Range("G6").Value = avg1to4
ActiveSheet.Range("G7").Value = avg5to9
ActiveSheet.Range("G8").Value = avg10plus



End Sub

Sub calcDFW()
' Count DFW occurances
Dim num As Integer, lastrow As Integer
Dim count1to4 As Integer, count5to9 As Integer, count10plus As Integer, NSIcount As Integer
Dim check As Integer
lastrow = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    'Non-SI Sessions
If ActiveSheet.Cells(num, "B").Value = 0 Then
    If ActiveSheet.Cells(num, "C").Value < 1.7 Or IsNumeric(ActiveSheet.Cells(num, "C").Value) = False Then
        NSIcount = NSIcount + 1
    End If
    
    'SI Session 1-4
ElseIf ActiveSheet.Cells(num, "B").Value >= 1 And ActiveSheet.Cells(num, "B").Value < 5 Then
    If ActiveSheet.Cells(num, "C").Value < 1.7 Or IsNumeric(ActiveSheet.Cells(num, "C").Value) = False Then
        count1to4 = count1to4t + 1
    End If
    
    'SI Session 5-9
ElseIf ActiveSheet.Cells(num, "B").Value >= 5 And ActiveSheet.Cells(num, "B").Value < 10 Then
    If ActiveSheet.Cells(num, "C").Value < 1.7 Or IsNumeric(ActiveSheet.Cells(num, "C").Value) = False Then
        count5to9 = count5to9 + 1
    End If
    
    'SI Session 10+
ElseIf ActiveSheet.Cells(num, "B").Value >= 10 Then
    If ActiveSheet.Cells(num, "C").Value < 1.7 Or IsNumeric(ActiveSheet.Cells(num, "C").Value) = False Then
        count10plus = count10plus + 1
    End If
End If

Next

' Plug in DFW counts for each sessions
ActiveSheet.Range("H5").Value = NSIcount
ActiveSheet.Range("H6").Value = count1to4t
ActiveSheet.Range("H7").Value = count5to9
ActiveSheet.Range("H8").Value = count10plus

' Display default value 0 if there's no value
For check = 5 To 8

If (ActiveSheet.Cells(check, "H").Value) = 0 Then
    ActiveSheet.Range("H" & check).Value = 0
End If

Next check

End Sub
