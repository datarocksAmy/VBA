Attribute VB_Name = "AcademicType_Pie_Graph"
Sub countacademicType_Pie()
' Count Occurance for Discipline Type
' Pie Graph

Dim lastrow As Integer
Dim num As Integer

Dim BUS As Long
Dim HMED As Long
Dim HUM As Long
Dim NS As Long
Dim SS As Long
Dim MATH As Long
Dim COMP As Long
Dim O As Long

lastrow = Sheets("Pie Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow

    If Sheets("Pie Graph").Range("D" & num) = "BUS" Then
        BUS = BUS + 1
            
    ElseIf Sheets("Pie Graph").Range("D" & num) = "HMED" Then
        HMED = HMED + 1
            
    ElseIf Sheets("Pie Graph").Range("D" & num) = "HUM" Then
        HUM = HUM + 1
        
    ElseIf Sheets("Pie Graph").Range("D" & num) = "NS" Then
        NS = NS + 1
        
    ElseIf Sheets("Pie Graph").Range("D" & num) = "SS" Then
        SS = SS + 1
        
    ElseIf Sheets("Pie Graph").Range("D" & num) = "MATH" Then
        MATH = MATH + 1
            
    ElseIf Sheets("Pie Graph").Range("D" & num) = "COMP" Then
        COMP = COMP + 1
            
    ElseIf Sheets("Pie Graph").Range("D" & num) = "O" Then
        O = O + 1
    End If

Next num

Sheets("Pie Graph").Range("L3").Value = BUS
Sheets("Pie Graph").Range("L4").Value = HMED
Sheets("Pie Graph").Range("L5").Value = HUM
Sheets("Pie Graph").Range("L6").Value = NS
Sheets("Pie Graph").Range("L7").Value = SS
Sheets("Pie Graph").Range("L8").Value = MATH
Sheets("Pie Graph").Range("L9").Value = COMP
Sheets("Pie Graph").Range("L10").Value = O


End Sub
