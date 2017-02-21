Attribute VB_Name = "AcademicType_DFW_Graph"
Sub countacademicType_DFW2()
' Count Occurance for Discipline Type
' DFW

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

lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("D" & num) = "BUS" Then
        BUS = BUS + 1
            
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HMED" Then
        HMED = HMED + 1
            
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HUM" Then
        HUM = HUM + 1
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "NS" Then
        NS = NS + 1
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "SS" Then
        SS = SS + 1
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "MATH" Then
        MATH = MATH + 1
            
    ElseIf Sheets("DFW Graph").Range("D" & num) = "COMP" Then
        COMP = COMP + 1
            
    ElseIf Sheets("DFW Graph").Range("D" & num) = "O" Then
        O = O + 1
    End If

Next num

Sheets("DFW Graph").Range("P55").Value = BUS
Sheets("DFW Graph").Range("P56").Value = HMED
Sheets("DFW Graph").Range("P57").Value = HUM
Sheets("DFW Graph").Range("P58").Value = NS
Sheets("DFW Graph").Range("P59").Value = SS
Sheets("DFW Graph").Range("P60").Value = MATH
Sheets("DFW Graph").Range("P61").Value = COMP
Sheets("DFW Graph").Range("P62").Value = O


End Sub

Sub sumSIgroup_DFW2()
' Sum up SI Group for each Discipline

Dim lastrow As Integer
Dim num As Integer
Dim BUS_SIG As Long
Dim HMED_SIG As Long
Dim HUM_SIG As Long
Dim NS_SIG As Long
Dim SS_SIG As Long
Dim MATH_SIG As Long
Dim COMP_SIG As Long
Dim O_SIG As Long


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("D" & num) = "BUS" Then
        BUS_SIG = BUS_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HMED" Then
        HMED_SIG = HMED_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HUM" Then
        HUM_SIG = HUM_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "NS" Then
        NS_SIG = NS_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "SS" Then
        SS_SIG = SS_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "MATH" Then
        MATH_SIG = MATH_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "COMP" Then
        COMP_SIG = COMP_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "O" Then
        O_SIG = O_SIG + Sheets("DFW Graph").Range("F" & num).Value
        
    End If

Next num

Sheets("DFW Graph").Range("Q55").Value = BUS_SIG
Sheets("DFW Graph").Range("Q56").Value = HMED_SIG
Sheets("DFW Graph").Range("Q57").Value = HUM_SIG
Sheets("DFW Graph").Range("Q58").Value = NS_SIG
Sheets("DFW Graph").Range("Q59").Value = SS_SIG
Sheets("DFW Graph").Range("Q60").Value = MATH_SIG
Sheets("DFW Graph").Range("Q61").Value = COMP_SIG
Sheets("DFW Graph").Range("Q62").Value = O_SIG


End Sub
Sub sumNonSIgroup_DFW2()
' Sum up NON-SI Group for each Discipline

Dim lastrow As Integer
Dim num As Integer
Dim BUS_NSIG As Long
Dim HMED_NSIG As Long
Dim HUM_NSIG As Long
Dim NS_NSIG As Long
Dim SS_NSIG As Long
Dim MATH_NSIG As Long
Dim COMP_NSIG As Long
Dim O_NSIG As Long


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("D" & num) = "BUS" Then
        BUS_NSIG = BUS_NSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HMED" Then
        HMED_NSIG = HMED_SNIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HUM" Then
        HUM_NSIG = HUM_NSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("" & num) = "NS" Then
        NS_NSIG = NS_NSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "SS" Then
        SS_NSIG = SS_NSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "MATH" Then
        MATH_NSIG = MATH_NSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "COMP" Then
        COMP_NSIG = COMP_NSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "O" Then
        O_NSIG = O_NSIG + Sheets("DFW Graph").Range("G" & num).Value

    End If

Next num

Sheets("DFW Graph").Range("R55").Value = BUS_NSIG
Sheets("DFW Graph").Range("R56").Value = HMED_NSIG
Sheets("DFW Graph").Range("R57").Value = HUM_NSIG
Sheets("DFW Graph").Range("R58").Value = NS_NSIG
Sheets("DFW Graph").Range("R59").Value = SS_NSIG
Sheets("DFW Graph").Range("R60").Value = MATH_NSIG
Sheets("DFW Graph").Range("R61").Value = COMP_NSIG
Sheets("DFW Graph").Range("R62").Value = O_NSIG



End Sub

Sub sumSIDFW_DFW2()
' Sum up SI DFW % for each Discipline

Dim lastrow As Integer
Dim num As Integer
Dim BUS_SIG As Variant
Dim HMED_SIG As Variant
Dim HUM_SIG As Variant
Dim NS_SIG As Variant
Dim SS_SIG As Variant
Dim MATH_SIG As Variant
Dim COMP_SIG As Variant
Dim O_SIG As Variant


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
On Error Resume Next
    If Sheets("DFW Graph").Range("D" & num) = "BUS" Then
        BUS_SIG = BUS_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HMED" Then
        HMED_SIG = HMED_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HUM" Then
        HUM_SIG = HUM_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "NS" Then
        NS_SIG = NS_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "SS" Then
        SS_SIG = SS_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "MATH" Then
        MATH_SIG = MATH_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "COMP" Then
        COMP_SIG = COMP_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "O" Then
        O_SIG = O_SIG + Sheets("DFW Graph").Range("J" & num).Value
        
    End If

Next num

Sheets("DFW Graph").Range("S55").Value = BUS_SIG
Sheets("DFW Graph").Range("S56").Value = HMED_SIG
Sheets("DFW Graph").Range("S57").Value = HUM_SIG
Sheets("DFW Graph").Range("S58").Value = NS_SIG
Sheets("DFW Graph").Range("S59").Value = SS_SIG
Sheets("DFW Graph").Range("S60").Value = MATH_SIG
Sheets("DFW Graph").Range("S61").Value = COMP_SIG
Sheets("DFW Graph").Range("S62").Value = O_SIG



End Sub

Sub sumNSIDFW_DFW2()
' Sum up NON-SI DFW % for each Discipline

Dim lastrow As Integer
Dim num As Integer
Dim BUS_NSIG As Variant
Dim HMED_NSIG As Variant
Dim HUM_NSIG As Variant
Dim NS_NSIG As Variant
Dim SS_NSIG As Variant
Dim MATH_NSIG As Variant
Dim COMP_NSIG As Variant
Dim O_NSIG As Variant


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
On Error Resume Next
    If Sheets("DFW Graph").Range("D" & num) = "BUS" Then
        BUS_NSIG = BUS_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HMED" Then
        HMED_NSIG = HMED_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "HUM" Then
        HUM_NSIG = HUM_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "NS" Then
        NS_NSIG = NS_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "SS" Then
        SS_NSIG = SS_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "MATH" Then
        MATH_NSIG = MATH_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "COMP" Then
        COMP_NSIG = COMP_NSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("D" & num) = "O" Then
        O_NSIG = O_NSIG + Sheets("DFW Graph").Range("K" & num).Value

    End If

Next num

Sheets("DFW Graph").Range("T55").Value = BUS_NSIG
Sheets("DFW Graph").Range("T56").Value = HMED_NSIG
Sheets("DFW Graph").Range("T57").Value = HUM_NSIG
Sheets("DFW Graph").Range("T58").Value = NS_NSIG
Sheets("DFW Graph").Range("T59").Value = SS_NSIG
Sheets("DFW Graph").Range("T60").Value = MATH_NSIG
Sheets("DFW Graph").Range("T61").Value = COMP_NSIG
Sheets("DFW Graph").Range("T62").Value = O_NSIG


End Sub



