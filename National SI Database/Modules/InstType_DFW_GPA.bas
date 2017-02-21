Attribute VB_Name = "instType_DFW_GPA"
Sub countInstType()
' Count occurance for 2PR/2PU/4PR/4PU
' GPA

Dim lastrow As Integer
Dim num As Integer
Dim twoprivate As Long
Dim twopublic As Long
Dim fourprivate As Long
Dim fourpublic As Long


lastrow = Sheets("GPA Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow

    If Sheets("GPA Graph").Range("C" & num) = "2PR" Then
        twoprivate = twoprivate + 1
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "2PU" Then
        twopublic = twopublic + 1
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PR" Then
        fourprivate = fourprivate + 1
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PU" Then
        fourpublic = fourpublic + 1
        
    End If

Next num
Sheets("GPA Graph").Range("P3").Value = twoprivate
Sheets("GPA Graph").Range("P4").Value = twopublic
Sheets("GPA Graph").Range("P5").Value = fourprivate
Sheets("GPA Graph").Range("P6").Value = fourpublic


End Sub
Sub sumSIgroup()
' Sum up SI Group for each Institution Type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateSIG As Long
Dim twopublicSIG As Long
Dim fourprivateSIG As Long
Dim fourpublicSIG As Long


lastrow = Sheets("GPA Graph").Cells(Rows.Count, "B").End(xlUp).Row
fourpublicSIG = 0
For num = 2 To lastrow

    If Sheets("GPA Graph").Range("C" & num) = "2PR" Then
        twoprivateSIG = twoprivateSIG + Sheets("GPA Graph").Range("F" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "2PU" Then
        twopublicSIG = twopublicSIG + Sheets("GPA Graph").Range("F" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PR" Then
        fourprivateSIG = fourprivateSIG + Sheets("GPA Graph").Range("F" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PU" Then
        fourpublicSIG = fourpublicSIG + Sheets("GPA Graph").Range("F" & num).Value
        
    Else
    On Error Resume Next
    End If

Next num
Sheets("GPA Graph").Range("Q3").Value = twoprivateSIG
Sheets("GPA Graph").Range("Q4").Value = twopublicSIG
Sheets("GPA Graph").Range("Q5").Value = fourprivateSIG
Sheets("GPA Graph").Range("Q6").Value = fourpublicSIG


End Sub
Sub sumNonSIgroup()
' Sum up NON-SI Group for each Institution Type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateNSIG As Long
Dim twopublicNSIG As Long
Dim fourprivateNSIG As Long
Dim fourpublicNSIG As Long


lastrow = Sheets("GPA Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow

    If Sheets("GPA Graph").Range("C" & num) = "2PR" Then
        twoprivateNSIG = twoprivateNSIG + Sheets("GPA Graph").Range("G" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "2PU" Then
        twopublicNSIG = twopublicNSIG + Sheets("GPA Graph").Range("G" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PR" Then
        fourprivateNSIG = fourprivateNSIG + Sheets("GPA Graph").Range("G" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PU" Then
        fourpublicNSIG = fourpublicNSIG + Sheets("GPA Graph").Range("G" & num).Value
        
    Else
    On Error Resume Next
    End If

Next num
Sheets("GPA Graph").Range("R3").Value = twoprivateNSIG
Sheets("GPA Graph").Range("R4").Value = twopublicNSIG
Sheets("GPA Graph").Range("R5").Value = fourprivateNSIG
Sheets("GPA Graph").Range("R6").Value = fourpublicNSIG


End Sub

Sub sumSIGPA()
' Sum up SI GPA for each Institution type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateSIG As Variant
Dim twopublicSIG As Variant
Dim fourprivateSIG As Variant
Dim fourpublicSIG As Variant


lastrow = Sheets("GPA Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow

    If Sheets("GPA Graph").Range("C" & num) = "2PR" Then
        twoprivateSIG = twoprivateSIG + Sheets("GPA Graph").Range("H" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "2PU" Then
        twopublicSIG = twopublicSIG + Sheets("GPA Graph").Range("H" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PR" Then
        fourprivateSIG = fourprivateSIG + Sheets("GPA Graph").Range("H" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PU" Then
        fourpublicSIG = fourpublicSIG + Sheets("GPA Graph").Range("H" & num).Value
        
    Else
    On Error Resume Next
    End If

Next num
Sheets("GPA Graph").Range("S3").Value = twoprivateSIG
Sheets("GPA Graph").Range("S4").Value = twopublicSIG
Sheets("GPA Graph").Range("S5").Value = fourprivateSIG
Sheets("GPA Graph").Range("S6").Value = fourpublicSIG


End Sub
Sub sumNSIGPA()
' Sum up NON-SI GPA for each Institution type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateNSIG As Variant
Dim twopublicNSIG As Variant
Dim fourprivateNSIG As Variant
Dim fourpublicNSIG As Variant


lastrow = Sheets("GPA Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow

    If Sheets("GPA Graph").Range("C" & num) = "2PR" Then
        twoprivateNSIG = twoprivateNSIG + Sheets("GPA Graph").Range("I" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "2PU" Then
        twopublicNSIG = twopublicNSIG + Sheets("GPA Graph").Range("I" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PR" Then
        fourprivateNSIG = fourprivateNSIG + Sheets("GPA Graph").Range("I" & num).Value
        
    ElseIf Sheets("GPA Graph").Range("C" & num) = "4PU" Then
        fourpublicNSIG = fourpublicNSIG + Sheets("GPA Graph").Range("H" & num).Value
        
    Else
    On Error Resume Next
    End If

Next num
Sheets("GPA Graph").Range("T3").Value = twoprivateNSIG
Sheets("GPA Graph").Range("T4").Value = twopublicNSIG
Sheets("GPA Graph").Range("T5").Value = fourprivateNSIG
Sheets("GPA Graph").Range("T6").Value = fourpublicNSIG



End Sub


