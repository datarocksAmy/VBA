Attribute VB_Name = "InstType_DFW_Graph"
Sub countInstType_DFW()
' Count occurance for 2PR/2PU/4PR/4PU
' DFW

Dim lastrow As Integer
Dim num As Integer
Dim twoprivate As Long
Dim twopublic As Long
Dim fourprivate As Long
Dim fourpublic As Long


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow

    If Sheets("DFW Graph").Range("C" & num) = "2PR" Then
        twoprivate = twoprivate + 1
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "2PU" Then
        twopublic = twopublic + 1
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PR" Then
        fourprivate = fourprivate + 1
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PU" Then
        fourpublic = fourpublic + 1
        
    End If

Next num

Sheets("DFW Graph").Range("P3").Value = twoprivate
Sheets("DFW Graph").Range("P4").Value = twopublic
Sheets("DFW Graph").Range("P5").Value = fourprivate
Sheets("DFW Graph").Range("P6").Value = fourpublic


End Sub
Sub sumSIgroup_DFW()
' Sum up SI Group for each Institution Type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateSIG As Long
Dim twopublicSIG As Long
Dim fourprivateSIG As Long
Dim fourpublicSIG As Long


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row
fourpublicSIG = 0
For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("C" & num) = "2PR" Then
        twoprivateSIG = twoprivateSIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "2PU" Then
        twopublicSIG = twopublicSIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PR" Then
        fourprivateSIG = fourprivateSIG + Sheets("DFW Graph").Range("F" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PU" Then
        fourpublicSIG = fourpublicSIG + Sheets("DFW Graph").Range("F" & num).Value

    End If

Next num

Sheets("DFW Graph").Range("Q3").Value = twoprivateSIG
Sheets("DFW Graph").Range("Q4").Value = twopublicSIG
Sheets("DFW Graph").Range("Q5").Value = fourprivateSIG
Sheets("DFW Graph").Range("Q6").Value = fourpublicSIG


End Sub
Sub sumNonSIgroup_DFW()
' Sum up NON-SI Group for each Institution Type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateNSIG As Long
Dim twopublicNSIG As Long
Dim fourprivateNSIG As Long
Dim fourpublicNSIG As Long


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("C" & num) = "2PR" Then
        twoprivateNSIG = twoprivateNSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "2PU" Then
        twopublicNSIG = twopublicNSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PR" Then
        fourprivateNSIG = fourprivateNSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PU" Then
        fourpublicNSIG = fourpublicNSIG + Sheets("DFW Graph").Range("G" & num).Value
        
    End If

Next num

Sheets("DFW Graph").Range("R3").Value = twoprivateNSIG
Sheets("DFW Graph").Range("R4").Value = twopublicNSIG
Sheets("DFW Graph").Range("R5").Value = fourprivateNSIG
Sheets("DFW Graph").Range("R6").Value = fourpublicNSIG


End Sub

Sub sumSIDFW_DFW()
' Sum up SI DFW % for each Institution type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateSIG As Variant
Dim twopublicSIG As Variant
Dim fourprivateSIG As Variant
Dim fourpublicSIG As Variant


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("C" & num) = "2PR" Then
        twoprivateSIG = twoprivateSIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "2PU" Then
        twopublicSIG = twopublicSIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PR" Then
        fourprivateSIG = fourprivateSIG + Sheets("DFW Graph").Range("J" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PU" Then
        fourpublicSIG = fourpublicSIG + Sheets("DFW Graph").Range("J" & num).Value

    End If

Next num

Sheets("DFW Graph").Range("S3").Value = twoprivateSIG
Sheets("DFW Graph").Range("S4").Value = twopublicSIG
Sheets("DFW Graph").Range("S5").Value = fourprivateSIG
Sheets("DFW Graph").Range("S6").Value = fourpublicSIG


End Sub
Sub sumNSIDFW_DFW()
' Sum up NON-SI DFW % for each Institution type

Dim lastrow As Integer
Dim num As Integer
Dim twoprivateNSIG As Variant
Dim twopublicNSIG As Variant
Dim fourprivateNSIG As Variant
Dim fourpublicNSIG As Variant


lastrow = Sheets("DFW Graph").Cells(Rows.Count, "B").End(xlUp).Row

For num = 2 To lastrow
    On Error Resume Next
    If Sheets("DFW Graph").Range("C" & num) = "2PR" Then
        twoprivateNSIG = twoprivateNSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "2PU" Then
        twopublicNSIG = twopublicNSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PR" Then
        fourprivateNSIG = fourprivateNSIG + Sheets("DFW Graph").Range("K" & num).Value
        
    ElseIf Sheets("DFW Graph").Range("C" & num) = "4PU" Then
        fourpublicNSIG = fourpublicNSIG + Sheets("DFW Graph").Range("K" & num).Value
    End If

Next num

Sheets("DFW Graph").Range("T3").Value = twoprivateNSIG
Sheets("DFW Graph").Range("T4").Value = twopublicNSIG
Sheets("DFW Graph").Range("T5").Value = fourprivateNSIG
Sheets("DFW Graph").Range("T6").Value = fourpublicNSIG


End Sub



