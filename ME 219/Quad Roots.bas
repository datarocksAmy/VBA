Attribute VB_Name = "Module1"
Option Explicit

Sub quadroots()

Dim runagain As Integer
runagain = 6
Do While runagain = 6

'Define variables
Dim a As Double, x As Double, b As Double, c As Double, r As Double, d As Double, r1 As String, r2 As String
Dim r11, r12 As Double
Dim naturesolution, rootanswers As String


'Input Section
'equation f(x)=ax^2+bx+c

a = InputBox(" Enter coefficient a: ")
b = InputBox(" Enter coefficient b: ")
c = InputBox(" Enter coefficient c: ")


'Nested Decision Structure


If a = 0 Then

    If b = 0 Then
       rootanswers = (" N/A ")
       naturesolution = (" Trivial Solution ")
       
    Else  'b<>0
       r = -(c / b)
       r = Round(r, 4)
       
       rootanswers = r 'Single root
       naturesolution = (" Single Root ")
    
    End If



Else   'a<>0
 d = (b ^ 2) - 4 * a * c
    
    If d > 0 Then
       
       r1 = (-b + d ^ (1 / 2)) / (2 * a)
       r2 = (-b - d ^ (1 / 2)) / (2 * a)
       
       r1 = Round(r1, 4)
       r2 = Round(r2, 4)
       
       rootanswers = (vbNewLine & " r1= " & r1 & vbNewLine & " r2= " & r2)
       naturesolution = (" Two Distinct Real Roots ")
       ' r is Two distinct Real Roots
       
    ElseIf d = 0 Then
        
        r = (-b) / 2 * a
        r = Round(r, 4)
        rootanswers = r
        naturesolution = (" (Identical) Single Root ")
        ' r is Double Roots
        
        
    Else  'd<0.
       
       r11 = (-b / 2 * a)
       r12 = (((Abs(d)) ^ (1 / 2)) / (2 * a))
       
       r1 = (Round(r11, 4) & " + " & Round(r12, 4) & " i ")
       r2 = (Round(r11, 4) & " - " & Round(r12, 4) & " i ")
       
       rootanswers = (vbNewLine & " r1= " & r1 & vbNewLine & " r2= " & r2)
       naturesolution = (" Complex Roots ")
       
       
       
    End If


End If



'Output Section

runagain = MsgBox(" Nature of the Solution is << " & naturesolution & " >>." & vbNewLine & " Answers of the root(s): " & rootanswers & vbNewLine & vbNewLine & " Do you want to try again? ", vbYesNo + vbInformation, " RESULT WINDOW ")


Loop

runagain = MsgBox(" See you next time! :) ", vbNo, " EXIT ")


End Sub

