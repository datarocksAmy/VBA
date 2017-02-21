Attribute VB_Name = "Module1"
Option Explicit

Public Const g = 980 '(cm/s^2)


Sub stokes()

Dim heightleft As Integer
Dim PP As Double, PF As Double, Ho As Double, H As Double, mu As Double, d As Double, v As Double, t As Double


'Input Section

Range("B1").Select

PP = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

PF = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

mu = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

d = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

Ho = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

H = ActiveCell.Value




'Calculation Section

Call velcalc(v, PP, PF, mu, d, H, Ho, g, heightleft, t)


'Output Section

Range("B7").Select

Selection.Font.Bold = True
Selection.Font.Italic = True
With Selection.Font
              .Color = -16776961
              .TintAndShade = 0
   
End With

ActiveCell.Value = v


If v > 0 Then
MsgBox (" The action of particle: SINK TO THE BOTTOM. " & vbNewLine & " Amount of time is " & t & " seconds. ")

Else
MsgBox (" The action of particle: FLOAT TO THE SURFACE. " & vbNewLine & " Amount of time is " & -t & " seconds. ")

End If


End Sub

Sub velcalc(v, PP, PF, mu, d, H, Ho, g, heightleft, t)

'v = (g / 18) * (((PP - PF)) / mu) * (d ^ 2)
v = (g * (PP - PF) * d ^ 2) / (18 * mu)
'MsgBox (v)


heightleft = H - Ho
t = heightleft / v


End Sub
