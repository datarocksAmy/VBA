Attribute VB_Name = "Beam2"
Sub beam2()

'Input Section

Range("B1").Select

b = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

h = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

maxdef = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

Esteel = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

Ealuminum = ActiveCell.Value
ActiveCell.Offset(1, 0).Select

F = ActiveCell.Value




'Calculation Section

Call calc(b, h, maxdef, E, F, I, Esteel, Ealuminum, maxLsteel, maxLaluminum)




'Output Section

MsgBox " Required beam length for steel is " & (maxLsteel) & " inches " & vbNewLine & " Required beam length for aluminum is " & (maxLaluminum) & " inches "


End Sub

Sub calc(b, h, maxdef, E, F, I, Esteel, Ealuminum, maxLsteel, maxLaluminum)

I = (b * (h ^ 3)) / 12
maxLsteel = Exp((1 / 3) * Log((3 * Esteel * I) / F))
maxLaluminum = Exp((1 / 3) * Log((3 * Ealuminum * I) / F))

End Sub

