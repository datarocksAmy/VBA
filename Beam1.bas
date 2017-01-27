Attribute VB_Name = "Beam1"
Sub beam1()

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

I = (b * h ^ 3) / 12
maxLsteel = Exp((1 / 3) * Log((3 * Esteel * I) / F))
maxLaluminum = Exp((1 / 3) * Log((3 * Ealuminum * I) / F))





'Output Selection

Range("B8").Select

ActiveCell.Value = I
ActiveCell.Offset(1, 0).Select

ActiveCell.Value = maxLsteel
ActiveCell.Offset(1, 0).Select

ActiveCell.Value = maxLaluminum






End Sub


