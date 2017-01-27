Attribute VB_Name = "Module1"
Sub loan()

    choice = 6
    Do While choice = 6
      
      Columns("A:D").Select                                                       'ActiveCell.Range("A2:D2").Resize(1000, 4).Select
      Selection.ClearContents
      
      Range("A1").Select
      ActiveCell.Value = " MONTH "
      ActiveCell.Offset(0, 1).Select
      ActiveCell.Value = " INTEREST "
      ActiveCell.Offset(0, 1).Select
      ActiveCell.Value = " TOTAL INTEREST "
      ActiveCell.Offset(0, 1).Select
      ActiveCell.Value = " BALANCE "
      
      
      '<< Gather Input Data >>
       P = Val(InputBox(" Enter the Loan Principle: "))
       d = Val(InputBox(" Enter the Loan Duration in months: "))
       yrR = Val(InputBox(" Enter the Yearly Interest Rate (%)"))
    
      
      '<< Initial Calculation >>
       r = ((yrR / 12) / 100)
       payment = P * (r + (r / ((1 + r) ^ d - 1)))
       payment = Round(payment, 4)
       B = P
       itotal = 0
      
      '<< Monthly Calculations >>
       Range("A2").Select
       For m = 1 To d
       
       imonth = r * B    'calculate the interest for that month
       B = B + imonth - payment
       itotal = itotal + imonth
       itotal = Round(itotal, 4)
       
       '<< Display Results >>
       
            
       ActiveCell.Value = m
       ActiveCell.Offset(0, 1).Select
    
       ActiveCell.Value = imonth
       ActiveCell.Offset(0, 1).Select
    
       ActiveCell.Value = itotal
       ActiveCell.Offset(0, 1).Select
       
       ActiveCell.Value = B
       
       ActiveCell.Offset(1, -3).Select
    
       Next
       
       
       
       MsgBox (" Monthly Payment= $ " & payment & Chr(13) & " Total Interest= $ " & itotal & Chr(13) & " Total Expense of Loan= $ " & P + itotal)
    
      choice = MsgBox(" Run again? ", vbYesNo, " Continue ")
      
    Loop


End Sub
