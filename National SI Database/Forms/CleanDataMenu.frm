VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CleanDataMenu 
   Caption         =   "Clean Data Menu"
   ClientHeight    =   5580
   ClientLeft      =   48
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "CleanDataMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CleanDataMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton3_Click()

End Sub

Private Sub CleanDataMenu_Click()
cleanMissingData
End Sub

Private Sub CommandButton10_Click()
unmarkColor
End Sub

Private Sub CommandButton4_Click()
DeleteOldUpdateNew
End Sub

Private Sub CommandButton8_Click()
deleteBlankRows
End Sub

Private Sub CommandButton9_Click()
MarkZeroColor
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub
