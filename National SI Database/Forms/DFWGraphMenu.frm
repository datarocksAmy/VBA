VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DFWGraphMenu 
   Caption         =   "DFW Graph Menu"
   ClientHeight    =   5562
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   5670
   OleObjectBlob   =   "DFWGraphMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DFWGraphMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
sortInstType
End Sub

Private Sub CommandButton3_Click()
sortAcademicDis
End Sub

Private Sub CommandButton4_Click()
copyandpasteDFW
End Sub

Private Sub CommandButton7_Click()
countInstType_DFW
sumSIgroup_DFW
sumNonSIgroup_DFW
sumSIDFW_DFW
sumNSIDFW_DFW
End Sub

Private Sub CommandButton8_Click()
countacademicType_DFW2
sumSIgroup_DFW2
sumNonSIgroup_DFW2
sumSIDFW_DFW2
sumNSIDFW_DFW2
End Sub

Private Sub CommandButton9_Click()
clearTable_DFW
End Sub

Private Sub Label1_Click()

End Sub
