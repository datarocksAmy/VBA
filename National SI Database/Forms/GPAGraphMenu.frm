VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} GPAGraphMenu 
   Caption         =   "GPA Graph Menu"
   ClientHeight    =   5550
   ClientLeft      =   90
   ClientTop       =   438
   ClientWidth     =   5640
   OleObjectBlob   =   "GPAGraphMenu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "GPAGraphMenu"
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

Private Sub CommandButton6_Click()

End Sub

Private Sub CommandButton4_Click()

copyandpasteGPA

End Sub

Private Sub CommandButton7_Click()
countInstType
sumSIgroup
sumNonSIgroup
sumSIGPA
sumNSIGPA




End Sub

Private Sub CommandButton8_Click()
countacademicType
DsumSIgroup
DsumNonSIgroup
DsumSIGPA
DsumNSIGPA

End Sub

Private Sub CommandButton9_Click()
clearTable_GPA
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label5_Click()

End Sub
