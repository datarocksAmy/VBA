Attribute VB_Name = "UI_Module"
' Module for calling the User Interface (Format Category)
Sub Workbook_Open_FormatCategory()
    Worksheets("Raw Data").Activate
    FormatCategoryMenu.Show
End Sub

' Module for calling the User Interface (GPA)
Sub Workbook_Open()
    Worksheets("GPA Graph").Activate
    GPAGraphMenu.Show
End Sub

' Module for calling the User Interface (DFW)
Sub Workbook_Open_DFW()
    Worksheets("DFW Graph").Activate
    DFWGraphMenu.Show
End Sub

' Module for calling the User Interface (Pie Graph)
Sub Workbook_Open_Pie()
    Worksheets("Pie Graph").Activate
    PieGraphMenu.Show
End Sub

' Module for calling the User Interface (Clean Data)
Sub Workbook_Open_CleanData()
    Worksheets("Data Clean").Activate
    CleanDataMenu.Show
End Sub

