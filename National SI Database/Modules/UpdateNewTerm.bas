Attribute VB_Name = "UpdateNewTerm"
Sub DeleteOldUpdateNew()
Dim CStartRow As Integer, CEndRow As Integer
Dim RStartRow As Integer, REndRow As Integer
Dim lastrow As Integer
Dim user_prompt As String
Dim num As Integer


' Last Row number contains Data
lastrow = Sheets("Data Clean").Cells(Rows.Count, "A").End(xlUp).Row

Re_Enter:
' Ask user what term they're updating
user_prompt = InputBox(Prompt:="Enter the semester you're updating< Semester(Fall:FA|Spring:SP|Summer:SUM)+Last 2 digit of year >: ", Title:="Update Term Data", Default:="e.g. FA17")

' Cancel
If user_prompt = vbNullString Then
    MsgBox ("You just canceled updating!")
    GoTo end_program
End If

' Check if input is in sheets or not
If Sheets("Data Clean").Range("A:A").Find(what:=user_prompt, LookAt:=xlWhole) Is Nothing Then
        MsgBox ("Invalid Input. Can't find any match! Please try again.")
        GoTo Re_Enter
End If


' First & last row number for term in "Data Clean"
CStartRow = Sheets("Data Clean").Range("A:A").Find(what:=user_prompt, after:=Range("A1")).Row
CEndRow = Sheets("Data Clean").Range("A:A").Find(what:=user_prompt, after:=Range("A1"), searchdirection:=xlPrevious).Row

' Delete the data for the update term from "Data Clean"
Sheets("Data Clean").Range(("A" & CStartRow), ("A" & CEndRow)).EntireRow.Delete

' First & last row number for term in "Raw Data"
RStartRow = Sheets("Raw Data").Range("A:A").Find(what:=user_prompt, after:=Range("A1")).Row
REndRow = Sheets("Raw Data").Range("A:A").Find(what:=user_prompt, after:=Range("A1"), searchdirection:=xlPrevious).Row

' Copy the Updated Data for the specific term from "Raw Data"
' Paste the Updated data to the "Data Clean"
lastrow = Sheets("Data Clean").Cells(Rows.Count, "A").End(xlUp).Row
Sheets("Raw Data").Range(("A" & RStartRow), ("A" & REndRow)).EntireRow.Copy Destination:=Sheets("Data Clean").Cells(lastrow, "A")

' Inform user update sucessfully update
MsgBox ("Finish Updating " & user_prompt)
end_program:

End Sub

