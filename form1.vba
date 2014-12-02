VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form1 
   Caption         =   "New Volunteer"
   ClientHeight    =   8355.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6810
   OleObjectBlob   =   "form1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()

Dim ctl As Control

Sheets("Active_Volunteers").Activate
RowCount = Worksheets("Active_Volunteers").Range("A6").CurrentRegion.Rows.Count + 5


' We get all the relevant column number from cell names and store them in vars.
nameColumn = Range("name").Column
phoneColumn = Range("phone").Column
emailColumn = Range("email").Column
emergencyContactNameColumn = Range("emergency_name").Column
emergencyContactNumberColumn = Range("emergency_number").Column
avaiabilityColumn = Range("availability").Column
restaurantLocationColumn = Range("restaurant_location").Column
numberOfDaysColumn = Range("number_of_days").Column
hearAboutColumn = Range("hear_about").Column
skillsColumn = Range("skills").Column
commentsColumn = Range("comments").Column

'We perform all relevant checks before using the entered values and store them in the spreadsheet

 If nameBox.Value = "" Then
     MsgBox "Please enter a First Name.", vbExclamation, "Volunteer Name"
        nameBox.SetFocus
        Exit Sub
    End If
    
If lastNameBox.Value = "" Then
     MsgBox "Please enter a Last Name.", vbExclamation, "Volunteer Name"
        lastNameBox.SetFocus
        Exit Sub
    End If
    
If phoneBox.Value = "" Then
    If emailBox.Value = "" Then
    
     MsgBox "Please enter a phone Number or an Email address.", vbExclamation, "Volunteer contact detail"
        phoneBox.SetFocus
        Exit Sub
    End If
    End If
    
    
If Not phoneBox.Value = "" Then
    If Not IsNumeric(phoneBox.Value) Then
    MsgBox "The phone number must contain only numbers.", vbExclamation, "Volunteer contact detail"
    phoneBox.SetFocus
    Exit Sub
    End If
End If

'We store all the values in the relevant cells in the spreadsheet.

'We increment by one the ID number
Cells(RowCount, 1).Value = Cells(RowCount - 1, 1).Value + 1

'store the name
Cells(RowCount, nameColumn).Value = nameBox.Value + " " + lastNameBox.Value

'store the phone number
Cells(RowCount, phoneColumn).Value = phoneBox.Value

'store the email
Cells(RowCount, emailColumn).Value = emailBox.Value

'store the emeergency contact name
Cells(RowCount, emergencyContactNameColumn).Value = emergencyBox.Value

'store the emergency contact number
Cells(RowCount, emergencyContactNumberColumn).Value = emergencyNumberBox.Value

'store the restaurant location
Cells(RowCount, restaurantLocationColumn).Value = locationSydney.Caption

'store all the skills and create boolean values for floor, kitchen and diswash
If kitchenCheck.Value = True Then Cells(RowCount, skillsColumn).Value = kitchenCheck.Caption
If kitchenCheck.Value = True Then kitchenText = "K " Else kitchenText = ""

If floorCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + floorCheck.Caption
If floorCheck.Value = True Then floortext = "F " Else floortext = ""

If dishwashCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + dishwashCheck.Caption
If dishwashCheck.Value = True Then dishwashText = "D " Else dishwashText = ""

If baristaCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + baristaCheck.Caption

If bartenderCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + bartenderCheck.Caption
If bartenderCheck.Value = True Then bartenderText = "B " Else bartenderText = ""

If fundraisingCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + fundraisingCheck.Caption

If adminCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + adminCheck.Caption

If commCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + commCheck.Caption

If otherCheck.Value = True Then Cells(RowCount, skillsColumn).Value = Cells(RowCount, skillsColumn).Value + ", " + otherCheckText.Value

'We store the availabilities. we create two array to store the days (mon to sun) and time of day (lunch or dinner)
Dim tab_day(7)
tab_day(0) = "Mon"
tab_day(1) = "Tue"
tab_day(2) = "Wed"
tab_day(3) = "Thu"
tab_day(4) = "Fri"
tab_day(5) = "Sat"
tab_day(6) = "Sun"

Dim tab_time(2)
tab_time(0) = "Lun"
tab_time(1) = "Din"

Dim checkBox As String
Dim checkBoxBackup As String

For i = 0 To 6
 For n = 0 To 1
    checkBox = tab_day(i) + tab_time(n)
    checkBoxBackup = tab_day(i) + tab_time(n) + "B"
    Column = Range(checkBox).Column
    checkObject = form1(checkBox).Value
    checkObjectB = form1(checkBoxBackup).Value
    If checkObject = True Then
        Cells(RowCount, Column).Value = kitchenText + floortext + dishwashText + bartenderText
        ElseIf checkObjectB = True Then Cells(RowCount, Column).Value = kitchenText + floortext + dishwashText + bartenderText + " (Backup)"
    End If
 Next
 
Next

'We store the number of shifts per week
Cells(RowCount, numberOfDaysColumn).Value = numShiftsBox.Value

'We store the specific skills and comments
Cells(RowCount, commentsColumn).Value = Cells(RowCount, commentsColumn).Value + " " + skillsText.Value

'We store the way he/she heard about lentil
If facebookCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + facebookCheck.Caption
If websiteCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + websiteCheck.Caption
If restaurantCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + restaurantCheck.Caption
If friendCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + friendCheck.Caption
If familyCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + familyCheck.Caption
If volunteerCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + volunteerCheck.Caption
If staffCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + staffCheck.Caption
If otherSourceCheck.Value = True Then Cells(RowCount, hearAboutColumn).Value = Cells(RowCount, hearAboutColumn).Value + " " + otherSourceBox.Value


For Each ctl In Me.Controls
If TypeName(ctl) = "TextBox" Or TypeName(ctl) = "ComboBox" Then
ctl.Value = ""
ElseIf TypeName(ctl) = "CheckBox" Then
ctl.Value = False
End If
Next ctl




End Sub


Private Sub CommandButton2_Click()
Unload form1
End Sub






Private Sub floorCheck_Click()

End Sub

Private Sub nameBox_Change()

End Sub
