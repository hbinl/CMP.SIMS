VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormSuppliers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Suppliers Management"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10470
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   9120
      TabIndex        =   22
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton Edit 
      Caption         =   "Edit..."
      Height          =   375
      Left            =   9120
      TabIndex        =   21
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add New"
      Height          =   375
      Left            =   9120
      TabIndex        =   20
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Details"
      Enabled         =   0   'False
      Height          =   2415
      Left            =   240
      TabIndex        =   1
      Top             =   4560
      Width           =   8655
      Begin VB.CommandButton Save 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Country 
         Height          =   285
         Left            =   4800
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox State 
         Height          =   285
         Left            =   4800
         TabIndex        =   9
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Postcode 
         Height          =   285
         Left            =   7200
         TabIndex        =   8
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox City 
         Height          =   285
         Left            =   4800
         TabIndex        =   7
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox Street 
         Height          =   285
         Left            =   4800
         TabIndex        =   6
         Top             =   360
         Width           =   3495
      End
      Begin VB.TextBox SEmail 
         Height          =   285
         Left            =   960
         TabIndex        =   5
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox STelephone 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox SCompany 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox SName 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Country:"
         Height          =   255
         Left            =   4080
         TabIndex        =   19
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "State:"
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Postcode:"
         Height          =   255
         Left            =   6360
         TabIndex        =   17
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "City/Town:"
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Street:"
         Height          =   255
         Left            =   4200
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Phone:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Company:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   7223
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "FormSuppliers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NewOrEdit As Integer
'To define whether the user is editing or adding a new record, Edit:0, New:1

Private Sub Form_Load()
'Load the column headers
With ListView1.ColumnHeaders
    .Add 1, , "ID", 400
    .Add 2, , "Dealer Name", 2000
    .Add 3, , "Company Name", 2000
    .Add 4, , "Telephone", 1200
    .Add 5, , "Email Address", 3000
    .Add 6, , "Address", 7000
End With

'Populate the listview
Call PopulateListView
Call LoadDetails
End Sub

Private Sub PopulateListView()
'Populate the list from database

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the supplier table
Set rst = db.OpenRecordset("Supplier")
'Clear the ListView
ListView1.ListItems.Clear


Do Until rst.EOF
    Set StockList = ListView1.ListItems.Add(, , rst(0))
    'Load each column into listview
    For i = 1 To 4
        StockList.SubItems(i) = rst(i)
    Next
    'For address, concatenate the separate fields into one
    StockList.SubItems(5) = rst(5) & ", " & rst(6) & " " & rst(7) & ", " & rst(8) & ", " & rst(9) & "."
    'Go to next record
    rst.MoveNext
Loop

'Close the database and record
Set rst = Nothing
db.Close
Set db = Nothing
End Sub

Private Sub LoadDetails()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query Supplier records that matches the selected criteria
SQL = "SELECT Supplier.* FROM Supplier WHERE Supplier.SupplierID = " & ListView1.SelectedItem
'Open the recordset as defined from the query
Set rst = db.OpenRecordset(SQL)
    'Start to load items
    SName.Text = rst!SName
    SCompany.Text = rst!SupplierCompany
    STelephone.Text = rst!STelephone
    SEmail.Text = rst!SEmail
    Street.Text = rst!SStreet
    City.Text = rst!SCity
    Postcode.Text = rst!SPostcode
    State.Text = rst!SState
    Country.Text = rst!SCountry
Set rst = Nothing
db.Close
Set db = Nothing

'Disable the frame from being edited
'Hide the save button
Frame1.Enabled = False
Save.Visible = False
End Sub

Private Sub ListView1_Click()
'Whenever the user click on an entry, automatically loads the details
Call LoadDetails
End Sub

Private Sub Add_Click()
'The user presses Add New button

Let NewOrEdit = 1
'Make the save button visible
Save.Visible = True
'Allow editing in the form below
Frame1.Enabled = True

'Clear all textboxes in the form
Dim Ctrl As Control
For Each Ctrl In Me.Controls
    If TypeOf Ctrl Is TextBox Then
        Ctrl.Text = ""
    End If
Next
End Sub

Private Sub Edit_Click()
'The user presses Edit button

Let NewOrEdit = 0
'Make the save button visible
Save.Visible = True
'Allow editing in the form below
Frame1.Enabled = True
End Sub
Private Sub Save_Click()
'Check for entry validity
Call EntryCheck
End Sub

Private Sub EntryCheck()
'Input validation
If SName = "" Then
    'Check if Name is empty
    Message = MsgBox("Please fill a valid name.", vbOKOnly, "Error")
    Exit Sub
ElseIf SCompany = "" Then
    'Check if Company Name is empty
    Message = MsgBox("Please fill a valid company name.", vbOKOnly, "Error")
    Exit Sub
ElseIf STelephone = "" Or IsNumeric(STelephone) = False Then
    'Check if Telephone is empty or Numerical
    Message = MsgBox("Please fill a valid telephone number.", vbOKOnly, "Error")
    Exit Sub
ElseIf SEmail = "" Then
    'Check if Email is empty
    Message = MsgBox("Please fill a valid email.", vbOKOnly, "Error")
    Exit Sub
ElseIf Street = "" Then
    'Check if Street is empty
    Message = MsgBox("Please fill a valid street address", vbOKOnly, "Error")
    Exit Sub
ElseIf City = "" Then
    'Check if City is empty
    Message = MsgBox("Please fill a valid city.", vbOKOnly, "Error")
    Exit Sub
ElseIf Postcode = "" Or IsNumeric(Postcode) = False Then
    'Check if Postcode is empty or numerical
    Message = MsgBox("Please fill a valid postcode.", vbOKOnly, "Error")
    Exit Sub
ElseIf State = "" Then
    'Check if State is empty
    Message = MsgBox("Please fill a valid state", vbOKOnly, "Error")
    Exit Sub
ElseIf Country = "" Then
    'Check if Country is empty
    Message = MsgBox("Please fill a valid country.", vbOKOnly, "Error")
    Exit Sub
Else
'No input error
Call SaveCalls
End If
End Sub

Private Sub SaveCalls()
'The user presses the save button

If NewOrEdit = 0 Then 'If the NewOrEdit flag is 0, then edit record
    Call EditRecord
ElseIf NewOrEdit = 1 Then 'If the NewOrEdit flag is 1, then add new record
    Call AddRecord
Else
    'Do nothing
End If
    'Reload the listview and disable editing for now
    Call PopulateListView
    Frame1.Enabled = False
    Save.Visible = False
End Sub

Private Sub AddRecord()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the supplier table
Set rst = db.OpenRecordset("Supplier")
    'Open a new record
    rst.AddNew
    'Start to add items to database
    rst!SName = SName.Text
    rst!SupplierCompany = SCompany.Text
    rst!STelephone = STelephone.Text
    rst!SEmail = SEmail.Text
    rst!SStreet = Street.Text
    rst!SCity = City.Text
    rst!SPostcode = Postcode.Text
    rst!SState = State.Text
    rst!SCountry = Country.Text
    'Save changes
    rst.Update
    rst.Close
'Close database and recordset
Set rst = Nothing
db.Close
Set db = Nothing
End Sub

Private Sub EditRecord()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the criteria from database
SQL = "SELECT Supplier.* FROM Supplier WHERE Supplier.SupplierID = " & ListView1.SelectedItem
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(SQL)
    'Open the record for edit
    rst.Edit
    'Saving each fields into the database
    rst!SName = SName.Text
    rst!SupplierCompany = SCompany.Text
    rst!STelephone = STelephone.Text
    rst!SEmail = SEmail.Text
    rst!SStreet = Street.Text
    rst!SCity = City.Text
    rst!SPostcode = Postcode.Text
    rst!SState = State.Text
    rst!SCountry = Country.Text
    'Save changes
    rst.Update
    rst.Close
'Close database and recordset
Set rst = Nothing
db.Close
Set db = Nothing
End Sub

Private Sub Delete_Click()
'Confirm with user on whether to delete or not
Message = MsgBox("Are you sure you want to delete this supplier?", vbYesNo, "Confirm Delete")
If Message = vbYes Then
    Call DeleteRecord
Else
    'Do nothing
End If
End Sub

Private Sub DeleteRecord()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the criteria from database
    SelectedDelete = "SELECT Supplier.* " & "FROM Supplier " & "WHERE Supplier.SupplierID = " & ListView1.SelectedItem & ""
    'Open the result set as defined by the query above
    Set MarkedForDeletion = db.OpenRecordset(SelectedDelete)
    'Delete records
    With MarkedForDeletion
        Do While Not .EOF
        .Delete
        .MoveNext
        Loop
    End With
    Set MarkedForDeletion = Nothing
    db.Close
    'Removing it from listview
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub


