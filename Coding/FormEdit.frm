VERSION 5.00
Begin VB.Form FormEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Inventory Item..."
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4785
   ScaleMode       =   0  'User
   ScaleWidth      =   10410
   Begin VB.CommandButton PreviousRecord 
      Caption         =   "<<"
      Height          =   375
      Left            =   7560
      TabIndex        =   31
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton NextRecord 
      Caption         =   ">>"
      Height          =   375
      Left            =   9840
      TabIndex        =   30
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton EditButton 
      Caption         =   "Save Changes"
      Default         =   -1  'True
      Height          =   375
      Left            =   8040
      TabIndex        =   22
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Close 
      Caption         =   "Discard Changes"
      Height          =   375
      Left            =   360
      TabIndex        =   21
      Top             =   4080
      Width           =   1575
   End
   Begin VB.TextBox ProductID 
      Height          =   285
      Left            =   1680
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox ProductName 
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox Cost 
      Height          =   285
      Left            =   1680
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Description 
      Height          =   735
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   17
      Top             =   2400
      Width           =   3375
   End
   Begin VB.TextBox Quantity 
      Height          =   285
      Left            =   1680
      TabIndex        =   16
      Top             =   1920
      Width           =   855
   End
   Begin VB.TextBox SalePrice 
      Height          =   285
      Left            =   3960
      TabIndex        =   15
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame ItemInfo 
      Caption         =   "Item Details"
      Height          =   3255
      Left            =   5640
      TabIndex        =   1
      Top             =   480
      Width           =   4575
      Begin VB.ComboBox Color 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox Size 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Category 
         Height          =   315
         ItemData        =   "FormEdit.frx":0000
         Left            =   1200
         List            =   "FormEdit.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox GenderMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1920
         Width           =   735
      End
      Begin VB.CheckBox GenderFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox DateDay 
         Height          =   315
         Left            =   3360
         TabIndex        =   4
         Text            =   "DateDay"
         Top             =   2400
         Width           =   735
      End
      Begin VB.ComboBox DateMonth 
         Height          =   315
         Left            =   2400
         TabIndex        =   3
         Text            =   "DateMonth"
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox DateYear 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Text            =   "DateYear"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Color:"
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Size:"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Category:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Gender 
         Caption         =   "Suited for:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Date Added:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2400
         Width           =   1095
      End
   End
   Begin VB.ComboBox Supplier 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Product ID:"
      Height          =   255
      Left            =   600
      TabIndex        =   29
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   28
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Cost Price:"
      Height          =   255
      Left            =   600
      TabIndex        =   27
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Description:"
      Height          =   255
      Left            =   480
      TabIndex        =   26
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier:"
      Height          =   255
      Left            =   600
      TabIndex        =   24
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label12 
      Caption         =   "Sale Price:"
      Height          =   255
      Left            =   3120
      TabIndex        =   23
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "FormEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the user to edit existing records selected in the Inventory form ListView

Dim db As Database
Dim rst As Recordset
Dim srst As Recordset

Private Sub Close_Click()
Message = MsgBox("Are you sure you want to discard all changes?", vbOKCancel, "Confirm Cancel")
If Message = vbOK Then
    Unload Me
Else
    'Do nothing
End If
End Sub

Private Sub EditButton_Click()
Call EmptyEntriesCheck
End Sub

Private Sub Form_Load()
'Populate the comboboxes and datelist
Call PopulateCategory
Call DateList

'Set date comboboxes to current date
DateYear = Year(Now)
DateMonth = Month(Now)
DateDay = Day(Now)

'Load record from database
Call LoadRecord
End Sub

Private Sub DateList()
'Populating the dates combobox

For i = 1 To 31
    DateDay.AddItem i   'Adding the day to Day combobox
Next

For i = 1 To 12
    DateMonth.AddItem MonthName(i)  'Adding the months to Month combobox
Next


For i = 1 To 30
    DateYear.AddItem (Year(Now) - i)    'Adding years to Years combobox for the past 30 years up to the current year
Next

End Sub
Private Sub PopulateCategory()
'Populate Categories List'Populate Categories and Supplier List at runtime
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the Category Table
Set rst = db.OpenRecordset("Category")


    'If it has not reached the end of the table, add the item into Category combobox, then rst.MoveNext will move the pointer to the next item and reiterate
    Do While Not rst.EOF
            Category.AddItem rst!Category
            rst.MoveNext
    Loop
    rst.Close   'Close the Category Table
    Set rst = Nothing


    'Open up Supplier Table
    Set rst = db.OpenRecordset("Supplier")
    'If it has not reached the end of Supplier table, it will concatenate the Supplier ID, Supplier Name and Supplier Company, then add the string into Supplier combobox
    Do While Not rst.EOF
        Supplier.AddItem rst!SupplierID & " - " & rst!SName & " - " & rst!SupplierCompany
        rst.MoveNext
    Loop
    Set rst = Nothing
    db.Close
    Set db = Nothing
    'Closes the recordset and database.

'Populate Colors and Size
Call AddColor
Call AddSize
End Sub

Private Sub AddColor()
'Adding each color item into the Color combobox selection
With Color
    .AddItem "Red"
    .AddItem "Green"
    .AddItem ("Blue")
    .AddItem ("White")
    .AddItem ("Yellow")
    .AddItem ("Pink")
    .AddItem ("Grey")
    .AddItem ("Black")
End With
End Sub


Private Sub AddSize()
'Adding each size item into the Size combobox selection
With Size
    .AddItem ("XS")
    .AddItem ("S")
    .AddItem ("M")
    .AddItem ("L")
    .AddItem ("XL")
    .AddItem ("None")
End With

End Sub
Private Sub LoadRecord()
On Error Resume Next

    'Set path of database
    Set db = OpenDatabase(App.Path & "/Database.MDB")
    'Form SQL query to select records that matches the selected item from ListView
    SelectedRecord = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & Inventory.ListView1.SelectedItem & ""
    'Open result set that matches the query
    Set rst = db.OpenRecordset(SelectedRecord)
        'Start loading items from record into respective fields
        ProductID.Text = rst!ProductID
        ProductName.Text = rst!PName
        'Apply formatting to the cost and price to 2 decimal places
        Cost.Text = FormatNumber(rst!Cost, 2)
        SalePrice.Text = FormatNumber(rst!Price, 2)
        Description.Text = rst!Description
        Quantity.Text = rst!Quantity
        Category.Text = rst!Category
        Color.Text = rst!Color
        Size.Text = rst!Size
        
        'Form SQL Query to select supplier names from Supplier table based on supplierID stored in Inventory table
        SQL = "Select Supplier.Sname, Supplier.SupplierCompany FROM Supplier WHERE Supplier.SupplierID = " & rst!Supplier
        'Open the result set that matches the query above
        Set srst = db.OpenRecordset(SQL)
        'Set supplier combobox to default
        Supplier.ListIndex = -1
        'If there are no records in the result set, leave the combobox default/blank
        If srst.RecordCount < 0 Then
            Supplier.ListIndex = -1
        Else
        'Else if there are records, combobox set selection according to the SupplierID
            Supplier.Text = rst!Supplier & " - " & srst(0) & " - " & srst(1)
        End If
        Set srst = Nothing
        
        
        If rst!Gender = "M" Then 'Check if it's male
            GenderMale.Value = 1
            GenderFemale.Value = 0
        ElseIf rst!Gender = "F" Then  'Check if it's female
            GenderMale.Value = 0
            GenderFemale.Value = 1
        ElseIf rst!Gender = "U" Then 'Both selected or deselected means it's suitable for both gender
            GenderMale.Value = 1
            GenderFemale.Value = 1
        End If
        
        'Extracting the date stored in the database to separate components of Year, Month, Day
        DateAddedValue = rst!DateAdded
        DateYear = Right(DateAddedValue, 4)
        DateMonth = MonthName(Left(DateAddedValue, InStr(1, DateAddedValue, "/", 1) - 1))
        DateExtract = Left(DateAddedValue, Len(DateAddedValue) - 5)
        DateDay = Right(DateExtract, Len(DateExtract) - InStr(1, DateExtract, "/", 1))
        
        'Finish, unload database and recordset
        Set rst = Nothing
        db.Close
        Set db = Nothing
End Sub


Private Sub ChangeRecord()
'The user clicks Commit Changes

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL query to request for the record that matches the product ID
EditedEntry = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & Inventory.ListView1.SelectedItem & ""
'Load the record that matches the query
Set rst = db.OpenRecordset(EditedEntry)
        'Open the record to edit
        rst.Edit
        'Start to load each fields into each corresponding fields in database
        rst!ProductID = ProductID.Text
        rst!PName = ProductName.Text
        rst!Cost = Cost.Text
        rst!Price = SalePrice.Text
        rst!Description = Description.Text
        rst!Quantity = Quantity.Text
        rst!Supplier = Left(Supplier.Text, InStr(1, Supplier.Text, " ") - 1)
        rst!Category = Category.Text
        rst!Color = Color.Text
        rst!Size = Size.Text
        
        'Check if it's male
        If GenderMale.Value = 1 And GenderFemale.Value = 0 Then
            rst!Gender = "M"
        'Check if it's female
        ElseIf GenderMale.Value = 0 And GenderFemale.Value = 1 Then
            rst!Gender = "F"
        'Both selected or deselected means it's suitable for both gender
        Else
            rst!Gender = "U"
        End If
        
        'Recombining the date into the database
        DateAddedValue = Month(CDate("1 " & DateMonth)) & "/" & DateDay & "/" & DateYear
        rst!DateAdded = DateAddedValue
    
    'Save changes and close the database
    rst.Update
    rst.Close
    
Set rst = Nothing
db.Close
Set db = Nothing
'Notify the user of successful changes.
MsgBox "Saved changes successfully!", vbOKOnly, "Edit"
    
End Sub

Private Sub EmptyEntriesCheck()
'Input validation
If ProductID = "" Or IsNumeric(ProductID) = False Then
    'If ProductID is empty or not numerical, then stop input.
    Message = MsgBox("Please check the product ID.", vbOKOnly, "Error")
    Exit Sub
ElseIf ProductName = "" Then
    'If product name is empty, then stop input
    Message = MsgBox("Please fill in the product name.", vbOKOnly, "Error")
    Exit Sub
ElseIf Cost = "" Or IsNumeric(Cost) = False Then
    'Check if cost is empty or the user has inputted a non-numerical value for cost
    Message = MsgBox("Please check your item cost.", vbOKOnly, "Error")
    Exit Sub
ElseIf SalePrice = "" Or IsNumeric(SalePrice) = False Then
    'Check if price is empty or the user has inputted a non-numerical value for price
    Message = MsgBox("Please check your price.", vbOKOnly, "Error")
    Exit Sub
ElseIf Quantity = "" Or IsNumeric(Quantity) = False Then
    'Check if quantity is empty or if the user has inputted a non-numerical value for quantity
    Message = MsgBox("Please check your quantity.", vbOKOnly, "Error")
    Exit Sub
ElseIf Category.ListIndex = -1 Then
    'Check if user has selected any category, if not, then stop input
    Message = MsgBox("Please choose a category for this item.", vbOKOnly, "Error")
    Exit Sub
ElseIf Color.ListIndex = -1 Then
    'Check if user has selected any color, if not, then stop input
    Message = MsgBox("Please choose a color for this item.", vbOKOnly, "Error")
    Exit Sub
ElseIf Size.ListIndex = -1 Then
    'Check if user has selected any size, if not, then stop input
    Message = MsgBox("Please choose a size for this item.", vbOKOnly, "Error")
    Exit Sub
Else
'If no input error, check for repeat product IDs
Call ChangeRecord
End If
End Sub

Private Sub NextRecord_Click()
'When the user click the >> button

'Increment the Selected item index in listview by 1
X = Inventory.ListView1.SelectedItem.Index + 1

'Check if end of listview
'If not, select the next record and load the next record
'Else, report end of list
If X < Inventory.ListView1.ListItems.Count Then
    Inventory.ListView1.ListItems(X).Selected = True
    Call LoadRecord
Else
    MsgBox "You have reached the end of the list.", vbOKOnly, "Info"
    Exit Sub
End If
End Sub

Private Sub PreviousRecord_Click()
'When the user click the << button

'Decrease the Selected item index in listview by 1
X = Inventory.ListView1.SelectedItem.Index - 1

'Check if reached beginning of listview
'If not, select the previous record and load the previous record
'Else, report start of list
If X > 0 Then
    Inventory.ListView1.ListItems(X).Selected = True
    Call LoadRecord
Else
    MsgBox "You have reached the start of the list.", vbOKOnly, "Info"
    Exit Sub
End If
End Sub
