VERSION 5.00
Begin VB.Form FormAddNew 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New Inventory Item..."
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10410
   Begin VB.ComboBox Supplier 
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3240
      Width           =   3375
   End
   Begin VB.CommandButton ClearAllButton 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   1680
      TabIndex        =   16
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Frame ItemInfo 
      Caption         =   "Item Details"
      Height          =   3255
      Left            =   5520
      TabIndex        =   25
      Top             =   360
      Width           =   4575
      Begin VB.ComboBox DateYear 
         Height          =   315
         Left            =   1200
         TabIndex        =   12
         Text            =   "DateYear"
         Top             =   2400
         Width           =   1095
      End
      Begin VB.ComboBox DateMonth 
         Height          =   315
         Left            =   2400
         TabIndex        =   13
         Text            =   "DateMonth"
         Top             =   2400
         Width           =   855
      End
      Begin VB.ComboBox DateDay 
         Height          =   315
         Left            =   3360
         TabIndex        =   14
         Text            =   "DateDay"
         Top             =   2400
         Width           =   735
      End
      Begin VB.CheckBox GenderFemale 
         Caption         =   "Female"
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox GenderMale 
         Caption         =   "Male"
         Height          =   255
         Left            =   1200
         TabIndex        =   10
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox Category 
         Height          =   315
         ItemData        =   "FormAddNew.frx":0000
         Left            =   1200
         List            =   "FormAddNew.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   2775
      End
      Begin VB.ComboBox Size 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox Color 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "Date Added:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Gender 
         Caption         =   "Suited for:"
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Category:"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Size:"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Color:"
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.TextBox SalePrice 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Quantity 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox Description 
      Height          =   735
      Left            =   1560
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2280
      Width           =   3375
   End
   Begin VB.TextBox Cost 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox ProductName 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox ProductID 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Close 
      Caption         =   "Discard"
      Height          =   375
      Left            =   240
      TabIndex        =   18
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "Add Item..."
      Default         =   -1  'True
      Height          =   375
      Left            =   8040
      TabIndex        =   15
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label12 
      Caption         =   "Sale Price:"
      Height          =   255
      Left            =   3000
      TabIndex        =   24
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Supplier:"
      Height          =   255
      Left            =   480
      TabIndex        =   23
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Quantity:"
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "Description:"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Cost Price:"
      Height          =   255
      Left            =   480
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Product Name:"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Product ID:"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "FormAddNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset
Dim GenderVariable As String

'This form allows the user to add new records to the database.
'The user will input the required information into each field, then the program will check the data for basic input errors
'then add it into the Inventory table as well as recording it in the transaction journal.

Private Sub Form_Load()
Call PopulateCategory   'Call procedure to load categories list from database
Call DateList   'Adding dates to the date combobox list
End Sub
Private Sub AddButton_Click()
    Call EmptyEntriesCheck 'Perform Error Checking before calling procedure InsertRecord
End Sub

Private Sub InsertRecord()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")

'Adding items to Inventory table
'Open recordset from Inventory table from database
Set rst = db.OpenRecordset("Inventory")
        rst.AddNew 'Open Empty Record
        
        'Adding each field values to the record
        rst!ProductID = ProductID.Text
        rst!PName = ProductName.Text
        rst!Cost = Cost.Text
        rst!Price = SalePrice.Text
        rst!Description = Description.Text
        rst!Quantity = Quantity.Text
        rst!Category = Category.Text
        rst!Color = Color.Text
        rst!Size = Size.Text
        
        'Take the numbers on the left as Supplier ID as defined in the Supplier table
        rst!Supplier = Left(Supplier.Text, InStr(1, Supplier.Text, " ", 1) - 1)

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
        
        'Concatenating the Month, Day, and Year into a single string, then add to database
        DateAddedValue = Month(CDate("1 " & DateMonth)) & "/" & DateDay & "/" & DateYear
        rst!DateAdded = DateAddedValue
   
        rst.Update 'Save changes to record
        rst.Close  'Closes the opened recordset
        
'Record the new incoming item to Journal
'Open the Journal table from database
Set rst = db.OpenRecordset("Journal")
        rst.AddNew 'Open new empty record
           'Add each field values to the corresponding record fields
           rst!TDate = DateAddedValue
           rst!ProductID = ProductID.Text
           rst!TransactionType = "Purchases"
           rst!TQuantity = Quantity.Text
           rst!TSaleValue = Cost.Text
           'Taking the supplier ID from the left of selected Supplier string
           rst!SupplierID = Left(Supplier.Text, InStr(1, Supplier.Text, " ", 1) - 1)
        rst.Update 'Save changes to the record
        rst.Close   'Closes the opened recordset
    
    'Ask user if wanted to add another entry
    NextEntry = MsgBox("New Entry Added to Database! Do you want to add another entry?", vbYesNo, Successful)
    If NextEntry = vbYes Then
        Call ClearAll   'If Yes, clear all the textboxes for the user to input another new entry
    Else
        'If No, unload the database and the form
        Set rst = Nothing
        db.Close
        Set db = Nothing
        Unload Me
    End If
    
End Sub

Private Sub ClearAllButton_Click()
Call ClearAll   'Call the ClearAll procedure to clear all input boxes
End Sub
Private Sub ClearAll()
Dim Ctrl As Control
For Each Ctrl In Me.Controls    'For all controls in the current form
    If TypeOf Ctrl Is TextBox Then
        Ctrl.Text = ""          'If the control is a textbox, set it to empty string
    End If
    If TypeOf Ctrl Is ComboBox Then
        Ctrl.ListIndex = -1     'If the control is a combobox, set the currently selected index to default (-1)
    End If
    If TypeOf Ctrl Is CheckBox Then
        Ctrl.Value = 0          'If the control is a checkbox, unset the checkbox.
    End If
Next
End Sub

Private Sub Close_Click()
'Ask user to confirm discarding all item
Message = MsgBox("Are you sure you want to discard this item?", vbOKCancel, "Confirm Cancel")

'If Yes, then unload the current form
If Message = vbOK Then
    Unload Me
Else
    'Do nothing
End If

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

'Set the default value for dates to the current date
DateYear = Year(Now)
DateMonth = MonthName(Month(Now))
DateDay = Day(Now)

End Sub

Private Sub PopulateCategory()
'Populate Categories and Supplier List at runtime
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open up the category table
Set rst = db.OpenRecordset("Category")

'If it has not reached the end of the table, add the item into Category combobox, then rst.MoveNext will move the pointer to the next item and reiterate
    Do While Not rst.EOF
            Category.AddItem rst!Category
            rst.MoveNext
    Loop
rst.Close   'Close the Category Table

'Open up the Supplier Table
Set rst = db.OpenRecordset("Supplier")
'If it has not reached the end of Supplier table, it will concatenate the Supplier ID, Supplier Name and Supplier Company, then add the string into Supplier combobox
    Do While Not rst.EOF
        Supplier.AddItem rst!SupplierID & " - " & rst!SName & " - " & rst!SupplierCompany
        rst.MoveNext    'After adding the string, move the pointer to the next record and reiterate
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
ElseIf Supplier.ListIndex = -1 Then
    'Check if quantity is empty or if the user has inputted a non-numerical value for quantity
    Message = MsgBox("Please check your supplier.", vbOKOnly, "Error")
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
Call RepeatIDCheck
End If

End Sub

Private Sub RepeatIDCheck()
'Check if there are existing product IDs in the database
'If yes, the program will offer to add new purchases to existing entries.

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
Set rst = db.OpenRecordset("Inventory")
'Open the Inventory table

'Initialise FoundFlag
Let FoundFlag = 0

'Repeat until the end of recordset, if there are any existing product ID that match the new input
Do Until rst.EOF
    If rst(0) = ProductID.Text Then
            Let FoundFlag = 1
            'If found, set FoundFlag
    End If
    'Move to next record
    rst.MoveNext
Loop

If FoundFlag = 0 Then
    'If there aren't any existing records with the same ID, move on to InsertRecord
    Call InsertRecord
Else
    'If there are already existing record with the same ID, ask the user whether to add new purchases instead
    AskIfIncrease = MsgBox("There are already an item with the same ID, would you like to add to the quantity of that item instead?", vbYesNo, "Duplicate ID Found")
    'If user say yes, call IncreaseStock
    If AskIfIncrease = vbYes Then
        Call IncreaseStock
    End If
End If
End Sub

Private Sub IncreaseStock()
'Add new purchases to existing entries

    AddQty = InputBox("The quantity of item with Product ID " & ProductID.Text & " to add: ", "Add Stock")
    'Call up an input box to receive input
    
    'Check for valid input
    If AddQty = 0 Or AddQty = "" Or IsNumeric(AddQty) = False Then
        MsgBox "Please input a valid value and try again.", vbOKOnly, "Error"
        Call IncreaseStock
    Else
        'If input is valid, update the new amount
        'Set path of database
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        'Form an SQL expression to select those records that match the requested productID
        EditedEntry = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & ProductID.Text
        'Create a Result Set that matches the SQL query
        Set rst = db.OpenRecordset(EditedEntry)
            
            'Load data for journaling purposes
            PriceforJournal = rst!Cost
            SIDforJournal = rst!Supplier
            
            rst.Edit 'Edit existing record
            NewQuantity = rst!Quantity + AddQty
            'Adding the user inputted quantity to existing records
                rst!Quantity = NewQuantity
            rst.Update
            rst.Close
            
        'Open the Journal table
        Set rst = db.OpenRecordset("Journal")
            rst.AddNew
            'Open a new record and add the following items to the journal
                rst!TDate = DateValue(Now)
                rst!ProductID = ProductID.Text
                rst!TransactionType = "Purchases"
                rst!TQuantity = AddQty
                rst!TSaleValue = PriceforJournal
                rst!SupplierID = SIDforJournal
        'Notify the user of successful entry
        MsgBox "The new stock level of item " & ProductID.Text & " is " & NewQuantity & ".", vbOKOnly, "Stock updated"
        
        'Unload everything
        rst.Update
        rst.Close
        Set rst = Nothing
        Set db = Nothing
        Unload Me
    End If
End Sub

