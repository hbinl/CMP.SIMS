VERSION 5.00
Begin VB.Form FormSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sale of Item..."
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   480
      ScaleHeight     =   795
      ScaleWidth      =   3555
      TabIndex        =   14
      Top             =   2640
      Width           =   3615
   End
   Begin VB.TextBox ProductName 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox DateYear 
      Height          =   315
      Left            =   1560
      TabIndex        =   11
      Text            =   "DateYear"
      Top             =   2160
      Width           =   855
   End
   Begin VB.ComboBox DateMonth 
      Height          =   315
      Left            =   2400
      TabIndex        =   10
      Text            =   "DateMonth"
      Top             =   2160
      Width           =   975
   End
   Begin VB.ComboBox DateDay 
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Text            =   "DateDay"
      Top             =   2160
      Width           =   615
   End
   Begin VB.TextBox SalePrice 
      Height          =   285
      Left            =   1560
      TabIndex        =   7
      Top             =   1680
      Width           =   975
   End
   Begin VB.CommandButton Discard 
      Caption         =   "Discard"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Commit 
      Caption         =   "Commit"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
   Begin VB.ComboBox QuantitySold 
      Height          =   315
      ItemData        =   "FormSales.frx":0000
      Left            =   1560
      List            =   "FormSales.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox PID 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "Date Sold:"
      Height          =   255
      Left            =   720
      TabIndex        =   12
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Sale Price:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Quantity Sold:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label ProductNameLabel 
      Caption         =   "Product Name:"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Product ID:"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "FormSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OriginalCost As Currency
Dim GrossProfit As Currency
Dim TransValue As Currency

Private Sub Discard_Click()
'Unload the current form if Discard is clicked
Unload Me
End Sub

Private Sub UpdateDisplay()
On Error GoTo ErrorHandling
'When the user selects the quantity to be sold, or changes the sale price
'Update the mini display picturebox below to show
'Original total cost, total value of current transaction, and the gross profit before tax

'Calculations
TransValue = QuantitySold * SalePrice.Text
OriginalTotalCost = OriginalCost * QuantitySold
GrossProfit = TransValue - OriginalTotalCost

'Clear the picturebox
Picture1.Cls
'Display the items
Picture1.Print "Original Total Cost: " & FormatCurrency(OriginalTotalCost)
Picture1.Print "Total Value of transaction: " & FormatCurrency(TransValue)
Picture1.Print "Gross Profit before Tax: " & FormatCurrency(GrossProfit)
Exit Sub


ErrorHandling:
Select Case Err.Number
    Case 13
        Picture1.Cls
        Picture1.Print "Please input a valid selling price."
    Case Else
        Picture1.Print "Unknown error occurred."
End Select
End Sub


Private Sub Form_Load()
'Initialisation
TransValue = 0
OriginalCost = 0
GrossProfit = 0

'Populate the comboboxes and load certain stock info from database
Call DateList
Call LoadStock
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

'Set date comboboxes to current date
DateYear = Year(Now)
DateMonth = MonthName(Month(Now))
DateDay = Day(Now)
End Sub
Private Sub LoadStock()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL query expression to search for items record that were selected in InventoryListView
SoldItem = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & Inventory.ListView1.SelectedItem
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(SoldItem)

'Load items
Let PID.Text = Inventory.ListView1.SelectedItem
ProductName.Text = rst(1)
SalePrice.Text = FormatNumber(rst(3), 2)

'Populate the selectable quantity sold to combobox
For i = 0 To rst(4)
    QuantitySold.AddItem i
Next i

'Initialise QuantitySold
Let QuantitySold = 0
Let OriginalCost = rst(2)

'If the quantity recorded in database is 0
'Report to user out of stock and stop sales
If rst(4) = 0 Then
    MsgBox "No sales can be committed because there isn't any stock left for this item (PID: " & PID.Text & "). Please contact your supplier to restock.", vbOKOnly, "Out of stock"
    QuantitySold.Enabled = False
    SalePrice.Enabled = False
    DateDay.Enabled = False
    DateMonth.Enabled = False
    DateYear.Enabled = False
    Commit.Enabled = False
End If

'Unload database and record
Set rst = Nothing
db.Close
Set db = Nothing
End Sub

Private Sub Commit_Click()
'When the user clicks Commit, check for entry validity
Call CheckEntry
End Sub

Private Sub CheckEntry()
'Check if quantity sold is valid
If QuantitySold <> 0 Then
    'Check if sale price is valid
    If SalePrice.Text <> "" Or IsNumeric(SalePrice.Text) Then
        Call WriteSales
    Else
        MsgBox "Please input the price at which the good is sold.", vbOKOnly, "Error"
    End If
Else
    MsgBox "Please select the quantity sold.", vbOKOnly, "Error"
End If
End Sub
Private Sub WriteSales()
On Error GoTo ErrHandling
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query records that matches the criteria selected
SalesEntry = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & Inventory.ListView1.SelectedItem & ""
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(SalesEntry)
        'Open the record for edit
        rst.Edit
        'Saving the field datas into record
        rst!Quantity = rst!Quantity - QuantitySold
        rst!DateAdded = Month(CDate("1 " & DateMonth)) & "/" & DateDay & "/" & DateYear
        'Save changes
        rst.Update
        'Close the record
        rst.Close

'Open the Journal table to record this transaction
Set rst = db.OpenRecordset("Journal")
    'Open a new record
    rst.AddNew
    'Adding items to this new record
    rst!TDate = Month(CDate("1 " & DateMonth)) & "/" & DateDay & "/" & DateYear
    rst!ProductID = Inventory.ListView1.SelectedItem
    rst!TransactionType = "Sales"
    rst!SupplierID = "0"
    rst!TSaleValue = TransValue
    rst!TQuantity = QuantitySold
    rst!GrossProfit = GrossProfit
    'Save Changes
    rst.Update
    'Close the record
    rst.Close
    
    Set rst = Nothing
    db.Close
    Set db = Nothing
    
'Notify the user of successful entry
MsgBox "Sales successfully recorded!"
Unload Me
Exit Sub

ErrHandling:
    Select Case Err.Number
        Case 13 'Type mismatch
            MsgBox "Invalid value entered.", vbOKOnly, "Error"
        Case Else
            MsgBox "Unknown error occurred.", vbOKOnly, "Error"
    End Select
End Sub


Private Sub QuantitySold_Click()
'Check if the user has selected valid quantity and sale prices
'Then call UpdateDisplay
If QuantitySold.ListIndex <> -1 And SalePrice.Text <> "" Then
    Call UpdateDisplay
Else
    Picture1.Cls
    Picture1.Print "Please choose a valid quantity and sale price."
End If
End Sub

Private Sub SalePrice_Change()
'Check if the user has selected valid quantity and sale prices
'Then call UpdateDisplay
If QuantitySold.ListIndex <> -1 And SalePrice.Text <> "" Then
    Call UpdateDisplay
Else
    Picture1.Cls
    Picture1.Print "Please choose a valid quantity and sale price."
End If
End Sub
