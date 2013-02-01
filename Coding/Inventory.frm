VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Inventory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory List"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   15705
   Begin VB.CommandButton AddStock 
      Caption         =   "Add Purchases to Current Stock"
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
      Left            =   9960
      TabIndex        =   12
      ToolTipText     =   "Add Purchases to this selected item."
      Top             =   8760
      Width           =   3135
   End
   Begin VB.CommandButton SalesButton 
      Caption         =   "Sale of Item..."
      Default         =   -1  'True
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
      Left            =   13200
      TabIndex        =   11
      ToolTipText     =   "Sale of selected stock item."
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton ClearSearch 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4560
      TabIndex        =   9
      ToolTipText     =   "Clear the Search Box"
      Top             =   240
      Width           =   735
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7695
      Left            =   240
      TabIndex        =   8
      Top             =   840
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13573
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.CommandButton Search 
      Caption         =   "Find!"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      ToolTipText     =   "Begin Search"
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox SearchQuery 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   6
      Top             =   240
      Width           =   2775
   End
   Begin VB.CommandButton DeleteItem 
      Caption         =   "Delete Item..."
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Deletes the selected item."
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton ExportList 
      Caption         =   "Export/Print List..."
      Height          =   375
      Left            =   12360
      TabIndex        =   3
      ToolTipText     =   "Allows you to export or print a list of inventory items."
      Top             =   240
      Width           =   1695
   End
   Begin VB.CommandButton EditItem 
      Caption         =   "Edit Item..."
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      ToolTipText     =   "Edit the selected item."
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton RefreshList 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   14280
      TabIndex        =   1
      ToolTipText     =   "Reload the items from database."
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton AddNew 
      Caption         =   "Add New Items..."
      Height          =   375
      Left            =   240
      TabIndex        =   0
      ToolTipText     =   "Adds a new item."
      Top             =   8760
      Width           =   1455
   End
   Begin VB.Label Status 
      Caption         =   "Status"
      Height          =   255
      Left            =   5640
      TabIndex        =   10
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Search:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   330
      Width           =   735
   End
End
Attribute VB_Name = "Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the main inventory list window that lists out all the current stock records in database.
'It shows automatically during startup.

Dim db As Database
Dim rst As Recordset
Dim srst As Recordset

Private Sub Form_Load()
On Error Resume Next

'Populating the ListView column headers
With ListView1.ColumnHeaders
    .Add 1, , "PID", 600
    .Add 2, , "Product Name", 1800
    .Add 3, , "Cost", 600
    .Add 4, , "Sale Price", 1150
    .Add 5, , "Qty", 500
    .Add 6, , "Total Cost", 1100
    .Add 7, , "Description", 2600
    .Add 8, , "Supplier", 2000
    .Add 9, , "Category", 1300
    .Add 10, , "Color", 1000
    .Add 11, , "Size", 700
    .Add 12, , "Gender", 900
    .Add 13, , "Date Added", 1500
End With

'Load the listview with stock records
Call PopulateListView
End Sub
Private Sub ClearSearch_Click()
'Clear the search box
SearchQuery.Text = ""

'Reload the listview
Call PopulateListView
End Sub

Private Sub AddStock_Click()
    AddQty = InputBox("The quantity of item with Product ID " & ListView1.SelectedItem & " to add: ", "Add Stock")
    If AddQty = 0 Or AddQty = "" Or IsNumeric(AddQty) = False Then
        MsgBox "Please input a valid value and try again.", vbOKOnly, "Error"
    Else
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        EditedEntry = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & ListView1.SelectedItem
        Set rst = db.OpenRecordset(EditedEntry)
        PriceforJournal = rst!Cost
        SIDforJournal = rst!Supplier
        rst.Edit
        NewQuantity = rst!Quantity + AddQty
        rst!Quantity = NewQuantity
        rst.Update
        rst.Close
        
        Set rst = db.OpenRecordset("Journal")
            rst.AddNew
            rst!TDate = DateValue(Now)
            rst!ProductID = ListView1.SelectedItem
            rst!TransactionType = "Purchases"
            rst!TQuantity = AddQty
            rst!TSaleValue = PriceforJournal
            rst!SupplierID = SIDforJournal
            rst!GrossProfit = 0
            MsgBox "The new stock level of item " & ListView1.SelectedItem & " is " & NewQuantity & ".", vbOKOnly, "Stock updated"
            rst.Update
            rst.Close
            Set rst = Nothing
            Set db = Nothing
    End If
Call PopulateListView
End Sub

Private Sub Form_Activate()
'Reload the listview
Call PopulateListView
End Sub

Private Sub SalesButton_Click()
'Load the Sale of Items form
Load FormSales
FormSales.Show
End Sub

Private Sub SQLSearch()
On Error GoTo ErrorHandling
'Get search query from search box
Y = SearchQuery.Text
'Initialise FoundFlag
Let FoundFlag = 0
'Clear Listview items
ListView1.ListItems.Clear

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the selected criteria from database
Query = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID LIKE '*" & Y & "*' or Inventory.PName LIKE '*" & Y & "*' or Inventory.Description LIKE '*" & Y & "*' or Inventory.Category LIKE '*" & Y & "*' or Inventory.Color LIKE '*" & Y & "*' or Inventory.Supplier LIKE '*" & Y & "*'"
'MsgBox "Debug message: " & Query, vbOKOnly '***For debug purposes only

'Open the result set as defined by the query above
Set BLA = db.OpenRecordset(Query)
    
    With BLA
        Do While Not .EOF
            'Load the first column of list view with first column of current pointed record
            Set StockList = ListView1.ListItems.Add(, , BLA(0))
            
            'For columns 1 to 12
            For i = 1 To 12
                    'For columns 1 to 4, load from database as usual
                    If i = 1 Or i = 2 Or i = 3 Or i = 4 Then
                        StockList.SubItems(i) = BLA(i)
                    
                    'For column 5, perform multiplication on the fly to calculate Total Cost
                    ElseIf i = 5 Then
                        StockList.SubItems(5) = BLA(2) * BLA(4)
                        
                    'For column 7 (Supplier), get supplier name from [Supplier]table based on the SupplierID in [Inventory] table
                    ElseIf i = 7 Then
                        'Form SQL expression to query items matching the selected criteria from database
                        SQL = "Select Supplier.Sname from Supplier where Supplier.SupplierID = " & BLA(6)
                        'Open the result set as defined by the query above
                        Set srst = db.OpenRecordset(SQL)
                        
                        If srst.RecordCount = 0 Then
                            'If no record found, then just display Deleted Supplier
                            StockList.SubItems(7) = "*Deleted Supplier*"
                        Else
                            'Else, just load as normal
                            StockList.SubItems(7) = srst(0)
                        End If
                        
                    Else
                        'Otherwise, just load like normal from the records
                        StockList.SubItems(i) = BLA(i - 1)
                    End If
            Next i
            'Move pointer to next record
            .MoveNext
            'Set FoundFlag to 1 to indicate item found
            Let FoundFlag = "1"
        Loop
    End With

'Unload recordset and database
Set BLA = Nothing
db.Close
Set db = Nothing

If FoundFlag = "0" Then
    'If nothing was found, update status
    Status.Caption = "No items matching criteria '" & Y & "' was found."
ElseIf FoundFlag = "1" Then
    'If found something, update number of items found
    Status.Caption = Inventory.ListView1.ListItems.Count & " item(s) found."
End If
Exit Sub

ErrorHandling:
Select Case Err.Number
    Case 93
        Status.Caption = "Invalid search string."
    Case Else
        Status.Caption = "Unknown error occurred."
End Select


End Sub
Private Sub PopulateListView()
On Error GoTo ErrorHandler

'Initialise ErrorFlag
ErrorFlag = 0

'Populate the list from database
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the Inventory table
Set rst = db.OpenRecordset("Inventory")
'Clear the ListView
ListView1.ListItems.Clear

'Repeat until end of Inventory table
Do Until rst.EOF
    'Load the first column of list view with first column of current pointed record
    Set StockList = ListView1.ListItems.Add(, , rst(0))
    
    'For columns 1 to 12, repeat
    For i = 1 To 12
    
        'For columns 1-4, just load from database
        If i = 1 Or i = 2 Or i = 3 Or i = 4 Then
            StockList.SubItems(i) = rst(i)
            
        ElseIf i = 5 Then
        'For column 5, perform on the fly multiplication for Total Cost
            StockList.SubItems(5) = rst(2) * rst(4)
            
        ElseIf i = 7 Then
        'For supplier column (7), load the supplier name from [Supplier]table based on the SupplierID in [Inventory] table
            
            'Form SQL expression to query items matching the selected criteria from database
            SQL = "Select Supplier.Sname from Supplier where Supplier.SupplierID = " & rst(6)
            'Open the result set as defined by the query above
            Set srst = db.OpenRecordset(SQL)
            
            'Check if supplier still exists
            If srst.RecordCount = 0 Then
                'If not, then just show Deleted Supplier
                StockList.SubItems(7) = "*Deleted Supplier*"
            Else
                'Else, loads his/her name.
                StockList.SubItems(7) = srst(0)
            End If
            
        Else
            'Otherwise, just load like normal from the records
            StockList.SubItems(i) = rst(i - 1)
        End If
        
    Next i
    'Move the pointer to next record
    rst.MoveNext
Loop

'Close the recordset and database.
Set rst = Nothing
db.Close
Set db = Nothing

If ErrorFlag = 0 Then
    'If no errors, then update status.
    Status.Caption = "All items loaded."
End If
Exit Sub

ErrorHandler:
ErrorFlag = 1   'To update status on errors
    Select Case Err.Number
        Case 13 'Empty records or type mismatch
            Status.Caption = "Records loaded but there were some information missing from your database. "
            Resume Next
        Case Else   'Other errors
            Status.Caption = "There were some problems loading records from the database. Please contact the administrator for more info."
            Resume Next
    End Select
End Sub

Private Sub RefreshList_Click()
'Reloads the items list from database.
Call PopulateListView
End Sub

Private Sub AddNew_Click()
'Loads the Add New Item window...
FormAddNew.Show
End Sub

Private Sub ExportList_Click()
'Show the preview window, and set the PreviewFlag as Inv
Load Preview
Preview.Show
PreviewFlag = "Inv"
End Sub


Private Sub DeleteItem_Click()
'Confirm with user on deleting the item
Message = MsgBox("Are you sure you want to delete this item with product ID " & ListView1.SelectedItem & "?", vbExclamation + vbYesNo, "Confirm Delete")
If Message = vbYes Then
    Call DeleteRecord
Else
    'Do nothing
End If
End Sub
Private Sub DeleteRecord()
'Delete the selected item from Listview

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items to be deleted matching the selected criteria from database
    SelectedDelete = "SELECT Inventory.* " & "FROM Inventory " & "WHERE Inventory.ProductID = " & ListView1.SelectedItem & ""
    'Open the result set as defined by the query above
    Set MarkedForDeletion = db.OpenRecordset(SelectedDelete)
    
    'Delete items
    With MarkedForDeletion
        Do While Not .EOF
        .Delete
        .MoveNext
        Loop
    End With
    
    'Close the database
    Set MarkedForDeletion = Nothing
    db.Close
    
    'Remove from listview
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
End Sub

Private Sub EditItem_Click()
'Show the Edit item form to edit the selected item
FormEdit.Show
End Sub

Private Sub Search_Click()
'When search button is clicked
'Check if search box is empty
If SearchQuery.Text <> "" Then
    Call SQLSearch
End If
End Sub

Private Sub SearchQuery_Change()
'If the searchbox is empty, load all records, else Search based on the query
If SearchQuery.Text = "" Then
    Call PopulateListView
Else
    Call SQLSearch
End If
End Sub

Private Sub SearchQuery_GotFocus()
'When the searchbox is focused, change its color
Inventory.SearchQuery.BackColor = vbYellow

'Set the search button as the default button
Search.Default = True
End Sub

Private Sub SearchQuery_LostFocus()
'Reset the background color of Search box
SearchQuery.BackColor = vbWhite
'Removes the default button status of the search button
Search.Default = False
'Clears the load status
Status.Caption = ""
End Sub

