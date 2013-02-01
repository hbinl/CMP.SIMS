VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Categories"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5550
   Begin VB.TextBox CatName 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Delete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3120
      Width           =   1935
   End
   Begin VB.CommandButton Refresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton AddNew 
      Caption         =   "Add New Category..."
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   5741
      View            =   3
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
   Begin VB.Label Label1 
      Caption         =   "Category Name:"
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "FormCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the user to manage categories.

Private Sub Form_Load()
'Adding each column header to the ListView
With ListView1.ColumnHeaders
    .Add 1, , "Category", 2500
End With

'After adding column headers, load categories from database
Call RefreshList
End Sub

Private Sub RefreshList()
'Clear all items from listview
ListView1.ListItems.Clear

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
Set rst = db.OpenRecordset("Category")
'Load the Category recordset table

'Load all category names from database until end of record
Do Until rst.EOF
    Set CategoryList = ListView1.ListItems.Add(, , rst(1))
    'Move the pointer to the next record
    rst.MoveNext
Loop

'Unload the recordset and database
Set rst = Nothing
db.Close
Set db = Nothing
End Sub

Private Sub Refresh_Click()
'Reload categories from database
Call RefreshList
End Sub

Private Sub AddNew_Click()
On Error GoTo ErrorHandling
'Adding new categories to the database

'Check if category name inputted is empty
If CatName <> "" Then
    'Set path of database
    Set db = OpenDatabase(App.Path & "/Database.MDB")
    'Open the category table
    Set rst = db.OpenRecordset("Category")
    'Open a new record to be added
    rst.AddNew
    'Add to record
    rst!Category = CatName
    'Save changes to record
    rst.Update
    'Close the record
    rst.Close
    MsgBox "New category successfully added!", vbOKOnly, "Success!"
    
    'Reload the categories from database
    Call RefreshList
    CatName.Text = ""
Else
'If it's empty, notify the user to input a valid name for category
    MsgBox "Please type in a valid category name."
End If
Exit Sub

ErrorHandling:
Select Case Err.Number
    Case 3163
        MsgBox "Please limit the new category name to less than 30 characters.", vbOKOnly, "Error"
        Exit Sub
    Case Else
        MsgBox "Error occurred:" & Err.Number & ". Please report the error to the developer.", vbOKOnly, "Error"
End Select
End Sub

Private Sub Delete_Click()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")

'Ask if user want to delete the selected category
Confirm = MsgBox("Are you sure you want to remove " & ListView1.SelectedItem & "? ", vbYesNo, "Confirm Delete")

If Confirm = vbYes Then
    'If yes, form an SQL query to look for the selected item from Category table
    SQL = "SELECT Category.* " & "FROM Category " & "WHERE Category.Category = '" & ListView1.SelectedItem & "'"
    'Load the result set based on the query
    Set MarkedForDeletion = db.OpenRecordset(SQL)
    
    'Delete all records that match the queried criteria in the result set
    With MarkedForDeletion
        Do While Not .EOF
        .Delete
        .MoveNext
        Loop
    End With
    'Unload database and resultset
    Set MarkedForDeletion = Nothing
    db.Close
    
    'Removing that item from listview
    ListView1.ListItems.Remove ListView1.SelectedItem.Index
Else
    'Do nothing
End If

'Reload items from database
Call RefreshList
End Sub
