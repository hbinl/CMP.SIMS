VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormJournal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transaction Journal"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10245
   Begin MSComctlLib.ListView ListView1 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   8493
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
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   5160
      Width           =   1575
   End
   Begin VB.CommandButton ExportPrint 
      Caption         =   "Export/Print List..."
      Height          =   375
      Left            =   6480
      TabIndex        =   0
      Top             =   5160
      Width           =   1815
   End
End
Attribute VB_Name = "FormJournal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form displays the list of inventory inflow/outflow journal records from database
'and allows the user to print/export the list.

Private Sub Done_Click()
'If user clicks Cancel, unload the form
Unload Me
End Sub

Private Sub ExportPrint_Click()
'When the user clicks the Export/Print Button, show preview window
Preview.Show
'Set the global preview flag to Jrn, to activate modules in FormPreview relating to Journals
PreviewFlag = "Jrn"
End Sub

Private Sub Form_Load()
'Populate the column headers
With ListView1.ColumnHeaders
    .Add 1, , "ID", 500
    .Add 2, , "Transaction Date", 1500
    .Add 3, , "Product ID", 1500
    .Add 4, , "Transaction Type", 2000
    .Add 5, , "Quantity", 1000
    .Add 6, , "Value", 1000
    .Add 7, , "Gross Profit", 1200
    .Add 8, , "Supplier ID", 1000
End With
'Load the journal entries from database
Call LoadLog
End Sub

Private Sub LoadLog()
'Clear the listview
ListView1.ListItems.Clear

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the Journal table
Set rst = db.OpenRecordset("Journal")

'Start to populate the journal entries
Do Until rst.EOF
    'Add each items to listview
    Set JournalList = ListView1.ListItems.Add(, , rst(0))
            For i = 1 To 7
                If rst(i) <> "" Then
                    If i = 5 Or i = 6 Then
                        'If the column is Cost or Gross Profit, format as currency
                        JournalList.SubItems(i) = FormatCurrency(rst(i))
                    Else    'Else just load directly as normal
                        JournalList.SubItems(i) = rst(i)
                    End If
                Else
                    'If there is missing data
                    JournalList.SubItems(i) = "*Null*"
                End If
            Next i
    'Move to next record
    rst.MoveNext
Loop

'Close the database and recordset after finish loading
Set rst = Nothing
db.Close
Set db = Nothing
End Sub
