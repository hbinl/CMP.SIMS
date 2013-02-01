VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormUsers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Users Management"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   3765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ChangeType 
      Caption         =   "Change Type"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2175
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton ChangePW 
      Caption         =   "Change Password"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton RemoveUser 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton AddUser 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "FormUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rst As Recordset

Private Sub AddUser_Click()
'When the user clicks the + button, call the NewUser form
Load FormNewUser
End Sub

Private Sub ChangePW_Click()
'When the user clicks change password
UserPassword.Show
End Sub

Private Sub ChangeType_Click()
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the selected criteria from database
TargetUser = "SELECT Users.* " & "FROM Users " & "WHERE Users.UserName = '" & ListView1.SelectedItem & "'"
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(TargetUser)
'Open the record for edit
rst.Edit

If rst!UserPriviledge = "Admin" Then
    'if the user is already an admin, change to user
    X = MsgBox("Are you sure you want to change this user's type to 'USER'?", vbYesNo, "Change User Type")
    If X = vbYes Then
        rst!UserPriviledge = "User"
    End If
ElseIf rst!UserPriviledge = "User" Then
    'If the user is already a user, change to admin
    X = MsgBox("Are you sure you want to change this user's type to 'ADMIN'?", vbYesNo, "Change User Type")
    If X = vbYes Then
        rst!UserPriviledge = "Admin"
    End If
End If

'Save changes, close the database and record
rst.Update
rst.Close
Set rst = Nothing
db.Close
Set db = Nothing

'Reload the users list
Call LoadUsers
End Sub

Private Sub Form_Activate()
Call LoadUsers
End Sub

Private Sub Form_Load()
'Populate the listview column headers
With ListView1.ColumnHeaders
    .Add 1, , "User", 2000
    .Add 2, , "Type", 1000
End With
'Populate user lists
Call LoadUsers
End Sub

Private Sub LoadUsers()
'Clear the listview
ListView1.ListItems.Clear
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the users table
Set rst = db.OpenRecordset("Users")

'Add each column and row into the listview
Do Until rst.EOF
    Set UserList = ListView1.ListItems.Add(, , rst(1))
    UserList.SubItems(1) = rst(3)
    'Move the pointer to next record
    rst.MoveNext
Loop

'Close the recordset and database
Set rst = Nothing
db.Close
Set db = Nothing
End Sub

Private Sub RemoveUser_Click()
'Confirm delete with user
DeleteCheck = MsgBox("Are you sure you want to delete this user:" & ListView1.SelectedItem & "?", vbYesNo, "Confirm Delete")
    If DeleteCheck = vbYes Then
        'If yes, set path of database
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        'Form SQL expression to query items matching the selected criteria from database
        SQL = "SELECT Users.* " & "FROM Users " & "WHERE Users.UserName = '" & ListView1.SelectedItem & "'"
        'Open the result set as defined by the query above
        Set MarkedForDeletion = db.OpenRecordset(SQL)
        
        'Delete all records in the result set
        With MarkedForDeletion
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
        End With
        Set MarkedForDeletion = Nothing
        'Close database
        db.Close
        'Remove from listview
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
End Sub
