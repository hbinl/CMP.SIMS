VERSION 5.00
Begin VB.Form UserPassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password..."
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox ConfirmPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton SavePW 
      Caption         =   "Save"
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2880
      Width           =   1215
   End
   Begin VB.TextBox NewPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox OldPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Caption         =   "Status"
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2280
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm Password:"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label NEWPWLABEL 
      Caption         =   "New Password:"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Old Password:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label LabelUser 
      Caption         =   "User Name: "
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "UserPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the user to change password, and is accessed from the Users Management window
Dim SelectedUser As String
Dim TargetedUserID As Integer

Private Sub Form_Load()
'Set selected user based on the selected user in FormUser's listview
SelectedUser = FormUsers.ListView1.SelectedItem
'Display selected user
LabelUser.Caption = "User Name:        " & SelectedUser
'Provide hint to user
Status.Caption = "For security, it is recommended that the password be at a length of minimum 6 characters."

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the selected criteria from database
TargetUser = "SELECT Users.* " & "FROM Users " & "WHERE Users.UserName = '" & SelectedUser & "'"
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(TargetUser)
'Get TargetedUserID from database
TargetedUserID = rst(0)

'Close database
Set rst = Nothing
db.Close
Set db = Nothing
End Sub
Private Sub Cancel_Click()
'The user presses cancel
Unload Me
End Sub

Private Sub SavePW_Click()
'The user presses save
Call PWCheck
End Sub

Private Sub PWCheck()
'Check for password validity

If Len(NewPW.Text) >= 6 Then    'Check if 6 characters or more
    If OldPW.Text = "" Or NewPW.Text = "" Or ConfirmPW.Text = "" Then   'Check if blank
        Status.Caption = "Please fill in the blank textboxes before continuing."
            'Highlight respective boxes if they are blank
            If OldPW.Text = "" Then
                OldPW.BackColor = vbYellow
            End If
            If NewPW.Text = "" Then
                NewPW.BackColor = vbYellow
            End If
            If ConfirmPW.Text = "" Then
                ConfirmPW.BackColor = vbYellow
            End If
        Else
            If NewPW.Text = OldPW.Text Then 'Check if same as old password
                Status.Caption = "New password cannot be the same as the old one."
                Else
                    If NewPW.Text <> ConfirmPW.Text Then    'Check if password is the same as confirm password
                        Status.Caption = "The confirm password does not match the first one."
                    Else    'Pass all tests, now verifying old password
                        Call CheckOldPassword
                    End If
            End If
        End If
    Else
        Status.Caption = "The new password must be more than 6 characters."
End If

End Sub
Private Sub CheckOldPassword()
Dim ToBeDecrypted As String

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the selected criteria from database
TargetUser = "SELECT Users.* " & "FROM Users " & "WHERE Users.UserID = " & TargetedUserID & ""
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(TargetUser)

'Load old password from database
ToBeDecrypted = rst(2)
'Decrypt old password
Call DecryptPassword(ToBeDecrypted)
'Compare new password with old password
'If they matches, then can proceed to change password
If OldPW.Text = ToBeDecrypted Then
    Call ChangePassword
Else
    'Else, block the password changing attempt.
    Status.Caption = "Incorrect old password."
End If

'Close the record and database
rst.Close
Set rst = Nothing
db.Close
Set db = Nothing
End Sub





Private Sub DecryptPassword(ByRef PW As String)
'Decrypt the encrypted hash from database
Dim Decrypted As String
'Initialise Decrypted variable
Let Decrypted = ""

'For each of the characters in the hash,
'Take it
'Convert the taken character to its ASCII code equivalent
'Then take the first letter of the username and convert it to ASCII code equivalent
'Subtract the ASCII code of username from the ASCII code of the taken character, then add 17
'Final value is converted back to character and placed into Decrypted variable
'Reiterate
'Finally pass back the decrypted password to the login dialog for comparison
For i = 1 To Len(PW)
        Let Char = Mid(PW, i, 1)
        Let Value = Asc(Char) - Asc(Mid(SelectedUser, 1, 1)) + 17
        Let Decrypted = Decrypted & Chr(Value)
Next i
Let PW = Decrypted
End Sub

Private Sub EncryptPassword(ByRef PW As String)
'Encrypt the password into a hash
Dim Encrypted As String
'Initialise Encrypted variable
Let Encrypted = ""

'For each of the characters in the hash,
'Take it
'Convert the taken character to its ASCII code equivalent
'Then take the first letter of the username and convert it to ASCII code equivalent
'Add the ASCII code of username to the ASCII code of the taken character, then subtract 17
'Final value is converted back to character and placed into Encrypted variable
'Reiterate
'Finally pass back the encrypted password to the SaveUser to be stored as an encrypted hash
For i = 1 To Len(PW)
        Let Char = Mid(PW, i, 1)
        Let Value = Asc(Char) + Asc(Mid(SelectedUser, 1, 1)) - 17
        Let Encrypted = Encrypted & Chr(Value)
Next i
Let PW = Encrypted
End Sub

Private Sub ChangePassword()
Dim PasswordTemp As String
PasswordTemp = NewPW.Text

'Call encryption module to encrypte the password
Call EncryptPassword(PasswordTemp)

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form SQL expression to query items matching the selected criteria from database
TargetUser = "SELECT Users.* " & "FROM Users " & "WHERE Users.UserID = " & TargetedUserID & ""
'Open the result set as defined by the query above
Set rst = db.OpenRecordset(TargetUser)
    'Open the record for edit
    rst.Edit
    'Saving the hashed password to database
    rst!UserPassword = PasswordTemp
    rst.Update
    rst.Close
    MsgBox "Password successfully changed for " & SelectedUser & "!", vbOKOnly, "Success"
'Closing the database, recordset and unload the current form
Set rst = Nothing
db.Close
Set db = Nothing
PasswordTemp = ""
Unload Me
End Sub


Private Sub OldPW_Click()
'Reset background color
OldPW.BackColor = vbWhite
End Sub
Private Sub NewPW_Click()
'Reset background color
NewPW.BackColor = vbWhite
End Sub
Private Sub ConfirmPW_Click()
'Reset background color
ConfirmPW.BackColor = vbWhite
End Sub
