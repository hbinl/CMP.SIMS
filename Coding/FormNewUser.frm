VERSION 5.00
Begin VB.Form FormNewUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New User..."
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4680
   Begin VB.Frame Frame1 
      Caption         =   "User Group"
      Height          =   735
      Left            =   720
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
      Begin VB.OptionButton OptionUser 
         Caption         =   "User"
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton OptionAdmin 
         Caption         =   "Admin"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.TextBox ConfirmPW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox PW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox UID 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Status 
      Alignment       =   2  'Center
      Caption         =   "Status"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   2880
      Width           =   3255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm Password:"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Password:"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "FormNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form will be called when the admin wants to add a new user to the system.

Dim UserPriviledge As String

Private Sub Cancel_Click()
'If the user clicks cancel, dismiss the form
Unload Me
End Sub

Private Sub Form_Load()
'Provide simple hints to the user on the minimum character requirement
Status.Caption = "Please choose minimum 3 characters for username and 6 characters for password."
End Sub

Private Sub OK_Click()
'The user clicks OK

'Check if Password or UserID is empty
If PW <> "" Or UID <> "" Then
    'If the user group selected is Admin
    If OptionAdmin.Value = True Then
        UserPriviledge = "Admin"
        Call UserNameCheck
    'If the user group selected is user
    ElseIf OptionUser.Value = True Then
        UserPriviledge = "User"
        Call UserNameCheck
    'If no user group is specified
    Else
        Status.Caption = "Please choose a user group above."
    End If
Else
'If password or userID is empty
    Status.Caption = "Please fill in the blanks."
End If
End Sub
Private Sub UserNameCheck()
'Check if username already exists in database

'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Open the Users table
Set rst = db.OpenRecordset("Users")

'Initialise FoundFlag
Let FoundFlag = 0

'Compare records with the user inputted string
'If match, set FoundFlag to 1
Do Until rst.EOF
    If rst(1) = UID.Text Then
            Let FoundFlag = 1
    End If
    rst.MoveNext
Loop
Set rst = Nothing
db.Close
Set db = Nothing

'If FoundFlag is still 0,
'then check if username is more than 3 characters
'or check if the password is more than 6 characters
'or check if the password and confirmed password match
If FoundFlag = 0 Then
    If Len(UID) < 3 Then
        Status.Caption = "Please choose a username with longer than 3 characters."
    Else
        If Len(PW) < 6 Then
            Status.Caption = "Please choose a password with longer than 6 characters."
        Else
            If PW <> ConfirmPW Then
                Status.Caption = "The passwords do not match. Please try again."
            Else
                'Passes all tests and ready to be saved to database
                Call SaveUser
            End If
        End If
    End If
Else
'If FoundFlag is set to 1 then there are already existing users with the same name
Status.Caption = "This username has already been used. Please choose another username."
End If
        
End Sub
Private Sub EncryptPassword(ByRef PW As String)
'Password is passed from SaveUser procedure to be encrypted
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
        Let Value = Asc(Char) + Asc(Mid(UID.Text, 1, 1)) - 17
        Let Encrypted = Encrypted & Chr(Value)
Next i
Let PW = Encrypted
End Sub
Private Sub SaveUser()
On Error GoTo ErrorHandling
Dim PasswordTemp As String
PasswordTemp = PW.Text
'Pass the password to EncryptPassword to be encrypted
Call EncryptPassword(PasswordTemp)

        'Encrypted hash is now being written to database
        'Set path of database
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        'Open Users table
        Set rst = db.OpenRecordset("Users")
        rst.AddNew 'Open Empty Record
        rst!UserName = UID.Text
        rst!UserPassword = PasswordTemp
        rst!UserPriviledge = UserPriviledge
        
        'Save changes and close the database/record
        rst.Update
        rst.Close
        Set rst = Nothing
        db.Close
        Set db = Nothing

'Notify the user of successful creation of new user
'Unload the form
MsgBox "User created successfully!", vbOKOnly, "Success!"
Load FormUsers
Unload Me
Exit Sub

ErrorHandling:
Select Case Err.Number
    Case 3163
        Status.Caption = "The username you entered is too long."
        UID.SetFocus
    Case Else
        MsgBox "An unknown error occurred.", vbOKOnly, "Error"
End Select

End Sub
