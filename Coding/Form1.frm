VERSION 5.00
Begin VB.Form FormLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Authentication"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton PWClear 
      Caption         =   "X"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton IDClear 
      Caption         =   "X"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   360
      Width           =   255
   End
   Begin VB.CommandButton LoginButton 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.TextBox PW 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox ID 
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is a security feature that allows the user to login to the system

Dim AQCount As Integer
'AQCount to keep track of number of times the user has tried to login but failed in the current session

Private Sub Form_Load()
'Initialise AQCount
AQCount = 0
'Initialise status to provide hints to user to login
Label3.Caption = "To proceed, enter your user ID and password."
End Sub


Private Sub IDClear_Click()
'Clear UserID
ID.Text = ""
End Sub

Private Sub PWClear_Click()
'Clear Password field
PW.Text = ""
End Sub

Private Sub LoginButton_Click()
'Check if the user has attempted login more than 5 times
If AQCount < 5 Then
    Call CheckLogin
End If
End Sub

Private Sub IncorrectLogin()
'If incorrect login
'Clear forms
        ID.Text = ""
        PW.Text = ""
        Message = MsgBox("Incorrect login ID/Password, please try again.", vbOKOnly, "Login Error")
        ID.SetFocus     'Refocus on UserID
        AQCount = AQCount + 1   'Increment AQCount
        Label3.Caption = "Remaining logins: " & (5 - AQCount) 'Change status to display remaining trials available
        If AQCount >= 5 Then
            'If tried more than 5 times, quit the application
            Message = MsgBox("You have tried more than 5 times, the program will now close.", vbOKOnly, "Login Failure")
            Close
            Unload Me
        End If
End Sub
Private Sub CheckLogin()
Dim ToBeDecrypted As String
'Set path of database
Set db = OpenDatabase(App.Path & "/Database.MDB")
'Form an SQL expression to find if user exists
TargetUser = "SELECT Users.* " & "FROM Users " & "WHERE Users.UserName = '" & ID.Text & "'"
'Open the result set based on the query
Set rst = db.OpenRecordset(TargetUser)

'If the result set is empty, user does not exist and call Incorrectlogin
If rst.EOF And rst.BOF Then
            Call IncorrectLogin
Else
'User exists, check password now
    ToBeDecrypted = rst(2)
    'Load the password hash from database
    'Pass to decrypting function
    Call DecryptPassword(ToBeDecrypted)
        If PW.Text = ToBeDecrypted Then
            'If the decrypted hash in database matches the password inputted by user
            'Load the main window and unload the login form
            Main1.Show
            Unload Me
            'Check if the current user is admin and set priviledge accordingly
            If rst(3) = "Admin" Then
                    SessionUserLevel = 0
            Else
                    SessionUserLevel = 1
            End If
        Else
        'If decrypted hash doesn't match the password inputted, call IncorrectLogin
            Call IncorrectLogin
        End If
End If

'Unload everything
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
        Let Value = Asc(Char) - Asc(Mid(ID.Text, 1, 1)) + 17
        Let Decrypted = Decrypted & Chr(Value)
Next i
Let PW = Decrypted
End Sub
Private Sub CancelButton_Click()
'Ask user if they want to cancel login
Dim Response As Integer
Response = MsgBox("Are you sure you want to cancel login?", vbYesNo, "Cancel")
If Response = vbYes Then
    'If yes, then unload form
    Close
    Unload Me
Else
    'Do nothing
End If
End Sub

