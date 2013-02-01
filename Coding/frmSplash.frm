VERSION 5.00
Begin VB.Form FormSplash 
   BackColor       =   &H00800080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   8385
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   3990
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8145
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   7680
         Top             =   3480
      End
      Begin VB.Image Image1 
         Height          =   1695
         Left            =   0
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   8175
      End
      Begin VB.Label LoadStatus 
         Caption         =   "Status"
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
         Left            =   120
         TabIndex        =   4
         Top             =   3600
         Width           =   3855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6870
         TabIndex        =   1
         Top             =   3120
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Inventory Management System"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   480
         TabIndex        =   3
         Top             =   2400
         Width           =   7215
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Sense Boutique"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   2
         Top             =   1920
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FormSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error GoTo FileError

'Display the current version
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & " rev " & App.Revision

'Check database and do startup maintenance
Call CheckDatabase
Call RepairDB

'After finish checking, enable the timer to countdown load
LoadStatus.Caption = "Database checking complete. "
Timer1.Enabled = True
Exit Sub

FileError:
    'If there is somekind of problem with the database
    'Ask the user whether they want to restore from backup
    AskIfRestore = MsgBox("Database.mdb cannot be found or corrupted. Would you like to restore from a backup?", vbQuestion + vbYesNo, "Error")
    If AskIfRestore = vbYes Then
        'If yes, then load limited FormBackup wiht only Restore visible
        Load FormBackup
        FormBackup.Show
        FormBackup.SetFocus
        FormBackup.SSTab1.Tab = 1
        FormBackup.SSTab1.TabVisible(0) = False
    Else
        'If no, then ask if the user would like to create a new one.
        'Then copy the Default database to the application directory
        AskIfReset = MsgBox("The program cannot continue without a database. Would you like to create a new one?", vbExclamation + vbYesNo, "Reset")
        If AskIfReset = vbYes Then
                FileCopy App.Path & "/Backup/Default.MDBDefault", App.Path & "/Database.mdb"
                MsgBox "Successfully created a new database!", vbOKOnly, "Restore Complete"
                MsgBox "Default username: bin, password: 000", vbInformation, "Default Login"
                Call Timer1_Timer
        Else
            'If the user chose No to both, the program will exit.
            MsgBox "No database can be found, the program cannot continue without a database. The program will now exit.", vbCritical + vbOKOnly, "Error"
        End If
    End If
Unload Me
End Sub
Private Sub CheckDatabase()
Set db = OpenDatabase(App.Path & "\Database.MDB") 'Check if Database can be accessed

LoadStatus.Caption = "Checking database..."
'Check if each required tables are available
Set rst = db.OpenRecordset("Inventory")
Set rst = db.OpenRecordset("Supplier")
Set rst = db.OpenRecordset("Category")
Set rst = db.OpenRecordset("Users")
Set rst = db.OpenRecordset("Journal")
Set rst = Nothing
db.Close
End Sub

Private Sub RepairDB()
LoadStatus.Caption = "Compacting Database..."

'Calling the CompactDatabase command to Compact and Repair database
'Removing redundant entries and reduce file size
DBEngine.CompactDatabase App.Path & "\Database.MDB", App.Path & "\Database2.MDB"
'Delete the old database (before compaction)
Kill (App.Path & "\Database.MDB")
'Renaming the newly compacted database to become the new database
Name App.Path & "\Database2.MDB" As App.Path & "\Database.MDB"
End Sub

Private Sub Timer1_Timer()
'Wait for 1 second before showing the login form
Load FormLogin
FormLogin.Show
Unload Me
End Sub
