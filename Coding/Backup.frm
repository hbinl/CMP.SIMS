VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FormBackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup/Restore..."
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9763
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   706
      TabMaxWidth     =   706
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Backup"
      TabPicture(0)   =   "Backup.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FreeSpace"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Drive1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Dir1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Backup"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Restore"
      TabPicture(1)   =   "Backup.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SelectedLabel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Restore"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Dir2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Drive2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "File2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.FileListBox File2 
         Height          =   2430
         Left            =   -72720
         Pattern         =   "*.MDBKP"
         TabIndex        =   10
         Top             =   1440
         Width           =   2535
      End
      Begin VB.DriveListBox Drive2 
         Height          =   315
         Left            =   -74760
         TabIndex        =   8
         Top             =   960
         Width           =   4575
      End
      Begin VB.DirListBox Dir2 
         Height          =   2340
         Left            =   -74760
         TabIndex        =   7
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CommandButton Restore 
         Caption         =   "Restore"
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
         Left            =   -73200
         TabIndex        =   6
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton Backup 
         Caption         =   "Backup"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   4920
         Width           =   1575
      End
      Begin VB.DirListBox Dir1 
         Height          =   3015
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   4575
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label FreeSpace 
         Alignment       =   2  'Center
         Caption         =   "Free space available:"
         Height          =   255
         Left            =   840
         TabIndex        =   13
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label SelectedLabel 
         Caption         =   "Selected: "
         Height          =   615
         Left            =   -74640
         TabIndex        =   11
         Top             =   4080
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Select Location of backup to restore from:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   3735
      End
      Begin VB.Label Label2 
         Caption         =   "Select a location for backup:"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   720
         Width           =   3495
      End
   End
   Begin VB.Label DBSize 
      Alignment       =   1  'Right Justify
      Caption         =   "FileSize"
      Height          =   255
      Left            =   2640
      TabIndex        =   12
      Top             =   6000
      Width           =   2655
   End
End
Attribute VB_Name = "FormBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form allows the user to backup and restore database.
'It will also be loaded during startup if there is a need for restoring database

Private Sub Form_Load()
On Error GoTo ErrorHandling

'Initialise the initial path selected in the directory view controls
Dir1.Path = App.Path & "\Backup\"
Dir2.Path = App.Path & "\Backup\"

'Call functions to check free space and database size
Call FreeSpaceCheck
Call DatabaseSize
Exit Sub

ErrorHandling:
Select Case Err.Number
    Case 53
        'Silencing the error if database is not found, to allow Restore Database dialog to load during startup
        DBSize.Caption = ""
    Case Else
        MsgBox "Unknown error occurred. Please contact the administrator for more info."
End Select
End Sub
Private Sub DatabaseSize()
'Check for the database file size
FileSize = FileLen(App.Path & "/Database.MDB")

'Convert obtained database file size to other more readable equivalent units by dividing 1024
If FileSize < 1000 Then
    DBSize.Caption = "Database size: " & (FileSize) & " bytes"
ElseIf FileSize > 1000 Then
    DBSize.Caption = "Database size: " & (FileSize / 1024) & " kB"
ElseIf FileSize > 1000000 Then
    DBSize.Caption = "Database size: " & (FileSize / 1024 ^ 2) & " MB"
ElseIf FileSize > 1000000000 Then
    DBSize.Caption = "Database size: " & (FileSize / 1024 ^ 3) & " GB"
End If
End Sub


Private Sub FreeSpaceCheck()
'Setting the variables for file system scripting operations
Dim FileSystem As Scripting.FileSystemObject
Dim DriveList As Drives
Dim Drive As Drive

Set FileSystem = New Scripting.FileSystemObject
Set DriveList = FileSystem.Drives
'Refreshing the list of drives currently connected to the system

'Taking the first character of the drive selected by user in the list, to obtain the Drive Letter
DriveLetter = Left(Drive1.Drive, 1)

'Looking through the list of drives connected to the system as defined by DriveList
'If the name of drive matches the one selected by user
'Display available space
'Convert bytes to equivalent more readable units
For Each Drive In DriveList
  If Drive.DriveLetter = UCase(DriveLetter) Then
    If Drive.AvailableSpace > 1000000000000# Then
        FreeSpace.Caption = "Free space available: " & Format(Drive.AvailableSpace / (1024 ^ 4), "###.00") & " TB"
    ElseIf Drive.AvailableSpace > 1000000000 Then
        FreeSpace.Caption = "Free space available: " & Format(Drive.AvailableSpace / (1024 ^ 3), "###.00") & " GB"
    ElseIf Drive.AvailableSpace > 1000000 Then
        FreeSpace.Caption = "Free space available: " & Format(Drive.AvailableSpace / (1024 ^ 2), "###.00") & " MB"
    ElseIf Drive.AvailableSpace > 1000 Then
        FreeSpace.Caption = "Free space available: " & Format(Drive.AvailableSpace / (1024 ^ 1), "###.00") & " kB"
    Else
        FreeSpace.Caption = "Free space available: " & Drive.AvailableSpace & " bytes"
    End If
  End If
Next
End Sub

Private Sub Drive1_Change()
'When the user selects a different drive from the list
On Error GoTo ErrorHandling

'Set path to the drive selected
Dir1.Path = Drive1.Drive
'Check for free space available on the drive
Call FreeSpaceCheck
Exit Sub

ErrorHandling:
    Select Case Err.Number
        Case 68
            'If the drive is not ready, not responding, or not connected
            FreeSpace.Caption = "Drive not connected or not available."
        Case Else
            Resume Next
    End Select
End Sub
Private Sub Backup_Click()
Dim DBKPCount As Integer, BKName As String
'Set path of database
Set Dbase = OpenDatabase(App.Path & "\Database.MDB")
'Check if database can be accessed.
Dbase.Close

'Calling the CompactDatabase command to Compact and Repair database
'Removing redundant entries and reduce file size
DBEngine.CompactDatabase App.Path & "\Database.MDB", App.Path & "\Database2.MDB"

'Delete the old database (before compaction)
Kill (App.Path & "\Database.MDB")

'Renaming the newly compacted database to become the new database
Name App.Path & "\Database2.MDB" As App.Path & "\Database.MDB"

'Ask for user confirmation on database restore
ConfirmBackup = MsgBox("This will start the backup process.", vbQuestion + vbOKCancel, "Confirmation")

'If yes, start the backup process
'The name of the backup is automatically generated from the current date and time
'Copy the database to the user selected directory and rename it to the *Name of Backup*
'Display a dialog box to confirm that backup operation has completed
If ConfirmBackup = vbOK Then
        BKName = "\Database" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & "-" & Hour(Now) & Minute(Now) & Second(Now) & ".mdbkp"
        FileCopy App.Path & "\Database.MDB", Dir1.Path & BKName
        MsgBox "Backup completed!", vbInformation, "Backup"
End If
    
End Sub

Private Sub Drive2_Change()
On Error GoTo ErrorHandling

'Set path displayed in Dir2
Dir2.Path = Drive2.Drive

Exit Sub

ErrorHandling:
    Select Case Err.Number
        Case 68
            'If the drive is not ready, not responding, or not connected
            SelectedLabel.Caption = "The drive is not connected or unavailable to access."
        Case Else
            Resume Next
    End Select
End Sub

Private Sub Dir2_Change()
'Set file paths to be displayed in File2
File2.Path = Dir2.Path
End Sub

Private Sub File2_Click()
'Obtain filepath and name from currently selected file
FileName = File2.Path & "\" & File2.FileName

'Show the currently selected file
SelectedLabel.Caption = "Selected: " & FileName
End Sub



Private Sub Restore_Click()
On Error Resume Next

ConfirmRestore = MsgBox("Are you sure you want to restore this database?", vbQuestion + vbYesNo, "Confirmation")
'Ask for user confirmation on database restore

'If yes, start restoration process

If ConfirmRestore = vbYes Then
    'Backing up the current database
    FileCopy App.Path & "/Database.MDB", App.Path & "/Backup/BackupBeforeRestore.mdbkp"
    'Copy user selected backup to main location
    FileCopy FileName, App.Path & "/Database.MDB"
    
    'Notify the user of successful completion
    MsgBox "Restoration completed!", vbInformation, "Restore Complete"
    MsgBox "Database have been restored, the program is now restarting.", vbInformation + vbOKOnly, "Restore Complete"
    
    'Unload everything and restarts the program
    Unload Me
    Unload Main1
    Load FormSplash
    FormSplash.Show
    Unload Me
End If

End Sub

Private Sub Cancel_Click()
'Unload the current form
Unload Me
End Sub

