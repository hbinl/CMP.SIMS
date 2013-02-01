VERSION 5.00
Begin VB.Form FormAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   2610
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      Caption         =   "Version"
      Height          =   225
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   3525
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "Sense Boutique Inventory Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3525
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      Caption         =   "For internal use in Sense Boutique Enterprise. Copyright by Loh Hao Bin (1201A18902)"
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   3885
   End
End
Attribute VB_Name = "FormAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form displays information about the current system

Private Sub Form_Load()
    Me.Caption = "About... " 'Set dialog title bar
    lblTitle.Caption = "Sense Boutique " & vbCrLf & "Inventory Management System" 'Set name of system
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision  'Display version numbers
    lblDescription.Caption = "Developed for internal use in Sense Boutique Inc. " & vbCrLf & "2012-2013 Copyright by Loh Hao Bin (1201A18902)" _
    & vbCrLf & "All rights reserved."   'Display disclaimers
  End Sub



Private Sub OKButton_Click()
Unload Me 'Unload the current form
End Sub
