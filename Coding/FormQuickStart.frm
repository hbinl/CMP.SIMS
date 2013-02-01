VERSION 5.00
Begin VB.Form FormQuickStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Start Guide"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6315
   Begin VB.CommandButton OKButton 
      Caption         =   "Got It!"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "To Backup: Maintenance > Backup/Restore... and choose the output directory."
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   5775
   End
   Begin VB.Label Label5 
      Caption         =   "To manage users: Maintenance > User Management..."
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "To look at daily journals: Log > Journal..."
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   5655
   End
   Begin VB.Label Label3 
      Caption         =   "To Search for items: Inventory > Search for Items... or Press F3"
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "To Add Items: Inventory > Add New Items..."
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Quick Start Guide"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "FormQuickStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form is a simple quick start guide that provides a brief guide to how to use the system.

Private Sub OKButton_Click()
'Unload the current form
Unload Me
End Sub

Private Sub Form_Load()
'Populate each of the labels with these texts
Label2.Caption = "To Add Items: Inventory > Add New Items..."
Label3.Caption = "To Search for items: Inventory > Search for Items... or Press F3"
Label4.Caption = "To look at stocktake journals: Inventory > Transaction Journal..."
Label5.Caption = "To manage users: Maintenance > User Management..."
Label6.Caption = "To Backup: Maintenance > Backup/Restore... and choose the output directory."
End Sub


