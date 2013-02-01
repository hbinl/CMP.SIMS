VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers Management"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   10545
   Begin MSComctlLib.ListView ListView1 
      Height          =   5055
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8916
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton EditCustomer 
      Caption         =   "Edit Customer..."
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.CommandButton DeleteCustomer 
      Caption         =   "Delete Customer..."
      Height          =   375
      Left            =   8280
      TabIndex        =   1
      Top             =   1560
      Width           =   1935
   End
   Begin VB.CommandButton AddNewCustomer 
      Caption         =   "Add New Customer..."
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "FormCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
