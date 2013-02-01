VERSION 5.00
Begin VB.MDIForm Main1 
   BackColor       =   &H8000000C&
   Caption         =   "Sense Boutique Inventory Management System"
   ClientHeight    =   9960
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu File_Menu 
      Caption         =   "File"
      Begin VB.Menu File_Logout 
         Caption         =   "Logout..."
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu File_Exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu Inventory_Menu 
      Caption         =   "Inventory"
      Begin VB.Menu Inventory_List 
         Caption         =   "Inventory List..."
         Shortcut        =   ^I
      End
      Begin VB.Menu Inventory_Search 
         Caption         =   "Search for Item..."
         Shortcut        =   {F3}
      End
      Begin VB.Menu S5 
         Caption         =   "-"
      End
      Begin VB.Menu Inventory_Add 
         Caption         =   "Add New Items..."
         Shortcut        =   ^N
      End
      Begin VB.Menu Transaction_Sales 
         Caption         =   "Sale of Item..."
         Shortcut        =   ^S
      End
      Begin VB.Menu S3 
         Caption         =   "-"
      End
      Begin VB.Menu Inventory_Suppliers 
         Caption         =   "Suppliers..."
      End
   End
   Begin VB.Menu Maintenance_Menu 
      Caption         =   "Maintenance"
      Begin VB.Menu Inventory_Categories 
         Caption         =   "Item Categories..."
      End
      Begin VB.Menu Transaction_Report 
         Caption         =   "Transaction Report..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu S4 
         Caption         =   "-"
      End
      Begin VB.Menu Maintenance_Users 
         Caption         =   "Users Management..."
         Shortcut        =   {F8}
      End
      Begin VB.Menu Maintenance_BR 
         Caption         =   "Backup/Restore..."
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu Window_Menu 
      Caption         =   "Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu Help_Menu 
      Caption         =   "Help"
      Begin VB.Menu Help_Quick 
         Caption         =   "Quick Start Guide"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu Help_HelpTopics 
         Caption         =   "Help Topics"
         Shortcut        =   {F1}
      End
      Begin VB.Menu S2 
         Caption         =   "-"
      End
      Begin VB.Menu Help_About 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Main1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Activate()
'Check user priviledge
Call UserCheck

'Show the Inventory list window automatically
Inventory.Show
End Sub


Private Sub UserCheck()
If SessionUserLevel = 0 Then
    'If admin user, show the Maintenance menu
    Maintenance_Menu.Enabled = True
    Maintenance_Menu.Visible = True
ElseIf SessionUserLevel = 1 Then
    'If ordinary user, hide the Maintenance menu
    Main1.Maintenance_Menu.Enabled = False
    Main1.Maintenance_Menu.Visible = False
End If

End Sub

Private Sub File_Exit_Click()
'Confirm exit with user
Message = MsgBox("Are you sure you want to quit?", vbYesNo, "Confirmation")
If Message = vbYes Then
    'Unload everything
    Unload Me
    Unload FormBackup
    Else
    'Do nothing
End If
End Sub

Private Sub File_Logout_Click()
'Confirm logout with user
Message = MsgBox("Are you sure you want to log out?", vbYesNo, "Confirmation")
If Message = vbYes Then
    'Unload everything
    Unload Me
    Unload FormBackup
    'Show the login dialog
    FormLogin.Show
    Else
    'Do nothing
End If
End Sub

Private Sub Help_About_Click()
'Show the About form
FormAbout.Show
End Sub

Private Sub Help_HelpTopics_Click()
'Show the Help form
FormHelp.Show
End Sub

Private Sub Help_Quick_Click()
'Show the Quick Start Guide
FormQuickStart.Show
End Sub

Private Sub Inventory_Add_Click()
'Show the add new items form
FormAddNew.Show
End Sub

Private Sub Inventory_Categories_Click()
'Show the Category management form
FormCategories.Show
End Sub

Private Sub Inventory_List_Click()
'Show the main Inventory list
Inventory.Show
End Sub

Private Sub Inventory_Search_Click()
'Focus on the Search box
Inventory.SearchQuery.SetFocus
End Sub

Private Sub Maintenance_BR_Click()
'Show the backup form
FormBackup.Show
End Sub

Private Sub Maintenance_Users_Click()
'Show the Users Management form
If SessionUserLevel = 0 Then
    FormUsers.Show
End If
End Sub

Private Sub Inventory_Suppliers_Click()
'Show the list of suppliers
FormSuppliers.Show
End Sub
Private Sub Transaction_Report_Click()
'Show the Transaction Journal
FormJournal.Show
End Sub

Private Sub Transaction_Sales_Click()
'Show the Sale of Items form
FormSales.Show
End Sub

