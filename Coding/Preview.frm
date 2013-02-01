VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Preview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preview"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   14685
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8160
      Top             =   7080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton SaveAs 
      Caption         =   "Save As Text File"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   7080
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   13080
      TabIndex        =   2
      Top             =   7080
      Width           =   1335
   End
   Begin VB.CommandButton PrintButton 
      Caption         =   "Print"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   7080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000E&
      Height          =   6615
      Left            =   240
      ScaleHeight     =   6555
      ScaleWidth      =   14115
      TabIndex        =   0
      Top             =   240
      Width           =   14175
   End
End
Attribute VB_Name = "Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This form generates a print preview before exporting or printing

Dim SaveDir As String
Dim db As Database
Dim rst As Recordset
Dim srst As Recordset

Private Sub Form_Activate()
On Error GoTo ErrorHandler

'Check for the previewflag set as they are passed from other forms
If PreviewFlag = "Inv" Or PreviewFlag = "Jrn" Then
    Call GeneratePreview
Else
    'If no flag are set or wrong previewflags are set, then output error
    Picture1.Print "Error generating preview: invalid preview flag set."
    'Disable buttons
    SaveAs.Enabled = False
    PrintButton.Enabled = False
End If
Exit Sub

ErrorHandler:
Call CommonErrorHandler
End Sub

Private Sub CommonErrorHandler()
 Select Case Err.Number
    Case 482    'Printer error
        MsgBox "Error 482: The print spooler service is not started or printer is not configured properly.", vbOKOnly, "Error"
    Case 32755
        'User cancelled operation, so do nothing
    Case Else
        MsgBox "Unknown error occurred, please contact the developer for more info.", vbOKOnly, "Error"
    End Select
End Sub

Private Sub CancelButton_Click()
'If the user pressed the cancel button
Unload Me
PreviewFlag = ""
End Sub

Private Sub GeneratePreview()
'Generate print previews

If PreviewFlag = "Inv" Then
    'Set path of database
    Set db = OpenDatabase(App.Path & "/Database.MDB")
    'Open the Inventory table
    Set rst = db.OpenRecordset("Inventory")
    
    'Print the columnn headers
    Picture1.Print Tab(0); "PID"; Tab(10); "Product"; Tab(40); "Cost"; Tab(55); "Selling"; Tab(70); "Quantity"; Tab(80); "Description"; Tab(110); _
                "Supplier"; Tab(130); "Category"; Tab(145); "Color"; Tab(155); "Size"; Tab(165); "Gender"; Tab(173); "Date Added"
                
    'Get the real supplier name from the [Supplier] table based on supplier ID from [Inventory] table
    Do While Not rst.EOF
        'Check if there are missing supplier information
        'Form SQL expression to query items matching the selected criteria from database
        SQL = "Select Supplier.Sname from Supplier where Supplier.SupplierID = " & rst(6)
        'Open the result set as defined by the query above
        Set srst = db.OpenRecordset(SQL)
            If srst.RecordCount = 0 Then
                'If the supplier record has already been deleted, set unknown
                SupplierName = "Unknown"
            Else
                SupplierName = srst(0)
            End If
        
        'Print each record
        Picture1.Print Tab(0); rst(0); Tab(10); Left(rst(1), 30); Tab(40); FormatCurrency(rst(2)); Tab(55); FormatCurrency(rst(3)); Tab(70); rst(4); Tab(80); Left(rst(5), 25); Tab(110); _
        Left(SupplierName, 20); Tab(130); Left(rst(7), 14); Tab(145); rst(8); Tab(155); rst(9); Tab(165); rst(10); Tab(173); rst(11)
        'Move pointer to next record
        rst.MoveNext
    Loop
End If

If PreviewFlag = "Jrn" Then
    'Set path of database
    Set db = OpenDatabase(App.Path & "/Database.MDB")
    'Open the Journal table
    Set rst = db.OpenRecordset("Journal")
    
    'Print the columnn headers
    Picture1.Print Tab(0); "Trans.ID"; Tab(10); "Date"; Tab(30); "Prod.ID"; Tab(43); "Type"; Tab(58); "Quantity"; Tab(68); "Trans. Value"; Tab(83); "Gross Profit"; Tab(100); "Supplier ID"
    Do While Not rst.EOF
        'Print each record
        Picture1.Print Tab(0); rst(0); Tab(10); Format(rst(1), "  dd/mm/yyyy"); Tab(30); rst(2); Tab(43); rst(3); Tab(58); rst(4); Tab(68); FormatCurrency(rst(5)); Tab(83); ; FormatCurrency(rst(7)); Tab(100); rst(6)
        'Move pointer to next record
        rst.MoveNext
    Loop
End If

End Sub






Private Sub PrintButton_Click()
On Error GoTo ErrorHandler

'Show the print dialog
CommonDialog1.ShowPrinter

'Check whether did user cancel the print operation
If Err.Number <> 32755 Then

    If PreviewFlag = "Inv" Then 'If the user came from the Inventory form
        'Set path of database
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        'Open Inventory table
        Set rst = db.OpenRecordset("Inventory")
        
        'Print column headers
        Printer.Print Tab(0); "PID"; Tab(10); "Product"; Tab(40); "Cost"; Tab(55); "Selling"; Tab(70); "Quantity"; Tab(80); "Description"; Tab(110); _
                "Supplier"; Tab(130); "Category"; Tab(145); "Color"; Tab(155); "Size"; Tab(165); "Gender"; Tab(173); "Date Added"
                
        Do While Not rst.EOF
            'Check if there are missing supplier information
            'Form SQL expression to query items matching the selected criteria from database
            SQL = "Select Supplier.Sname from Supplier where Supplier.SupplierID = " & rst(6)
            'Open the result set as defined by the query above
            Set srst = db.OpenRecordset(SQL)
                
                If srst.RecordCount = 0 Then
                    'If the supplier record has already been deleted, set unknown
                    SupplierName = "Unknown"
                Else
                    SupplierName = srst(0)
                End If
                
            'Print each record
            Printer.Print Tab(0); rst(0); Tab(10); Left(rst(1), 30); Tab(40); FormatCurrency(rst(2)); Tab(55); FormatCurrency(rst(3)); Tab(70); rst(4); Tab(80); Left(rst(5), 25); Tab(110); _
            Left(SupplierName, 20); Tab(130); Left(rst(7), 14); Tab(145); rst(8); Tab(155); rst(9); Tab(165); rst(10); Tab(173); rst(11)
            'Move pointer to next record
            rst.MoveNext
            
        Loop
        'Ending the print operation
        Printer.EndDoc
        'Notify the user of successfully sending the document to printer
        MsgBox "The inventory database has successfully been sent to the printer!", vbOKOnly, "Print"
    End If

    If PreviewFlag = "Jrn" Then 'If the user came from the Journal form
        'Set path of database
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        'Open journal table
        Set rst = db.OpenRecordset("Journal")
        
        'Print column headers
        Printer.Print Tab(0); "Trans.ID"; Tab(10); "Date"; Tab(30); "Prod.ID"; Tab(43); "Type"; Tab(58); "Quantity"; Tab(68); "Trans. Value"; Tab(83); "Gross Profit"; Tab(100); "Supplier ID"
        
        'Print each record
        Do While Not rst.EOF
            Printer.Print Tab(0); rst(0); Tab(10); Format(rst(1), "  dd/mm/yyyy"); Tab(30); rst(2); Tab(43); rst(3); Tab(58); rst(4); Tab(68); FormatCurrency(rst(5)); Tab(83); FormatCurrency(rst(7)); Tab(100); rst(6)
            'Move pointer to next record
            rst.MoveNext
        Loop
        
        'Ending the print operation
        Printer.EndDoc
        'Notify the user of successfully sending the document to printer
        MsgBox "The transaction report has successfully been sent to the printer!", vbOKOnly, "Print"
    End If
End If

'Unset the previewflag
PreviewFlag = ""

'Unload the current form
Unload Me
Exit Sub

ErrorHandler:
Call CommonErrorHandler
End Sub



Private Sub SaveAs_Click()
On Error GoTo ErrorHandler

Call SaveDialog
'Load the save dialog

'Check if the user cancelled the save dialog or not
If SaveDir <> "" Then
    If PreviewFlag = "Inv" Then     'If the user came from the Inventory form
        Open SaveDir For Append As #1
        'Open the new file for appending
            'Set path of database
            Set db = OpenDatabase(App.Path & "/Database.MDB")
            'Opening the Inventory table
            Set rst = db.OpenRecordset("Inventory")
            'Print column headers
            Print #1, Tab(0); "PID"; Tab(10); "Product"; Tab(40); "Cost"; Tab(55); "Selling"; Tab(70); "Quantity"; Tab(80); "Description"; Tab(110); _
                "Supplier"; Tab(130); "Category"; Tab(145); "Color"; Tab(155); "Size"; Tab(165); "Gender"; Tab(173); "Date Added"
            Do While Not rst.EOF
                'Check if there are missing supplier information
                'Form SQL expression to query items matching the selected criteria from database
                SQL = "Select Supplier.Sname from Supplier where Supplier.SupplierID = " & rst(6)
                'Open the result set as defined by the query above
                Set srst = db.OpenRecordset(SQL)
                    If srst.RecordCount = 0 Then
                        'If the supplier record has already been deleted, set unknown
                        SupplierName = "Unknown"
                    Else
                        SupplierName = srst(0)
                    End If
                'Print each entry
                Print #1, Tab(0); rst(0); Tab(10); Left(rst(1), 30); Tab(40); FormatCurrency(rst(2)); Tab(55); FormatCurrency(rst(3)); Tab(70); rst(4); Tab(80); Left(rst(5), 25); Tab(110); _
                Left(SupplierName, 20); Tab(130); Left(rst(7), 14); Tab(145); rst(8); Tab(155); rst(9); Tab(165); rst(10); Tab(173); rst(11)
                'Move pointer to next record
                rst.MoveNext
            Loop
        'Closing the file
        Close #1
        'Notify the user of successful export
        Message = MsgBox("The list of inventory items has successfully been saved to a text file at " & SaveDir, vbOKOnly, "Save Success!")
    End If

    If PreviewFlag = "Jrn" Then     'If the user came from the Journal form
        Open SaveDir For Append As #1
        'Open the new file for appending
        
        'Set path of database
        Set db = OpenDatabase(App.Path & "/Database.MDB")
        'Opening the Journal table
        Set rst = db.OpenRecordset("Journal")
        
        'Print column headers
        Print #1, Tab(0); "Trans.ID"; Tab(10); "Date"; Tab(30); "Prod.ID"; Tab(43); "Type"; Tab(58); "Quantity"; Tab(68); "Trans. Value"; Tab(83); "Gross Profit"; Tab(100); "Supplier ID"
        Do While Not rst.EOF
            'Print each entry
            Print #1, Tab(0); rst(0); Tab(10); Format(rst(1), "  dd/mm/yyyy"); Tab(30); rst(2); Tab(43); rst(3); Tab(58); rst(4); Tab(68); FormatCurrency(rst(5)); Tab(83); FormatCurrency(rst(7)); Tab(100); rst(6)
            'Move pointer to next record
            rst.MoveNext
        Loop
        'Closing the file
        Close #1
        'Notify the user of successful export
        Message = MsgBox("The journal has successfully been saved to a text file at " & SaveDir, vbOKOnly, "Save Success!")
    End If
End If

'Unset the previewflag
PreviewFlag = ""

'Unload the current form
Unload Me
Exit Sub

ErrorHandler:
Call CommonErrorHandler
End Sub

Private Sub SaveDialog()
'Call the Save Common Dialog
'Define the save dialog title
CommonDialog1.DialogTitle = "Save as File..."

'Define flags for the commondialog
'cdlOFNHideReadOnly - Hides the read only checkbox
'cdlOFNPathMustExist - The user can only input valid paths
'cdlOFNOverwritePrompt - The user will be notified if there is an existing file with same name, and ask if to overwrite it
CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNOverwritePrompt

'Define the file type
CommonDialog1.Filter = "Text Documents(*.txt)|*.txt"

'Display the Save common dialog
CommonDialog1.ShowSave

'After the save dialog is closed, set SaveDir to the selected file path
SaveDir = CommonDialog1.FileName
Exit Sub
End Sub
