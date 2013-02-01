Attribute VB_Name = "Base"
'*********************************************************************************
'*                      SENSE INVENTORY MANAGEMENT SYSTEM                        *
'*********************************************************************************
'Program Title      :       Sense Inventory Management System
'Developer          :       Loh Hao Bin (1201A18902)
'Version Number     :       1.0
'Date               :       03/03/2013
'Language           :       Microsoft Visual Basic 6.0
'Dependencies       :       Microsoft Common Dialog Control 6.0 (SP3)
'                           Microsoft ADO Data Control 6.0 (OLEDB)
'                           Microsoft Internet Controls
'                           Microsoft Tabbed Dialog Controls 6.0
'                           Microsoft Windows Common Controls 6.0 (SP6)
'                           Microsoft DAO 3.6 Object Library
'                           Microsoft Scripting Runtime
'
'Name of Client     :       Sense Boutique Inc.
'Main Features      :       Keeping records of stock
'                           Add, Edit and remove inventory records
'                           Search through records
'                           Manage suppliers records
'                           Produce a stock inflow and outflow journal report
'
'******************************************************************************

Public SessionUserLevel As Integer
'To store the current session user priviledge: Admin or user?

Public PreviewFlag As String
'To store the current session flag for Preview window, in order to determine whether to generate print preview for Inventory or Transaction Journal
