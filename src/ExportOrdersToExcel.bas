Attribute VB_Name = "ExportOrdersToExcel"
Option Compare Database
Option Explicit

Sub ExportOrdersToExcel()
    Dim filePath As String
    
    ' Choose where the file will be saved
    filePath = "C:\Users\" & Environ("Username") & "\Documents\OrdersExport.xlsx"
    
    ' Export the query to Excel
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel12Xml, "qryOrdersExport", filePath, True
    
    MsgBox "Orders exported to: " & filePath
End Sub
