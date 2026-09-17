Attribute VB_Name = "CheckDataQuality"
Option Compare Database
Option Explicit

Sub CheckDataQuality()
    Dim dupCount As Long
    Dim suspCount As Long
    
    ' Count how many rows each query found
    dupCount = DCount("*", "qryDuplicateOrders")
    suspCount = DCount("*", "qrySuspiciousQuantities")
    
    ' Tell the user what was found
    MsgBox "Duplicate orders found: " & dupCount
    MsgBox "Suspicious quantities found: " & suspCount
    
    ' Open the duplicate orders query if there are any
    If dupCount > 0 Then
        DoCmd.OpenQuery "qryDuplicateOrders"
    End If
    
    ' Open the suspicious quantities query if there are any
    If suspCount > 0 Then
        DoCmd.OpenQuery "qrySuspiciousQuantities"
    End If
    
    ' If nothing was found at all, let the user know everything looks fine
    If dupCount = 0 And suspCount = 0 Then
        MsgBox "No issues detected."
    End If
End Sub
