Attribute VB_Name = "SeedAllValues"
Option Compare Database
Option Explicit

' ==== SeedCategories ====
Sub SeedCategories()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim categories As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    categories = Array("Electronics", "Food", "Office Supplies", "Furniture", "Cleaning Supplies")
    
    'Create a query to reuse for every category
    Set qdf = db.CreateQueryDef("", "INSERT INTO Category (CategoryName) VALUES ([catName])")
    
    On Error GoTo ErrHandler
    
    
    
    For i = 0 To UBound(categories)
        qdf.Parameters("catName") = categories(i)
        qdf.Execute dbFailOnError
    Next i
    
    MsgBox "Categories added successfully!"
    
CleanExit:
    qdf.Close
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error seeding categories: " & Err.Description, vbExclamation
    Resume CleanExit
    
End Sub

' ==== SeedSupplier ====
Sub SeedSuppliers()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim suppliers As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    ' Each row: SupplierName, ContactInfo
    suppliers = Array( _
        Array("Acme Distribution", "orders@acme.com"), _
        Array("Global Parts Co.", "sales@globalparts.com"), _
        Array("Northwind Traders", "info@northwind.com") _
    )
    
    Set qdf = db.CreateQueryDef("", "INSERT INTO Supplier (SupplierName, ContactInfo) VALUES ([supName], [supContact])")
    
    On Error GoTo ErrHandler
    
    For i = 0 To UBound(suppliers)
        qdf.Parameters("supName") = suppliers(i)(0)
        qdf.Parameters("supContact") = suppliers(i)(1)
        qdf.Execute dbFailOnError
    Next i
    
    MsgBox "Suppliers seeded."
    
CleanExit:
    qdf.Close
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error seeding suppliers: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub


' ==== SeedProducts ====
Sub SeedProducts()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim products As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    ' Each row: ProductName, CategoryName (looked up), UnitPrice
    products = Array( _
        Array("Wireless Mouse", "Electronics", 19.99), _
        Array("Printer Paper A4", "Office Supplies", 4.5), _
        Array("Office Chair", "Furniture", 89.99) _
    )
    
    Set qdf = db.CreateQueryDef("", "INSERT INTO Product (ProductName, CategoryID, UnitPrice) VALUES ([proName], [catID], [proPrice])")
    
    On Error GoTo ErrHandler
    
    For i = 0 To UBound(products)
        qdf.Parameters("proName") = products(i)(0)
        qdf.Parameters("catID") = GetCategoryID(db, products(i)(1))
        qdf.Parameters("proPrice") = products(i)(2)
        qdf.Execute dbFailOnError
    Next i
    
    MsgBox "Products seeded."
    
CleanExit:
    qdf.Close
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error seeding products: " & Err.Description, vbExclamation
    Resume CleanExit
End Sub
Function GetCategoryID(db As DAO.Database, catName As String) As Long
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set qdf = db.CreateQueryDef("", "SELECT CategoryID FROM Category  WHERE CategoryName = [catName]")
    qdf.Parameters("catName") = catName
    
    Set rs = qdf.OpenRecordset()
    If Not rs.EOF Then 'if at least one match is found (End of File)
        GetCategoryID = rs!CategoryID
    Else                'if there is no match set errorNumbe´ + 1
        Err.Raise vbObjectError + 1, "GetCategoryID", "Category not found: " & catName 'error Number
    End If
    
    rs.Close
    qdf.Close
    Set rs = Nothing
    Set qdf = Nothing
End Function


' ==== SeedOrders ====
Sub SeedOrders()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim orders As Variant
    Dim i As Integer
    Dim prodID As Long, suppID As Long
    
    Set db = CurrentDb
    ' Each row: ProductName, SupplierName, Quantity, OrderDate
    orders = Array( _
        Array("Wireless Mouse", "Acme Distribution", 50, #9/1/2026#), _
        Array("Office Chair", "Northwind Traders", 10, #9/10/2026#) _
    )
    Set qdf = db.CreateQueryDef("", _
        "INSERT INTO Orders (ProductID, SupplierID, Quantity, OrderDate) VALUES ([ordProdID], [ordSuppID], [ordQty], [ordDate])")
    
    On Error GoTo ErrHandler
    
    For i = 0 To UBound(orders)
        qdf.Parameters("ordProdID") = GetProductID(db, orders(i)(0))
        qdf.Parameters("ordSuppID") = GetSupplierID(db, orders(i)(1))
        qdf.Parameters("ordQty") = orders(i)(2)
        qdf.Parameters("ordDate") = orders(i)(3)
        qdf.Execute dbFailOnError
    Next i
    
    MsgBox "Orders seeded."
    
CleanExit:
    qdf.Close
    Set qdf = Nothing
    Set db = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "Error seeding orders: " & Err.Description, vbExclamation
    Resume CleanExit
    
    
End Sub

Function GetProductID(db As DAO.Database, proName As String) As Long
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set qdf = db.CreateQueryDef("", "SELECT ProductID FROM Product WHERE ProductName = [proName]")
    qdf.Parameters("proName") = proName
    
    Set rs = qdf.OpenRecordset()
    If Not rs.EOF Then
        GetProductID = rs!ProductID
    Else
        Err.Raise vbObjectError + 2, "GetProductID", "Product not found: " & proName 'error Number
    End If
    
    rs.Close
    qdf.Close
    Set rs = Nothing
    Set qdf = Nothing
End Function

Function GetSupplierID(db As DAO.Database, supName As String) As Long
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set qdf = db.CreateQueryDef("", "SELECT SupplierID FROM Supplier WHERE SupplierName = [supName]")
    qdf.Parameters("supName") = supName
    
    Set rs = qdf.OpenRecordset()
    If Not rs.EOF Then
        GetSupplierID = rs!SupplierID
    Else
        Err.Raise vbObjectError + 3, "GetSupplierID", "Supplier not found: " & supName
    End If
    
    rs.Close
    qdf.Close
    Set rs = Nothing
    Set qdf = Nothing
End Function
    
' ==== Master runner ====
Sub SeedAllData()
    If DCount("*", "Category") > 0 Or DCount("*", "Supplier") > 0 Or DCount("*", "Product") > 0 Or DCount("*", "Orders") > 0 Then
            MsgBox "Records already exist"
                    Exit Sub
                        End If
    SeedCategories
    SeedSuppliers
    SeedProducts
    SeedOrders
    MsgBox "All tables seeded successfully!"
End Sub

