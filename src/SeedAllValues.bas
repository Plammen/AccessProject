Attribute VB_Name = "SeedAllValues"
Option Compare Database
' ==== SeedCategories ====
Sub SeedCategories()
    Dim db As DAO.Database
    Dim categories As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    categories = Array("Electronics", "Food", "Office Supplies", "Furniture", "Cleaning Supplies")
    
    For i = 0 To UBound(categories)
        db.Execute "INSERT INTO Category (CategoryName) VALUES ('" & categories(i) & "')", dbFailOnError
    Next i
    
    MsgBox "Categories added successfully!"
End Sub

' ==== SeedSupplier ====
Sub SeedSuppliers()
    Dim db As DAO.Database
    Dim suppliers As Variant
    Dim i As Integer
    
    Set db = CurrentDb
    ' Each row: SupplierName, ContactInfo
    suppliers = Array( _
        Array("Acme Distribution", "orders@acme.com"), _
        Array("Global Parts Co.", "sales@globalparts.com"), _
        Array("Northwind Traders", "info@northwind.com") _
    )
    
    For i = 0 To UBound(suppliers)
        db.Execute "INSERT INTO Supplier (SupplierName, ContactInfo) VALUES ('" & _
            suppliers(i)(0) & "', '" & suppliers(i)(1) & "')", dbFailOnError
    Next i
    
    MsgBox "Suppliers seeded."
End Sub


' ==== SeedProducts ====
Sub SeedProducts()
    Dim db As DAO.Database
    Dim products As Variant
    Dim i As Integer
    Dim catID As Long
    
    Set db = CurrentDb
    ' Each row: ProductName, CategoryName (looked up), UnitPrice
    products = Array( _
        Array("Wireless Mouse", "Electronics", 19.99), _
        Array("Printer Paper A4", "Office Supplies", 4.5), _
        Array("Office Chair", "Furniture", 89.99) _
    )
    
    For i = 0 To UBound(products)
        catID = DLookup("CategoryID", "Category", "CategoryName='" & products(i)(1) & "'")
        db.Execute "INSERT INTO Product (ProductName, CategoryID, UnitPrice) VALUES ('" & _
            products(i)(0) & "', " & catID & ", " & products(i)(2) & ")", dbFailOnError
    Next i
    
    MsgBox "Products seeded."
End Sub


' ==== SeedOrders ====
Sub SeedOrders()
    Dim db As DAO.Database
    Dim orders As Variant
    Dim i As Integer
    Dim prodID As Long, suppID As Long
    
    Set db = CurrentDb
    ' Each row: ProductName, SupplierName, Quantity, OrderDate
    orders = Array( _
        Array("Wireless Mouse", "Acme Distribution", 50, #9/1/2026#), _
        Array("Office Chair", "Northwind Traders", 10, #9/10/2026#) _
    )
    
    For i = 0 To UBound(orders)
        prodID = DLookup("ProductID", "Product", "ProductName='" & orders(i)(0) & "'")
        suppID = DLookup("SupplierID", "Supplier", "SupplierName='" & orders(i)(1) & "'")
        db.Execute "INSERT INTO Orders (ProductID, SupplierID, Quantity, OrderDate) VALUES (" & _
            prodID & ", " & suppID & ", " & orders(i)(2) & ", #" & orders(i)(3) & "#)", dbFailOnError
    Next i
    
    MsgBox "Orders seeded."
End Sub


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

