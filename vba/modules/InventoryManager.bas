' Inventory Management System - Main Module
' Advanced Excel VBA module for inventory management
' Author: Shivam Sahu
' Version: 1.0

Option Explicit

' Global variables
Public Const INVENTORY_SHEET As String = "Inventory"
Public Const TRANSACTIONS_SHEET As String = "Transactions"
Public Const SUPPLIERS_SHEET As String = "Suppliers"
Public Const DASHBOARD_SHEET As String = "Dashboard"

' Main inventory management functions
Public Sub UpdateInventoryLevels()
    ' Update inventory levels based on transactions
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim startTime As Double
    
    startTime = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' Update progress
    Call UpdateProgressBar(0, "Updating inventory levels...")
    
    For i = 2 To lastRow
        Call CalculateCurrentStock(i)
        Call CheckReorderPoint(i)
        
        ' Update progress
        If i Mod 10 = 0 Then
            Call UpdateProgressBar((i - 1) / (lastRow - 1) * 100, "Processing row " & i & " of " & lastRow)
        End If
    Next i
    
    ' Refresh dashboard
    Call RefreshInventoryDashboard
    
    ' Update progress
    Call UpdateProgressBar(100, "Inventory update completed!")
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    MsgBox "Inventory levels updated successfully!" & vbCrLf & _
           "Processing time: " & Format(Timer - startTime, "0.00") & " seconds", _
           vbInformation, "Inventory Update Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error updating inventory: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub CalculateCurrentStock(rowNum As Long)
    ' Calculate current stock level for a product
    Dim ws As Worksheet
    Dim productId As String
    Dim currentStock As Long
    Dim initialStock As Long
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    productId = ws.Cells(rowNum, 1).Value
    
    If productId = "" Then Exit Sub
    
    ' Get initial stock
    initialStock = ws.Cells(rowNum, 5).Value
    
    ' Calculate current stock from transactions
    currentStock = GetStockFromTransactions(productId) + initialStock
    
    ' Update current stock
    ws.Cells(rowNum, 6).Value = currentStock
    
    ' Update stock status
    Call UpdateStockStatus(rowNum, currentStock)
    
    ' Update inventory value
    ws.Cells(rowNum, 12).Value = currentStock * ws.Cells(rowNum, 11).Value ' Unit Cost
End Sub

Public Function GetStockFromTransactions(productId As String) As Long
    ' Get stock level from transaction history
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim stock As Long
    
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    stock = 0
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = productId Then
            If ws.Cells(i, 4).Value = "IN" Then
                stock = stock + ws.Cells(i, 5).Value
            ElseIf ws.Cells(i, 4).Value = "OUT" Then
                stock = stock - ws.Cells(i, 5).Value
            End If
        End If
    Next i
    
    GetStockFromTransactions = stock
End Function

Public Sub UpdateStockStatus(rowNum As Long, currentStock As Long)
    ' Update stock status based on current stock level
    Dim ws As Worksheet
    Dim reorderPoint As Long
    Dim maxStock As Long
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    reorderPoint = ws.Cells(rowNum, 8).Value
    maxStock = ws.Cells(rowNum, 9).Value
    
    ' Clear existing formatting
    ws.Cells(rowNum, 7).Interior.ColorIndex = xlNone
    
    ' Update status and formatting
    If currentStock <= 0 Then
        ws.Cells(rowNum, 7).Value = "Out of Stock"
        ws.Cells(rowNum, 7).Interior.Color = RGB(255, 0, 0) ' Red
    ElseIf currentStock <= reorderPoint Then
        ws.Cells(rowNum, 7).Value = "Low Stock"
        ws.Cells(rowNum, 7).Interior.Color = RGB(255, 165, 0) ' Orange
    ElseIf currentStock >= maxStock Then
        ws.Cells(rowNum, 7).Value = "Overstocked"
        ws.Cells(rowNum, 7).Interior.Color = RGB(255, 255, 0) ' Yellow
    Else
        ws.Cells(rowNum, 7).Value = "In Stock"
        ws.Cells(rowNum, 7).Interior.Color = RGB(0, 255, 0) ' Green
    End If
End Sub

Public Sub CheckReorderPoint(rowNum As Long)
    ' Check if product needs reordering
    Dim ws As Worksheet
    Dim currentStock As Long
    Dim reorderPoint As Long
    Dim productName As String
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    currentStock = ws.Cells(rowNum, 6).Value
    reorderPoint = ws.Cells(rowNum, 8).Value
    productName = ws.Cells(rowNum, 2).Value
    
    If currentStock <= reorderPoint And currentStock > 0 Then
        ' Add to reorder list
        Call AddToReorderList(rowNum)
    End If
End Sub

Public Sub AddToReorderList(rowNum As Long)
    ' Add product to reorder list
    Dim ws As Worksheet
    Dim reorderWs As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    Set reorderWs = ThisWorkbook.Worksheets("Reorder List")
    
    ' Check if already in reorder list
    If Not IsInReorderList(ws.Cells(rowNum, 1).Value) Then
        lastRow = reorderWs.Cells(reorderWs.Rows.Count, "A").End(xlUp).Row + 1
        
        ' Add product to reorder list
        reorderWs.Cells(lastRow, 1).Value = ws.Cells(rowNum, 1).Value ' Product ID
        reorderWs.Cells(lastRow, 2).Value = ws.Cells(rowNum, 2).Value ' Product Name
        reorderWs.Cells(lastRow, 3).Value = ws.Cells(rowNum, 6).Value ' Current Stock
        reorderWs.Cells(lastRow, 4).Value = ws.Cells(rowNum, 8).Value ' Reorder Point
        reorderWs.Cells(lastRow, 5).Value = CalculateSuggestedOrderQty(rowNum) ' Suggested Qty
        reorderWs.Cells(lastRow, 6).Value = ws.Cells(rowNum, 10).Value ' Supplier
        reorderWs.Cells(lastRow, 7).Value = Date ' Date Added
    End If
End Sub

Public Function IsInReorderList(productId As String) As Boolean
    ' Check if product is already in reorder list
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Reorder List")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    IsInReorderList = False
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = productId Then
            IsInReorderList = True
            Exit Function
        End If
    Next i
End Function

Public Function CalculateSuggestedOrderQty(rowNum As Long) As Long
    ' Calculate suggested order quantity
    Dim ws As Worksheet
    Dim currentStock As Long
    Dim reorderPoint As Long
    Dim maxStock As Long
    Dim suggestedQty As Long
    Dim avgUsage As Double
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    currentStock = ws.Cells(rowNum, 6).Value
    reorderPoint = ws.Cells(rowNum, 8).Value
    maxStock = ws.Cells(rowNum, 9).Value
    
    ' Calculate average usage
    avgUsage = CalculateAverageUsage(ws.Cells(rowNum, 1).Value)
    
    ' Calculate suggested quantity
    If avgUsage > 0 Then
        ' Based on usage and lead time
        suggestedQty = Int(avgUsage * 30) ' 30 days supply
    Else
        ' Based on max stock and current stock
        suggestedQty = maxStock - currentStock
    End If
    
    ' Ensure minimum order quantity
    If suggestedQty < reorderPoint Then
        suggestedQty = reorderPoint
    End If
    
    ' Ensure we don't exceed max stock
    If currentStock + suggestedQty > maxStock Then
        suggestedQty = maxStock - currentStock
    End If
    
    CalculateSuggestedOrderQty = suggestedQty
End Function

Public Function CalculateAverageUsage(productId As String) As Double
    ' Calculate average daily usage for a product
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim totalUsage As Long
    Dim dayCount As Long
    Dim currentDate As Date
    Dim lastDate As Date
    
    Set ws = ThisWorkbook.Worksheets(TRANSACTIONS_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    totalUsage = 0
    dayCount = 0
    lastDate = Date - 30 ' Look at last 30 days
    
    For i = 2 To lastRow
        If ws.Cells(i, 2).Value = productId And ws.Cells(i, 4).Value = "OUT" Then
            currentDate = ws.Cells(i, 3).Value
            If currentDate >= lastDate Then
                totalUsage = totalUsage + ws.Cells(i, 5).Value
                dayCount = dayCount + 1
            End If
        End If
    Next i
    
    If dayCount > 0 Then
        CalculateAverageUsage = totalUsage / 30 ' Average per day
    Else
        CalculateAverageUsage = 0
    End If
End Function

Public Sub RefreshInventoryDashboard()
    ' Refresh the inventory dashboard
    Dim ws As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    
    ' Update summary statistics
    Call UpdateDashboardStatistics
    
    ' Refresh charts
    Call RefreshDashboardCharts
    
    ' Update alerts
    Call UpdateDashboardAlerts
End Sub

Public Sub UpdateDashboardStatistics()
    ' Update dashboard statistics
    Dim ws As Worksheet
    Dim inventoryWs As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set inventoryWs = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    
    lastRow = inventoryWs.Cells(inventoryWs.Rows.Count, "A").End(xlUp).Row
    
    ' Calculate statistics
    Dim totalProducts As Long
    Dim totalValue As Double
    Dim avgStock As Double
    Dim lowStockItems As Long
    Dim outOfStockItems As Long
    
    totalProducts = lastRow - 1
    totalValue = Application.WorksheetFunction.Sum(inventoryWs.Range("L2:L" & lastRow)) ' Inventory Value
    avgStock = Application.WorksheetFunction.Average(inventoryWs.Range("F2:F" & lastRow)) ' Current Stock
    lowStockItems = Application.WorksheetFunction.CountIf(inventoryWs.Range("G2:G" & lastRow), "Low Stock")
    outOfStockItems = Application.WorksheetFunction.CountIf(inventoryWs.Range("G2:G" & lastRow), "Out of Stock")
    
    ' Update dashboard
    ws.Cells(3, 2).Value = totalProducts
    ws.Cells(4, 2).Value = Format(totalValue, "$#,##0.00")
    ws.Cells(5, 2).Value = Format(avgStock, "0")
    ws.Cells(6, 2).Value = lowStockItems
    ws.Cells(7, 2).Value = outOfStockItems
End Sub

Public Sub UpdateDashboardAlerts()
    ' Update dashboard alerts
    Dim ws As Worksheet
    Dim reorderWs As Worksheet
    Dim lastRow As Long
    
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    Set reorderWs = ThisWorkbook.Worksheets("Reorder List")
    
    lastRow = reorderWs.Cells(reorderWs.Rows.Count, "A").End(xlUp).Row
    
    ' Update reorder alerts
    ws.Cells(10, 2).Value = lastRow - 1
    
    ' Update alert messages
    If lastRow - 1 > 0 Then
        ws.Cells(11, 2).Value = "Items need reordering!"
        ws.Cells(11, 2).Interior.Color = RGB(255, 0, 0)
    Else
        ws.Cells(11, 2).Value = "All items in stock"
        ws.Cells(11, 2).Interior.Color = RGB(0, 255, 0)
    End If
End Sub

Public Sub UpdateProgressBar(percent As Double, message As String)
    ' Update progress bar (if exists)
    On Error Resume Next
    ThisWorkbook.Worksheets("Progress").Cells(1, 1).Value = percent
    ThisWorkbook.Worksheets("Progress").Cells(2, 1).Value = message
    On Error GoTo 0
End Sub

Public Sub RefreshDashboardCharts()
    ' Refresh dashboard charts
    Dim ws As Worksheet
    Dim chart As ChartObject
    
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    
    ' Refresh all charts
    For Each chart In ws.ChartObjects
        chart.Chart.Refresh
    Next chart
End Sub

' Data validation functions
Public Sub ValidateInventoryData()
    ' Validate inventory data integrity
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errors As Collection
    Dim errorMsg As String
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    Set errors = New Collection
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        ' Check for required fields
        If ws.Cells(i, 1).Value = "" Then
            errors.Add "Row " & i & ": Product ID is required"
        End If
        
        If ws.Cells(i, 2).Value = "" Then
            errors.Add "Row " & i & ": Product Name is required"
        End If
        
        If ws.Cells(i, 8).Value <= 0 Then
            errors.Add "Row " & i & ": Reorder Point must be greater than 0"
        End If
        
        If ws.Cells(i, 9).Value <= ws.Cells(i, 8).Value Then
            errors.Add "Row " & i & ": Max Stock must be greater than Reorder Point"
        End If
    Next i
    
    ' Display errors
    If errors.Count > 0 Then
        errorMsg = "Data validation errors found:" & vbCrLf
        For i = 1 To errors.Count
            errorMsg = errorMsg & errors(i) & vbCrLf
        Next i
        MsgBox errorMsg, vbCritical, "Data Validation Errors"
    Else
        MsgBox "Data validation completed successfully!", vbInformation, "Validation Complete"
    End If
End Sub

' Utility functions
Public Sub ClearInventoryData()
    ' Clear all inventory data
    Dim result As VbMsgBoxResult
    
    result = MsgBox("Are you sure you want to clear all inventory data?", _
                   vbYesNo + vbQuestion, "Confirm Clear Data")
    
    If result = vbYes Then
        ThisWorkbook.Worksheets(INVENTORY_SHEET).Range("A2:M1000").Clear
        ThisWorkbook.Worksheets(TRANSACTIONS_SHEET).Range("A2:F1000").Clear
        ThisWorkbook.Worksheets("Reorder List").Range("A2:G1000").Clear
        MsgBox "All inventory data has been cleared.", vbInformation, "Data Cleared"
    End If
End Sub

Public Sub ExportInventoryData()
    ' Export inventory data to CSV
    Dim ws As Worksheet
    Dim fileName As String
    Dim filePath As String
    
    Set ws = ThisWorkbook.Worksheets(INVENTORY_SHEET)
    
    ' Get file path
    fileName = "Inventory_Export_" & Format(Date, "yyyy-mm-dd") & ".csv"
    filePath = ThisWorkbook.Path & "\" & fileName
    
    ' Export data
    ws.Range("A1").CurrentRegion.Copy
    Workbooks.Add
    ActiveSheet.Paste
    ActiveWorkbook.SaveAs filePath, xlCSV
    ActiveWorkbook.Close
    
    MsgBox "Inventory data exported to: " & filePath, vbInformation, "Export Complete"
End Sub
