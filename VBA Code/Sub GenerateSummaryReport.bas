Sub GenerateSummaryReport()
    Dim wsReport As Worksheet
    Dim wsTrans As Worksheet
    Dim wsAnalysis As Worksheet
    Dim lastRow As Long
    Dim nextRow As Long
    
    Application.ScreenUpdating = False
    
    Set wsTrans = ThisWorkbook.Sheets("Transaction Data")
    Set wsAnalysis = ThisWorkbook.Sheets("Analysis 1")
    
    ' ===== Create or clear "Sales Summary Report" sheet =====
    On Error Resume Next
    Set wsReport = ThisWorkbook.Sheets("Sales Summary Report")
    On Error GoTo 0
    
    If wsReport Is Nothing Then
        Set wsReport = ThisWorkbook.Sheets.Add
        wsReport.Name = "Sales Summary Report"
    Else
        wsReport.Cells.Clear
    End If
    
    nextRow = 1
    
    ' ===== Report Header =====
    wsReport.Cells(nextRow, 1).Value = "Sales Summary Report"
    wsReport.Cells(nextRow, 1).Font.Bold = True
    wsReport.Cells(nextRow, 1).Font.Size = 16
    nextRow = nextRow + 2
    
    wsReport.Cells(nextRow, 1).Value = "Report Generated:"
    wsReport.Cells(nextRow, 2).Value = Now
    wsReport.Cells(nextRow, 2).NumberFormat = "mm/dd/yyyy h:mm AM/PM" ' <-- 12-hour timestamp format
    nextRow = nextRow + 2
    
    ' ===== Revenue & Profit =====
    wsReport.Cells(nextRow, 1).Value = "Total Revenue:"
    wsReport.Cells(nextRow, 2).Value = wsAnalysis.Range("E8").Value
    wsReport.Cells(nextRow, 1).Font.Bold = True
    
    wsReport.Cells(nextRow + 1, 1).Value = "Total Profit:"
    wsReport.Cells(nextRow + 1, 2).Value = wsAnalysis.Range("F8").Value
    wsReport.Cells(nextRow + 1, 1).Font.Bold = True
    nextRow = nextRow + 3
    
    ' ===== Number of Transactions & Avg Sale =====
    lastRow = wsTrans.Cells(wsTrans.Rows.Count, 1).End(xlUp).Row
    wsReport.Cells(nextRow, 1).Value = "Number of Transactions:"
    wsReport.Cells(nextRow, 2).Value = lastRow - 1 ' assuming headers in row 1
    wsReport.Cells(nextRow, 1).Font.Bold = True
    
    wsReport.Cells(nextRow + 1, 1).Value = "Average Sale Value:"
    wsReport.Cells(nextRow + 1, 2).Value = wsAnalysis.Range("E8").Value / (lastRow - 1)
    wsReport.Cells(nextRow + 1, 1).Font.Bold = True
    nextRow = nextRow + 3
    
    ' ===== Top Product =====
    wsReport.Cells(nextRow, 1).Value = "Top Product"
    wsReport.Cells(nextRow, 1).Font.Bold = True
    wsReport.Cells(nextRow, 2).Value = "Revenue"
    wsReport.Cells(nextRow, 2).Font.Bold = True
    nextRow = nextRow + 1
    wsReport.Cells(nextRow, 1).Value = wsAnalysis.Range("I9").Value
    wsReport.Cells(nextRow, 2).Value = wsAnalysis.Range("J9").Value
    nextRow = nextRow + 2
    
    ' ===== Best Buyer Location =====
    wsReport.Cells(nextRow, 1).Value = "Best Buyer Location"
    wsReport.Cells(nextRow, 1).Font.Bold = True
    wsReport.Cells(nextRow, 2).Value = "Profit"
    wsReport.Cells(nextRow, 2).Font.Bold = True
    nextRow = nextRow + 1
    wsReport.Cells(nextRow, 1).Value = wsAnalysis.Range("I14").Value
    wsReport.Cells(nextRow, 2).Value = wsAnalysis.Range("J14").Value
    nextRow = nextRow + 2
    
    ' ===== Top Salesperson =====
    wsReport.Cells(nextRow, 1).Value = "Top Salesperson"
    wsReport.Cells(nextRow, 1).Font.Bold = True
    wsReport.Cells(nextRow, 2).Value = "Profit"
    wsReport.Cells(nextRow, 2).Font.Bold = True
    nextRow = nextRow + 1
    wsReport.Cells(nextRow, 1).Value = wsAnalysis.Range("I24").Value
    wsReport.Cells(nextRow, 2).Value = wsAnalysis.Range("J24").Value
    nextRow = nextRow + 2
    
    ' ===== Profit by Age Group =====
    wsReport.Cells(nextRow, 1).Value = "Profit by Age Group"
    wsReport.Cells(nextRow, 1).Font.Bold = True
    wsReport.Cells(nextRow, 2).Value = "Profit"
    wsReport.Cells(nextRow, 2).Font.Bold = True
    nextRow = nextRow + 1
    
    wsAnalysis.Range("I29:J33").Copy
    wsReport.Cells(nextRow, 1).PasteSpecial xlPasteValues
    wsReport.Cells(nextRow, 1).PasteSpecial xlPasteFormats
    
    wsReport.Columns.AutoFit
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    
    MsgBox "Sales Summary Report updated successfully!", vbInformation
End Sub


