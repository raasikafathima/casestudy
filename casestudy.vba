Sub GenerateReports()
    Dim dataSheet As Worksheet
    Dim reportSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim reportRow As Long
    
    ' Set references to data and report sheets
    Set dataSheet = ThisWorkbook.Sheets("Data") ' Change "Data" to your actual data sheet name
    Set reportSheet = ThisWorkbook.Sheets("Report") ' Change "Report" to your actual report sheet name
    
    ' Clear existing data in report sheet
    reportSheet.Cells.Clear
    
    ' Find the last row of data in the data sheet
    lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Define the data range
    Set dataRange = dataSheet.Range("A2:D" & lastRow) ' Assuming data range is from A2:D(lastRow)
    
    ' Add headers to the report sheet
    With reportSheet
        .Range("A1:D1").Value = Array("Date", "Product", "Quantity", "Amount") ' Adjust headers as needed
    End With
    
    ' Loop through each row of data and populate the report sheet
    reportRow = 2 ' Start from row 2 in the report sheet
    For Each cell In dataRange.Rows
        ' Copy data to report sheet
        reportSheet.Cells(reportRow, 1).Value = dataSheet.Cells(cell.Row, 1).Value ' Assuming date is in column A
        reportSheet.Cells(reportRow, 2).Value = dataSheet.Cells(cell.Row, 2).Value ' Assuming product is in column B
        reportSheet.Cells(reportRow, 3).Value = dataSheet.Cells(cell.Row, 3).Value ' Assuming quantity is in column C
        reportSheet.Cells(reportRow, 4).Value = dataSheet.Cells(cell.Row, 4).Value ' Assuming amount is in column D
        reportRow = reportRow + 1
    Next cell
    
    ' Apply formatting to the report sheet (optional)
    With reportSheet
        .Range("A1:D1").Font.Bold = True
        .Columns("C:D").NumberFormat = "#,##0.00" ' Apply number format to quantity and amount columns
    End With
    
    ' Notify user that report generation is complete
    MsgBox "Reports generated successfully!", vbInformation
End Sub
