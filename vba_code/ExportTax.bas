Attribute VB_Name = "ExportTax"

Public Sub ExportToCompuTax()
    Dim savePath As Variant
    Dim txSheet As Worksheet
    Dim newWb As Workbook
    Dim lastRow As Long
    Dim i As Long, destRow As Long
    
    On Error GoTo ErrorHandler
    
    Set txSheet = ThisWorkbook.Sheets("TX")
    
    ' Select export path
    savePath = Application.GetSaveAsFilename("CompuTax_Import_Data.xlsx", "Excel Files (*.xlsx), *.xlsx", , "Export CompuTax Data")
    
    If savePath = False Then
        MsgBox "Export cancelled by user.", vbExclamation, "Export Cancelled"
        Exit Sub
    End If
    
    ' Create new workbook
    Set newWb = Workbooks.Add
    
    ' Write headers
    newWb.Sheets(1).Cells(1, 1).Value = "Field Name"
    newWb.Sheets(1).Cells(1, 2).Value = "Amount CY"
    newWb.Sheets(1).Cells(1, 3).Value = "Amount PY"
    
    ' Copy data rows from TX sheet (starting from row 6)
    lastRow = txSheet.Cells(txSheet.Rows.Count, "A").End(xlUp).Row
    destRow = 2
    
    For i = 6 To lastRow
        If Trim(txSheet.Cells(i, 1).Value) <> "" Then
            newWb.Sheets(1).Cells(destRow, 1).Value = txSheet.Cells(i, 1).Value
            newWb.Sheets(1).Cells(destRow, 2).Value = txSheet.Cells(i, 2).Value
            newWb.Sheets(1).Cells(destRow, 3).Value = txSheet.Cells(i, 3).Value
            destRow = destRow + 1
        End If
    Next i
    
    ' Format columns
    newWb.Sheets(1).Columns("A:C").AutoFit
    
    ' Save and close
    newWb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook
    newWb.Close SaveChanges:=False
    
    MsgBox "CompuTax export file created successfully!", vbInformation, "Export Complete"
    Exit Sub
    
ErrorHandler:
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    MsgBox "An error occurred during CompuTax export: " & Err.Description, vbCritical, "Export Error"
End Sub

Public Sub ExportToWinman()
    Dim savePath As Variant
    Dim txSheet As Worksheet
    Dim newWb As Workbook
    Dim lastRow As Long
    Dim i As Long, destRow As Long
    
    On Error GoTo ErrorHandler
    
    Set txSheet = ThisWorkbook.Sheets("TX")
    
    ' Winman software import often expects a specific CSV layout. We will write to .csv.
    savePath = Application.GetSaveAsFilename("Winman_Import_Data.csv", "CSV Files (*.csv), *.csv", , "Export Winman Data")
    
    If savePath = False Then
        MsgBox "Export cancelled by user.", vbExclamation, "Export Cancelled"
        Exit Sub
    End If
    
    ' Create new workbook
    Set newWb = Workbooks.Add
    
    ' Write headers
    newWb.Sheets(1).Cells(1, 1).Value = "Field Name"
    newWb.Sheets(1).Cells(1, 2).Value = "Amount"
    
    ' Copy data rows from TX sheet (Winman usually imports CY data)
    lastRow = txSheet.Cells(txSheet.Rows.Count, "A").End(xlUp).Row
    destRow = 2
    
    For i = 6 To lastRow
        If Trim(txSheet.Cells(i, 1).Value) <> "" Then
            newWb.Sheets(1).Cells(destRow, 1).Value = txSheet.Cells(i, 1).Value
            newWb.Sheets(1).Cells(destRow, 2).Value = txSheet.Cells(i, 2).Value ' CY value
            destRow = destRow + 1
        End If
    Next i
    
    ' Save as CSV
    newWb.SaveAs Filename:=savePath, FileFormat:=xlCSV
    newWb.Close SaveChanges:=False
    
    MsgBox "Winman export file created successfully!", vbInformation, "Export Complete"
    Exit Sub
    
ErrorHandler:
    If Not newWb Is Nothing Then newWb.Close SaveChanges:=False
    MsgBox "An error occurred during Winman export: " & Err.Description, vbCritical, "Export Error"
End Sub
