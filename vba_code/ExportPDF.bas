Attribute VB_Name = "ExportPDF"

Public Sub ExportStatementsToPDF()
    Dim savePath As Variant
    Dim currentSheet As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set currentSheet = ActiveSheet
    
    ' Select PDF save location
    savePath = Application.GetSaveAsFilename("Schedule_III_Financial_Statements.pdf", "PDF Files (*.pdf), *.pdf", , "Save Financial Statements PDF")
    
    If savePath = False Then
        MsgBox "Export cancelled by user.", vbExclamation, "Export Cancelled"
        Exit Sub
    End If
    
    ' Select target sheets for compiled statements
    Sheets(Array( _
        "BS", _
        "PL", _
        "CF", _
        "Ageing", _
        "N1", _
        "N2", _
        "N3", _
        "N4", _
        "N5", _
        "N6" _
    )).Select
    
    ' Export selected sheets to a single combined PDF
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=savePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
        
    ' Restore focus to Home
    Sheets("Home").Activate
    
    MsgBox "Financial statements compiled and exported to PDF successfully!", vbInformation, "Export Complete"
    Exit Sub
    
ErrorHandler:
    currentSheet.Activate
    MsgBox "An error occurred during PDF generation: " & Err.Description, vbCritical, "Export Error"
End Sub
