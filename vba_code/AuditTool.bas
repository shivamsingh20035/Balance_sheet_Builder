Attribute VB_Name = "AuditTool"
Public Sub TraceCellToLedgers(ByVal Sh As Object, ByVal Target As Range, ByRef Cancel As Boolean)
    Dim nodeCode As String
    Dim glCode As String
    Dim sourceSheet As String
    Dim tbSheet As Worksheet
    Dim drillSheet As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim writeRow As Long
    Dim foundCount As Long
    Dim cySum As Double, pySum As Double
    Dim c As Integer
    
    ' Only intercept double-clicks on amount columns (Col 3 or 4)
    If Target.Column < 3 Or Target.Column > 4 Then Exit Sub
    
    sourceSheet = Sh.Name
    If sourceSheet <> "BS" And sourceSheet <> "PL" And sourceSheet <> "N2" And sourceSheet <> "N3" And sourceSheet <> "N4" And sourceSheet <> "N5" Then Exit Sub
    
    ' Cancel default Excel edit-in-cell behavior
    Cancel = True
    
    Set tbSheet = ThisWorkbook.Sheets("TB")
    Set drillSheet = ThisWorkbook.Sheets("Drilldown_View")
    
    ' Identify what to trace: Node Code or GL Code
    nodeCode = ""
    glCode = ""
    
    If sourceSheet = "BS" Or sourceSheet = "PL" Then
        nodeCode = Trim(Sh.Cells(Target.Row, 1).Value)
        If nodeCode = "" Then Exit Sub
    Else
        ' Notes sheets: GL Code is in Column A
        glCode = Trim(Sh.Cells(Target.Row, 1).Value)
        If glCode = "" Or glCode = "-" Or glCode = "GL Code" Or InStr(glCode, "Total") > 0 Then Exit Sub
    End If
    
    ' Clear Drilldown_View sheet
    drillSheet.Cells.Clear
    
    ' Setup headers
    drillSheet.Cells.Font.Name = "Calibri"
    drillSheet.Cells(1, 1).Value = "<- Return to " & sourceSheet
    drillSheet.Hyperlinks.Add Anchor:=drillSheet.Cells(1, 1), Address:="", SubAddress:="'" & sourceSheet & "'!" & Target.Address, TextToDisplay:="<- Return to " & sourceSheet
    drillSheet.Cells(1, 1).Font.Bold = True
    drillSheet.Cells(1, 1).Font.Color = RGB(0, 0, 255)
    
    drillSheet.Cells(3, 1).Value = "Audit Drilldown View"
    drillSheet.Cells(3, 1).Font.Bold = True
    drillSheet.Cells(3, 1).Font.Size = 14
    
    If nodeCode <> "" Then
        drillSheet.Cells(4, 1).Value = "Drilldown for Node Code: " & nodeCode
    Else
        drillSheet.Cells(4, 1).Value = "Drilldown for GL Code: " & glCode
    End If
    drillSheet.Cells(4, 1).Font.Italic = True
    drillSheet.Cells(4, 1).Font.Size = 10
    
    ' Table headers
    drillSheet.Cells(6, 1).Value = "GL Code"
    drillSheet.Cells(6, 2).Value = "GL Name"
    drillSheet.Cells(6, 3).Value = "31 March 2025 (CY)"
    drillSheet.Cells(6, 4).Value = "31 March 2024 (PY)"
    drillSheet.Cells(6, 5).Value = "Type"
    drillSheet.Cells(6, 6).Value = "Primary Group"
    drillSheet.Cells(6, 7).Value = "Secondary Group"
    
    For c = 1 To 7
        drillSheet.Cells(6, c).Font.Bold = True
        drillSheet.Cells(6, c).Font.Color = RGB(255, 255, 255)
        drillSheet.Cells(6, c).Interior.Color = RGB(79, 79, 79)
        drillSheet.Cells(6, c).HorizontalAlignment = xlCenter
    Next c
    
    writeRow = 7
    foundCount = 0
    
    ' Scan TB sheet
    lastRow = tbSheet.Cells(tbSheet.Rows.Count, 1).End(xlUp).Row
    
    For r = 6 To lastRow - 1 ' Skip the total row
        Dim matchFound As Boolean
        matchFound = False
        
        If nodeCode <> "" Then
            Dim rowNodeCode As String
            rowNodeCode = Trim(tbSheet.Cells(r, 15).Value) ' Column O (Link)
            If rowNodeCode = nodeCode Or Left(rowNodeCode, Len(nodeCode) + 1) = nodeCode & "." Then
                matchFound = True
            End If
        ElseIf glCode <> "" Then
            Dim rowGlCode As String
            rowGlCode = Trim(tbSheet.Cells(r, 1).Value) ' Column A (GL Code)
            If rowGlCode = glCode Then
                matchFound = True
            End If
        End If
        
        If matchFound Then
            drillSheet.Cells(writeRow, 1).Value = tbSheet.Cells(r, 1).Value ' GL Code
            drillSheet.Cells(writeRow, 2).Value = tbSheet.Cells(r, 2).Value ' GL Name
            
            ' CY Final is Col 11 (K)
            ' PY Final is Col 8 (H)
            drillSheet.Cells(writeRow, 3).Value = tbSheet.Cells(r, 11).Value
            drillSheet.Cells(writeRow, 4).Value = tbSheet.Cells(r, 8).Value
            
            drillSheet.Cells(writeRow, 5).Value = tbSheet.Cells(r, 12).Value ' Type
            drillSheet.Cells(writeRow, 6).Value = tbSheet.Cells(r, 13).Value ' Primary
            drillSheet.Cells(writeRow, 7).Value = tbSheet.Cells(r, 14).Value ' Secondary
            
            ' Apply number formats
            drillSheet.Cells(writeRow, 3).NumberFormat = "#,##0.00"
            drillSheet.Cells(writeRow, 4).NumberFormat = "#,##0.00"
            
            ' Add borders
            For c = 1 To 7
                drillSheet.Cells(writeRow, c).Borders.LineStyle = xlContinuous
                drillSheet.Cells(writeRow, c).Borders.Color = RGB(217, 217, 217)
            Next c
            
            writeRow = writeRow + 1
            foundCount = foundCount + 1
        End If
    Next r
    
    If foundCount = 0 Then
        drillSheet.Cells(writeRow, 1).Value = "-"
        drillSheet.Cells(writeRow, 2).Value = "No matching ledger records found."
        For c = 1 To 7
            drillSheet.Cells(writeRow, c).Borders.LineStyle = xlContinuous
            drillSheet.Cells(writeRow, c).Borders.Color = RGB(217, 217, 217)
        Next c
        writeRow = writeRow + 1
    Else
        ' Add Totals Row
        drillSheet.Cells(writeRow, 2).Value = "Total Mapped Ledgers"
        drillSheet.Cells(writeRow, 2).Font.Bold = True
        
        drillSheet.Cells(writeRow, 3).Value = "=SUM(C7:C" & (writeRow - 1) & ")"
        drillSheet.Cells(writeRow, 3).Font.Bold = True
        drillSheet.Cells(writeRow, 3).NumberFormat = "#,##0.00"
        
        drillSheet.Cells(writeRow, 4).Value = "=SUM(D7:D" & (writeRow - 1) & ")"
        drillSheet.Cells(writeRow, 4).Font.Bold = True
        drillSheet.Cells(writeRow, 4).NumberFormat = "#,##0.00"
        
        For c = 1 To 7
            ' Double bottom border for total
            drillSheet.Cells(writeRow, c).Borders(xlEdgeTop).LineStyle = xlContinuous
            drillSheet.Cells(writeRow, c).Borders(xlEdgeTop).Color = RGB(0, 0, 0)
            drillSheet.Cells(writeRow, c).Borders(xlEdgeBottom).LineStyle = xlDouble
            drillSheet.Cells(writeRow, c).Borders(xlEdgeBottom).Color = RGB(0, 0, 0)
        Next c
    End If
    
    ' Autofit columns
    drillSheet.Columns("A:G").AutoFit
    drillSheet.Columns("A").ColumnWidth = 25 ' Make return link column wider
    
    ' Make sheet visible and select it
    drillSheet.Visible = xlSheetVisible
    drillSheet.Activate
End Sub
