Attribute VB_Name = "ImportTB"

Public Sub ImportTrialBalance()
    Dim fd As Object
    Dim fileSelected As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim destWs As Worksheet
    Dim rCount As Long
    Dim i As Long
    Dim lastRowDest As Long
    
    Dim colCode As Long
    Dim colName As Long
    Dim colDebit As Long
    Dim colCredit As Long
    Dim headerRow As Long
    Dim startRow As Long
    
    Dim yrAns As String
    Dim destCol As Long
    
    On Error GoTo ErrorHandler
    
    Set destWs = ThisWorkbook.Sheets("TB")
    
    ' Ask which year to import into
    yrAns = InputBox("Which year's Books column to import into (2023, 2024, or 2025)?", "Select Import Year", "2025")
    If yrAns = "" Then Exit Sub
    
    If yrAns = "2023" Then
        destCol = 3 ' Column C
    ElseIf yrAns = "2024" Then
        destCol = 6 ' Column F
    ElseIf yrAns = "2025" Then
        destCol = 9 ' Column I
    Else
        MsgBox "Invalid year selected. Defaulting to 2025.", vbInformation, "Select Year"
        destCol = 9
    End If
    
    ' Select File Dialog
    fileSelected = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm; *.csv), *.xls; *.xlsx; *.xlsm; *.csv", , "Select Trial Balance File")
    
    If fileSelected = "False" Then
        MsgBox "No file selected. Operation cancelled.", vbExclamation, "Import Cancelled"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Open source workbook
    Set srcWb = Workbooks.Open(fileSelected, ReadOnly:=True)
    Set srcWs = srcWb.Sheets(1)
    
    ' Smart Column Detection
    colName = 1
    colDebit = 3
    colCredit = 4
    colCode = 0
    headerRow = 1
    
    Dim r As Long, c As Long
    Dim cellVal As String
    For r = 1 To 5
        For c = 1 To 30
            cellVal = LTrim(RTrim(CStr(srcWs.Cells(r, c).Value)))
            If cellVal <> "" Then
                Select Case LCase(cellVal)
                    Case "name", "particulars", "account name", "gl name", "account", "ledger name", "ledger"
                        colName = c
                        headerRow = r
                    Case "dr bal ho closing", "closing debit", "debit", "dr", "dr balance", "debit amount", "closing dr"
                        colDebit = c
                        headerRow = r
                    Case "cr bal ho closing", "closing credit", "credit", "cr", "cr balance", "credit amount", "closing cr"
                        colCredit = c
                        headerRow = r
                    Case "code", "gl code", "account code", "glcode", "ledger code", "account number", "ac code"
                        colCode = c
                        headerRow = r
                End Select
            End If
        Next c
    Next r
    
    If colCode = 0 Then
        If colName > 1 Then
            colCode = 1
        Else
            colCode = colName
        End If
    End If
    
    startRow = headerRow + 1
    rCount = srcWs.Cells(srcWs.Rows.Count, colName).End(xlUp).Row
    If rCount < startRow Then
        MsgBox "Source file contains no data below the header row.", vbCritical, "Import Error"
        srcWb.Close SaveChanges:=False
        Exit Sub
    End If
    
    ' For the selected column, clear existing values in rows 6 to 500
    Dim rowIdx As Long
    For rowIdx = 6 To 500
        destWs.Cells(rowIdx, destCol).Value = 0
    Next rowIdx
    
    Dim glCode As String, glName As String
    Dim drVal As Double, crVal As Double
    Dim netVal As Double
    Dim foundRange As Range
    Dim nextRowDest As Long
    
    For i = startRow To rCount
        glName = Trim(CStr(srcWs.Cells(i, colName).Value))
        If glName <> "" Then
            glCode = Trim(CStr(srcWs.Cells(i, colCode).Value))
            drVal = 0
            crVal = 0
            If IsNumeric(srcWs.Cells(i, colDebit).Value) Then drVal = CDbl(srcWs.Cells(i, colDebit).Value)
            If IsNumeric(srcWs.Cells(i, colCredit).Value) Then crVal = CDbl(srcWs.Cells(i, colCredit).Value)
            netVal = drVal - crVal
            
            ' Find account in destWs
            Set foundRange = Nothing
            If glCode <> "" Then
                Set foundRange = destWs.Columns(1).Find(What:=glCode, LookIn:=xlValues, LookAt:=xlWhole)
            End If
            
            If foundRange Is Nothing Then
                ' Find next available row (looking for empty GL Code starting from row 6)
                nextRowDest = 6
                Do While destWs.Cells(nextRowDest, 1).Value <> "" And destWs.Cells(nextRowDest, 1).Value <> "Total"
                    nextRowDest = nextRowDest + 1
                Loop
                
                ' If it hits "Total", insert a row before Total
                If destWs.Cells(nextRowDest, 1).Value = "Total" Then
                    destWs.Rows(nextRowDest).Insert Shift:=xlDown
                End If
                
                ' Write GL info
                destWs.Cells(nextRowDest, 1).Value = glCode
                destWs.Cells(nextRowDest, 2).Value = glName
                destWs.Cells(nextRowDest, destCol).Value = netVal
                
                ' Formulas for Final columns
                destWs.Cells(nextRowDest, 5).Formula = "=C" & nextRowDest & "+D" & nextRowDest ' 2023 Final
                destWs.Cells(nextRowDest, 8).Formula = "=F" & nextRowDest & "+G" & nextRowDest ' 2024 Final
                destWs.Cells(nextRowDest, 11).Formula = "=I" & nextRowDest & "+J" & nextRowDest ' 2025 Final
                
                ' Formula for Link Code
                destWs.Cells(nextRowDest, 15).Formula = "=IFERROR(INDEX('Chart of Accounts'!$A$4:$A$100, MATCH(N" & nextRowDest & ", 'Chart of Accounts'!$B$4:$B$100, 0)), """")"
                
                ' Style and formatting
                Dim cIdx As Long
                For cIdx = 1 To 15
                    destWs.Cells(nextRowDest, cIdx).Font.Name = "Calibri"
                    destWs.Cells(nextRowDest, cIdx).Font.Size = 10
                    With destWs.Cells(nextRowDest, cIdx).Borders
                        .LineStyle = xlContinuous
                        .Color = RGB(217, 217, 217)
                        .Weight = xlThin
                    End With
                Next cIdx
                
                ' Number formats
                destWs.Cells(nextRowDest, 3).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 4).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 5).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 6).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 7).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 8).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 9).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 10).NumberFormat = "#,##0.00"
                destWs.Cells(nextRowDest, 11).NumberFormat = "#,##0.00"
            Else
                ' Account exists, update balance
                destWs.Cells(foundRange.Row, destCol).Value = netVal
            End If
        End If
    Next i
    
    srcWb.Close SaveChanges:=False
    
    ' Recalculate totals row
    Dim totalRow As Long
    totalRow = 6
    Do While destWs.Cells(totalRow, 1).Value <> "Total" And totalRow < 1000
        totalRow = totalRow + 1
    Loop
    
    If destWs.Cells(totalRow, 1).Value = "Total" Then
        ' Update SUM formulas for totals
        Dim colLetter As String
        For cIdx = 3 To 11
            colLetter = Split(destWs.Cells(1, cIdx).Address, "$")(1)
            destWs.Cells(totalRow, cIdx).Formula = "=SUM(" & colLetter & "6:" & colLetter & (totalRow - 1) & ")"
        Next cIdx
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    ' Rebuild Note sheets dynamically
    Call RebuildAllNotes
    
    MsgBox "Trial Balance imported successfully for year " & yrAns & "!", vbInformation, "Import Complete"
    Exit Sub
    
ErrorHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "An error occurred during import: " & Err.Description, vbCritical, "Error"
End Sub

Public Sub RebuildAllNotes()
    Dim tbWs As Worksheet
    Dim wsN2 As Worksheet
    Dim wsN3 As Worksheet
    Dim wsN4 As Worksheet
    Dim wsN5 As Worksheet
    Dim rowStart As Long
    
    Set tbWs = ThisWorkbook.Sheets("TB")
    Set wsN2 = ThisWorkbook.Sheets("N2")
    Set wsN3 = ThisWorkbook.Sheets("N3")
    Set wsN4 = ThisWorkbook.Sheets("N4")
    Set wsN5 = ThisWorkbook.Sheets("N5")
    
    ' Clear notes sheets from row 4 down to 500
    wsN2.Range("A4:C500").ClearContents
    wsN2.Range("A4:C500").Borders.LineStyle = xlNone
    wsN2.Range("A4:C500").Font.Bold = False
    
    wsN3.Range("A4:C500").ClearContents
    wsN3.Range("A4:C500").Borders.LineStyle = xlNone
    wsN3.Range("A4:C500").Font.Bold = False
    
    wsN4.Range("A4:C500").ClearContents
    wsN4.Range("A4:C500").Borders.LineStyle = xlNone
    wsN4.Range("A4:C500").Font.Bold = False
    
    wsN5.Range("A4:C500").ClearContents
    wsN5.Range("A4:C500").Borders.LineStyle = xlNone
    wsN5.Range("A4:C500").Font.Bold = False
    
    ' N2 - Liabilities
    rowStart = 4
    Call WriteNoteVBA(wsN2, "1.1.1", "Share Capital", 2, rowStart, tbWs)
    Call WriteNoteVBA(wsN2, "1.1.2", "Reserves and Surplus", 3, rowStart, tbWs)
    Call WriteNoteVBA(wsN2, "1.3.2", "Trade Payables", 4, rowStart, tbWs)
    
    ' N3 - Fixed Assets
    rowStart = 4
    Call WriteNoteVBA(wsN3, "2.1.1.1", "Property, Plant and Equipment", 5, rowStart, tbWs)
    
    ' N4 - Assets
    rowStart = 4
    Call WriteNoteVBA(wsN4, "2.2.3", "Trade Receivables", 6, rowStart, tbWs)
    Call WriteNoteVBA(wsN4, "2.2.4", "Cash and Cash Equivalents", 7, rowStart, tbWs)
    
    ' N5 - Income & Expense
    rowStart = 4
    Call WriteNoteVBA(wsN5, "3.1", "Revenue from Operations", 8, rowStart, tbWs)
    Call WriteNoteVBA(wsN5, "4.3", "Employee Benefit Expenses", 9, rowStart, tbWs)
    Call WriteNoteVBA(wsN5, "4.5", "Depreciation and Amortization", 10, rowStart, tbWs)
End Sub

Private Sub WriteNoteVBA(ws As Worksheet, nodeCode As String, nodeName As String, noteNum As Integer, ByRef rowStart As Long, tbWs As Worksheet)
    Dim lastRowTb As Long
    Dim i As Long
    Dim glCode As String
    Dim glName As String
    Dim mappedNode As String
    Dim startDataRow As Long
    Dim endDataRow As Long
    Dim matchingCount As Long
    Dim c As Long
    
    ' Write Note Header
    ws.Cells(rowStart, 1).Value = "Note " & noteNum & ": " & nodeName
    ws.Cells(rowStart, 1).Font.Name = "Calibri"
    ws.Cells(rowStart, 1).Font.Size = 11
    ws.Cells(rowStart, 1).Font.Bold = True
    rowStart = rowStart + 1
    
    ' Table Headers
    ws.Cells(rowStart, 1).Value = "GL Code"
    ws.Cells(rowStart, 2).Value = "GL Particulars"
    ws.Cells(rowStart, 3).Value = "Amount"
    
    For c = 1 To 3
        ws.Cells(rowStart, c).Font.Bold = True
        ws.Cells(rowStart, c).Font.Color = RGB(255, 255, 255)
        ws.Cells(rowStart, c).Interior.Color = RGB(51, 51, 51)
    Next c
    rowStart = rowStart + 1
    
    startDataRow = rowStart
    matchingCount = 0
    
    lastRowTb = tbWs.Cells(tbWs.Rows.Count, 1).End(xlUp).Row
    If lastRowTb >= 6 Then
        For i = 6 To lastRowTb
            ' Skip the Total row
            If tbWs.Cells(i, 1).Value <> "Total" And Trim(CStr(tbWs.Cells(i, 1).Value)) <> "" Then
                glCode = CStr(tbWs.Cells(i, 1).Value)
                glName = CStr(tbWs.Cells(i, 2).Value)
                mappedNode = CStr(tbWs.Cells(i, 15).Value) ' Column O (15) is Node Code
                
                ' Match node code (either exact or child nodes)
                If mappedNode = nodeCode Or Left(mappedNode, Len(nodeCode) + 1) = nodeCode & "." Then
                    ws.Cells(rowStart, 1).Value = glCode
                    ws.Cells(rowStart, 2).Value = glName
                    
                    ' Formula links dynamically to the Trial Balance (column K - 11 - for 2025 Final)
                    ws.Cells(rowStart, 3).Formula = "=IFERROR(VLOOKUP(A" & rowStart & ", 'TB'!$A:$O, 11, FALSE), 0) / BD!$B$8"
                    ws.Cells(rowStart, 3).NumberFormat = "#,##0.00"
                    
                    For c = 1 To 3
                        With ws.Cells(rowStart, c).Borders
                            .LineStyle = xlContinuous
                            .Color = RGB(217, 217, 217)
                            .Weight = xlThin
                        End With
                    Next c
                    
                    rowStart = rowStart + 1
                    matchingCount = matchingCount + 1
                End If
            End If
        Next i
    End If
    
    If matchingCount = 0 Then
        ws.Cells(rowStart, 1).Value = "-"
        ws.Cells(rowStart, 2).Value = "No GL accounts mapped"
        ws.Cells(rowStart, 3).Value = 0
        ws.Cells(rowStart, 3).NumberFormat = "#,##0.00"
        
        For c = 1 To 3
            With ws.Cells(rowStart, c).Borders
                .LineStyle = xlContinuous
                .Color = RGB(217, 217, 217)
                .Weight = xlThin
            End With
        Next c
        rowStart = rowStart + 1
    End If
    
    endDataRow = rowStart - 1
    
    ws.Cells(rowStart, 2).Value = "Total " & nodeName
    ws.Cells(rowStart, 2).Font.Bold = True
    
    ws.Cells(rowStart, 3).Formula = "=SUM(C" & startDataRow & ":C" & endDataRow & ")"
    ws.Cells(rowStart, 3).Font.Bold = True
    ws.Cells(rowStart, 3).NumberFormat = "#,##0.00"
    
    For c = 1 To 3
        With ws.Cells(rowStart, c).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = RGB(0, 0, 0)
            .Weight = xlThin
        End With
        With ws.Cells(rowStart, c).Borders(xlEdgeBottom)
            .LineStyle = xlDouble
            .Color = RGB(0, 0, 0)
            .Weight = xlThick
        End With
    Next c
    
    rowStart = rowStart + 3
End Sub

Public Sub RollForwardYear()
    Dim tbWs As Worksheet
    Dim bdWs As Worksheet
    Dim aeWs As Worksheet
    Dim bsWs As Worksheet
    Dim plWs As Worksheet
    Dim txWs As Worksheet
    Dim lastRowTb As Long
    Dim i As Long
    
    ' Confirm with the user if run interactively
    If Application.UserControl Then
        Dim response As VbMsgBoxResult
        response = MsgBox("Are you sure you want to roll forward the financial statements to the next fiscal year? This will shift CY to PY, PY to PPY, and clear current year books and adjustments.", vbYesNo + vbQuestion, "Fiscal Year Rollover")
        If response <> vbYes Then Exit Sub
    End If
    
    Set tbWs = ThisWorkbook.Sheets("TB")
    Set bdWs = ThisWorkbook.Sheets("BD")
    Set aeWs = ThisWorkbook.Sheets("AE")
    Set bsWs = ThisWorkbook.Sheets("BS")
    Set plWs = ThisWorkbook.Sheets("PL")
    Set txWs = ThisWorkbook.Sheets("TX")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. Shift Trial Balance Columns
    lastRowTb = tbWs.Cells(tbWs.Rows.Count, 1).End(xlUp).Row
    If lastRowTb >= 6 Then
        For i = 6 To lastRowTb
            ' Skip total row
            If tbWs.Cells(i, 1).Value <> "Total" And Trim(CStr(tbWs.Cells(i, 1).Value)) <> "" Then
                ' New FY23 Books (Col C) = Old FY24 Final (Col H)
                tbWs.Cells(i, 3).Value = tbWs.Cells(i, 8).Value
                tbWs.Cells(i, 4).Value = 0 ' New FY23 Adjustments (Col D)
                
                ' New FY24 Books (Col F) = Old FY25 Final (Col K)
                tbWs.Cells(i, 6).Value = tbWs.Cells(i, 11).Value
                tbWs.Cells(i, 7).Value = 0 ' New FY24 Adjustments (Col G)
                
                ' Clear New FY25 Books (Col I)
                tbWs.Cells(i, 9).Value = 0
                tbWs.Cells(i, 10).Value = "=SUMIFS('AE'!$C$6:$C$55, 'AE'!$A$6:$A$55, A" & i & ") - SUMIFS('AE'!$D$6:$D$55, 'AE'!$A$6:$A$55, A" & i & ")"
            End If
        Next i
    End If
    
    ' 2. Clear Adjustments sheet
    Dim lastRowAe As Long
    lastRowAe = aeWs.Cells(aeWs.Rows.Count, 1).End(xlUp).Row
    If lastRowAe >= 6 Then
        aeWs.Range("A6:A" & lastRowAe).ClearContents
        aeWs.Range("C6:D" & lastRowAe).ClearContents
        aeWs.Range("E6:E" & lastRowAe).ClearContents
    End If
    
    ' 3. Parse and Increment Financial Year in BD
    Dim currentFy As String
    Dim nextFy As String
    currentFy = bdWs.Range("B5").Value ' e.g. "FY 2025-26"
    If Left(currentFy, 3) = "FY " Then
        Dim yearParts() As String
        yearParts = Split(Mid(currentFy, 4), "-")
        If UBound(yearParts) = 1 Then
            Dim startYear As Integer
            Dim endYear As Integer
            startYear = CInt(yearParts(0)) + 1
            endYear = CInt(yearParts(1)) + 1
            nextFy = "FY " & startYear & "-" & Right(CStr(endYear), 2)
            bdWs.Range("B5").Value = nextFy
        End If
    Else
        ' Fallback
        bdWs.Range("B5").Value = "FY 2026-27"
    End If
    
    ' 4. Update TB Date Headers (Row 4)
    Dim yr23 As String, yr24 As String, yr25 As String
    yr23 = tbWs.Range("C4").Value ' "31 March 2023"
    yr24 = tbWs.Range("F4").Value ' "31 March 2024"
    yr25 = tbWs.Range("I4").Value ' "31 March 2025"
    
    tbWs.Range("C4").Value = yr24
    tbWs.Range("F4").Value = yr25
    
    Dim lastYr As Integer
    lastYr = CInt(Right(yr25, 4)) + 1
    tbWs.Range("I4").Value = "31 March " & lastYr
    
    ' 5. Update BS and PL Title Headers and Column Headers
    Dim bsTitle As String
    bsTitle = bsWs.Range("A3").Value
    bsWs.Range("A3").Value = Left(bsTitle, Len(bsTitle) - 4) & lastYr
    bsWs.Range("C5").Value = "31 March " & lastYr
    bsWs.Range("D5").Value = "31 March " & (lastYr - 1)
    
    Dim plTitle As String
    plTitle = plWs.Range("A3").Value
    plWs.Range("A3").Value = Left(plTitle, Len(plTitle) - 4) & lastYr
    plWs.Range("C5").Value = "31 March " & lastYr
    plWs.Range("D5").Value = "31 March " & (lastYr - 1)
    
    Dim txTitle As String
    txTitle = txWs.Range("A3").Value
    txWs.Range("A3").Value = Left(txTitle, Len(txTitle) - 4) & lastYr
    
    ' 6. Rebuild Notes
    Call RebuildAllNotes
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    If Application.UserControl Then
        MsgBox "Rollover completed successfully! The reporting year is now shifted to " & nextFy & ".", vbInformation, "Rollover Complete"
    End If
End Sub
