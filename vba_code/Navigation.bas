Attribute VB_Name = "Navigation"

' Macro to navigate to a sheet from the Home dashboard
Public Sub GoToBD()
    Sheets("BD").Activate
End Sub

Public Sub GoToTB()
    Sheets("TB").Activate
End Sub

Public Sub GoToAE()
    Sheets("AE").Activate
End Sub

Public Sub GoToBS()
    Sheets("BS").Activate
End Sub

Public Sub GoToPL()
    Sheets("PL").Activate
End Sub

Public Sub GoToCF()
    Sheets("CF").Activate
End Sub

Public Sub GoToN1()
    Sheets("N1").Activate
End Sub

Public Sub GoToN2()
    Sheets("N2").Activate
End Sub

Public Sub GoToN3()
    Sheets("N3").Activate
End Sub

Public Sub GoToN4()
    Sheets("N4").Activate
End Sub

Public Sub GoToN5()
    Sheets("N5").Activate
End Sub

Public Sub GoToN6()
    Sheets("N6").Activate
End Sub

Public Sub GoToTX()
    Sheets("TX").Activate
End Sub

Public Sub GoToHome()
    Sheets("Home").Activate
End Sub

' Macro to show/hide standard Excel toolbars and headings for an app-like feel
Public Sub ToggleToolbar()
    Application.DisplayFormulaBar = Not Application.DisplayFormulaBar
    ActiveWindow.DisplayHeadings = Not ActiveWindow.DisplayHeadings
    ActiveWindow.DisplayGridLines = Not ActiveWindow.DisplayGridLines
End Sub

' Page-fitting macros for financial statements formatting
Public Sub FitToOnePage()
    On Error GoTo ErrorHandler
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    MsgBox "Page setup adjusted to fit on a single page.", vbInformation, "Page Fit"
    Exit Sub
ErrorHandler:
    MsgBox "Could not adjust page setup: " & Err.Description, vbExclamation, "Page Fit Error"
End Sub

Public Sub FitToTwoPages()
    On Error GoTo ErrorHandler
    With ActiveSheet.PageSetup
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 2
    End With
    MsgBox "Page setup adjusted to fit on two pages.", vbInformation, "Page Fit"
    Exit Sub
ErrorHandler:
    MsgBox "Could not adjust page setup: " & Err.Description, vbExclamation, "Page Fit Error"
End Sub
