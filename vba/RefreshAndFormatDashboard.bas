Attribute VB_Name = "Module_RefreshAndFormatDashboard"
Option Explicit

' Refreshes all formulas/connections in the workbook, then reapplies standard
' number formatting and column widths on the two reporting tabs.
' Run this after pulling in a new monthly claims extract into Claims_Data.
Sub RefreshAndFormatDashboard()

    Dim wsFinancial As Worksheet
    Dim wsRootCause As Worksheet
    Dim lastCol As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    ' Step 1: force a full recalculation, including any external connections
    ThisWorkbook.RefreshAll
    Application.CalculateFullRebuild

    ' Step 2: reformat Financial_Model
    Set wsFinancial = ThisWorkbook.Sheets("Financial_Model")
    With wsFinancial
        lastCol = .Cells(12, .Columns.Count).End(xlToLeft).Column
        .Range(.Cells(1, 2), .Cells(1, lastCol)).EntireColumn.AutoFit
        .Range("F13:H30").NumberFormat = "$#,##0.00;($#,##0.00);""-"""
    End With

    ' Step 3: reformat Root_Cause_Dashboard
    Set wsRootCause = ThisWorkbook.Sheets("Root_Cause_Dashboard")
    With wsRootCause
        .Columns("B:D").AutoFit
        .Range("D7:D20").NumberFormat = "$#,##0.00;($#,##0.00);""-"""
    End With

    Application.ScreenUpdating = True
    MsgBox "Dashboard refreshed and reformatted.", vbInformation

End Sub
