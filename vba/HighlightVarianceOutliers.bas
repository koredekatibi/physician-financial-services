Attribute VB_Name = "Module_HighlightVarianceOutliers"
Option Explicit

' Loops through Complex_IF_Variance_Flags and color-codes rows by their
' Variance_Status / Priority_Flag so reviewers can triage visually without
' re-sorting the sheet. Safe to re-run — it clears prior formatting first.
Sub HighlightVarianceOutliers()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim statusVal As String
    Dim priorityVal As String

    Set ws = ThisWorkbook.Sheets("Complex_IF_Variance_Flags")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False

    ' Clear existing highlight fills first
    ws.Range("A2:H" & lastRow).Interior.ColorIndex = xlNone

    For i = 2 To lastRow
        statusVal = ws.Cells(i, "G").Value
        priorityVal = ws.Cells(i, "H").Value

        If priorityVal = "High Priority Review" Then
            ws.Range("A" & i & ":H" & i).Interior.Color = RGB(255, 199, 206) ' red
        ElseIf statusVal = "Underpaid - Escalate" Then
            ws.Range("A" & i & ":H" & i).Interior.Color = RGB(255, 235, 156) ' yellow
        ElseIf statusVal = "Denied - RCA Needed" Then
            ws.Range("A" & i & ":H" & i).Interior.Color = RGB(221, 235, 247) ' light blue
        End If
    Next i

    Application.ScreenUpdating = True
    MsgBox "Highlighted " & (lastRow - 1) & " rows by variance/priority status.", vbInformation

End Sub
