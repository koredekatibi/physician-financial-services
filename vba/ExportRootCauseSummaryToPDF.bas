Attribute VB_Name = "Module_ExportRootCauseSummaryToPDF"
Option Explicit

' Exports the Root_Cause_Dashboard sheet to a dated PDF in the same folder
' as this workbook -- used for the weekly leadership packet.
Sub ExportRootCauseSummaryToPDF()

    Dim ws As Worksheet
    Dim exportPath As String
    Dim fileName As String

    Set ws = ThisWorkbook.Sheets("Root_Cause_Dashboard")

    fileName = "RootCause_Summary_" & Format(Date, "YYYY-MM-DD") & ".pdf"
    exportPath = ThisWorkbook.Path & Application.PathSeparator & fileName

    If ThisWorkbook.Path = "" Then
        MsgBox "Save the workbook to a folder first, then re-run this macro.", vbExclamation
        Exit Sub
    End If

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=exportPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    MsgBox "Exported to: " & exportPath, vbInformation

End Sub
