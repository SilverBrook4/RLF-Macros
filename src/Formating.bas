Attribute VB_Name = "Formating"
' This module contains functions for formatting a page to the RLF letter head
' Written by: Mason Ritchie

Option Explicit

' checks for mutliple pages and sets them up for formatting if there are
Public Function FormatMultPages()
    Dim numPages As Integer
    Dim pageRange As Range
    numPages = ActiveDocument.ComputeStatistics(wdStatisticPages)
    ' checks if there is more than 1 page
    If numPages > 1 Then
        Debug.Print "formatting for mult pages"
        Set pageRange = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=2)
        pageRange.Collapse Direction:=wdCollapseStart
        pageRange.InsertBreak Type:=wdSectionBreakContinuous
        SetMarginsNorm
    End If
End Function

' sets the margins of pages after the first to normal margins
Public Function SetMarginsNorm()
    Dim finalSec As Integer
    Dim sec As Section
    finalSec = ActiveDocument.Sections.Count
    Set sec = ActiveDocument.Sections(finalSec)
    With sec.PageSetup
        .TextColumns.SetCount numColumns:=1
        .TopMargin = 72
        .BottomMargin = 72
        .LeftMargin = 72
        .RightMargin = 72
    End With
End Function
