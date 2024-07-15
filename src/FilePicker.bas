Attribute VB_Name = "FilePicker"
' This module contains functions to find a file
' with the file selector
' Written by: Mason Ritchie

Option Explicit

' this creates a file search window for the user to
' pick a message file from
Public Function FindFile()
    Dim fd As fileDialog
    Dim selectedFile As Variant
    Set fd = Application.fileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Title = "Select a File"
    If fd.Show = -1 Then
        selectedFile = fd.SelectedItems(1)
    End If
    Set fd = Nothing
    FindFile = selectedFile
End Function
