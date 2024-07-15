Attribute VB_Name = "Driver"
' This module contains the macros for formatting the RLF
' letter head and printing an envelope
' Written by: Mason Ritchie
' Current Version: 1.0

' Version History:
' V1.0: 7/9/2024 "original release"

Option Explicit

' populates the letter head document with
' the message and formats

Sub FillLetterHead()
    Dim msgFile As New InputFile
    Dim msg As Range
    msgFile.file_name = FindFile
    Set msg = ActiveDocument.Bookmarks("message").Range
    msg.Text = msgFile.RetrieveContents
    msgFile.CloseDoc
    FormatMultPages
End Sub

' Prints an envelope
Sub EnvelopePrint()
    Dim Envelope As New MyEnvelope
    Envelope.SelectAddress
    Envelope.PrintEnvelope
End Sub

' runs FillLetterHead and EnvelopePrint together
Sub AutoNew()
    Dim check As Boolean
    Dim checkForm As EnvelopePrintForm
    Set checkForm = New EnvelopePrintForm
    
    ' runs FillLetterHead
    FillLetterHead
    
    ' runs the form to decide if the user want to
    ' print an envelope
    Load checkForm
    checkForm.Show
    check = checkForm.envCheck
    Unload checkForm
    
    ' determines if envelope should be printed or not
    If check Then
        EnvelopePrint
    Else
        MsgBox "Everything Looks Good!"
    End If
End Sub
