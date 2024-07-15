VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InputFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class represents the file that is being added to the letter head
' Written by: Mason Ritchie
Public file_name As String
Private file_contents As Range
Private fInput As Document

' accesses the contents of the message doc and returns it
Public Function RetrieveContents()
    Dim contents As Range
    Set fInput = Documents.Open(file_name)
    fInput.Activate
    Set RetrieveContents = fInput.Content
End Function

' closes message document
Public Function CloseDoc()
    fInput.Close SaveChanges:=wdDoNotSaveChanges, OriginalFormat:=wdOriginalDocumentFormat
End Function
