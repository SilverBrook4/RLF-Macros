VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyEnvelope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' This class represents an envelope to be sent out
' Written by: Mason Ritchie
' Note: uses the Windows API

' Windows API declaration
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Const VK_RETURN As Long = &HD

Private send_address As String
Const RETURN_ADDRESS As String = "71 Court Square Suite A, vbCrLf Harrisonburg, VA 22801"

' Lets the user select an address
Public Sub SelectAddress()
    MsgBox "Highlight the address then press CTRL + SHIFT (close this to contiue)"
    Do ' checks for ENTER press before continuing
    
        DoEvents ' lets ms word be functional before press
        
            Loop Until GetAsyncKeyState(&H10) <> 0 And GetAsyncKeyState(&H11) <> 0 ' checks for ENTER press
            
    send_address = Selection.Text ' saves highlighted contents
End Sub

' Prints an envelope with gven address
Public Function PrintEnvelope()
    ActivePrinter = "NPI246441 (HP LaserJet M608)"
    ActiveDocument.Envelope.PrintOut address:=send_address
End Function
