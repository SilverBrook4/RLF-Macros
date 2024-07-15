VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EnvelopePrintForm 
   Caption         =   "Print Envelope"
   ClientHeight    =   2532
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3372
   OleObjectBlob   =   "EnvelopePrintForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EnvelopePrintForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' this code runs the EnvelopePrintForm

Option Explicit

Public envCheck As Boolean

' no button
Private Sub no_Click()
    envCheck = False
    Me.Hide
End Sub

' yes button
Private Sub yes_Click()
    envCheck = True
    Me.Hide
End Sub
