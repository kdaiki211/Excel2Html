VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Preview 
   Caption         =   "Preview"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13785
   OleObjectBlob   =   "UI_Preview.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public HtmlToPreview As String

Private Sub UserForm_Activate()
    web_main.Navigate "about:blank"

    While web_main.Busy = True
    Wend

    web_main.Document.Open
    web_main.Document.Write HtmlToPreview
    web_main.Document.Close
End Sub

Private Sub UserForm_Resize()
    
End Sub
