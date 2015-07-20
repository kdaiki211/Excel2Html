VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UI_Excel2Html 
   Caption         =   "Excel2Html"
   ClientHeight    =   4692
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   6480
   OleObjectBlob   =   "UI_Excel2Html.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UI_Excel2Html"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_toHtml_Click()
    Dim res As String
    res = ConvertSelectedRangeToHtml
    
    With tbox_output
        .Text = res
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
