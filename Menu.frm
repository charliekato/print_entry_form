VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Menu"
   ClientHeight    =   3870
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   12270
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButtonPrintGo_Click()

        Call print_kohyou.go_ahead
    
End Sub
