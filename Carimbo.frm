VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Carimbo 
   Caption         =   "Carimbo"
   ClientHeight    =   1515
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   OleObjectBlob   =   "Carimbo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Carimbo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btConfirmar_Click()
    If cdSenha = "Cho1co2la3te!" Then
        Carimbar
        Unload Carimbo
    Else
        MsgBox "Senha incorreta"
        Exit Sub
    End If
End Sub

Private Sub btSair_Click()
    Unload Carimbo
End Sub
