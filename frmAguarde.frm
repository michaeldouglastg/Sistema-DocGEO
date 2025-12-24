VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAguarde 
   ClientHeight    =   615
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4755
   OleObjectBlob   =   "frmAguarde.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmAguarde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub AtualizarStatus(Mensagem As String)
    On Error Resume Next
    Me.lblMensagem.Caption = Mensagem
    Me.Repaint
    DoEvents
    On Error GoTo 0
End Sub
