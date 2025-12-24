VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmControle 
   Caption         =   "Mapa10X Controller"
   ClientHeight    =   3585
   ClientLeft      =   30
   ClientTop       =   180
   ClientWidth     =   5505
   OleObjectBlob   =   "frmControle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbVoltar_Click()
    ThisWorkbook.Sheets(M_Config.SH_PAINEL).Activate
    frmControle.Hide
End Sub

