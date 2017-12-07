VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVerificaEsc 
   Caption         =   "Verificar Escala"
   ClientHeight    =   2385
   ClientLeft      =   6045
   ClientTop       =   6375
   ClientWidth     =   7905
   OleObjectBlob   =   "frmVerificaEsc.frx":0000
End
Attribute VB_Name = "frmVerificaEsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private valFil As String

Private Sub btnCorregir_Click()
    Vesp.ErrRes = False
    Unload Me
End Sub

Private Sub btnIgnorar_Click()
    Vesp.ErrRes = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim x, y As Integer
    x = Application.Width / 4
    y = Application.Height / 2
    
    Me.Left = (x * 3) - 200
    Me.Top = y - 60
    
    loadInfo
    
End Sub

Private Function loadInfo()
    Dim fil As Integer
    fil = Vesp.i
    lblEst.Caption = Range("C" & fil)
    lblEsc.Caption = Format(Range("H" & fil), "0.00")
    lblUlt.Caption = Format(Range("L" & fil), "0.00")
    lblS.Caption = "+ - (" & Format(Range("M" & fil), "0.0000") & ")"
End Function
