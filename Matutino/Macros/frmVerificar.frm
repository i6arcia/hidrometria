VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVerificar 
   Caption         =   "Verificar Nivel"
   ClientHeight    =   2385
   ClientLeft      =   6045
   ClientTop       =   6375
   ClientWidth     =   7905
   OleObjectBlob   =   "frmVerificar.frx":0000
End
Attribute VB_Name = "frmVerificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private nomEstacion As String
Public respuestaFrm As Boolean

Private Sub btnCorregir_Click()
    CapturaMatutino.respuestaFrm = False
    Unload Me
End Sub

Private Sub btnIgnorar_Click()
    CapturaMatutino.respuestaFrm = True
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim x, y As Integer
    x = Application.Width / 4
    y = Application.Height / 2
    
    Me.Left = (x * 3) - 200
    Me.Top = y - 60
End Sub

Sub loadInfo(clv As String, nivel As String, ultNivel As String, desviacionStd As String)
    nomEstacion = dataBase.getNombreEstacion(clv)
    
    'Modifica variables a mostrar
    lblEst.Caption = nomEstacion
    lblEsc.Caption = Format(nivel, "0.00")
    lblUlt.Caption = ultNivel
    lblS.Caption = "+ - (" & Format(desviacionStd, "0.0000") & ")"
End Sub
