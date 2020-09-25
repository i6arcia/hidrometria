VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnActualizar_Click()
    CapturaMatutino.actualizaHojas
    btnCapturar.Enabled = False
End Sub

Private Sub btnCapturar_Click()
    hidrometrica.tieneEditadosHid
    btnCapturar.Enabled = False
    'btnActualizar_Click
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not CapturaMatutino.editando Then
        'On Error Resume Next
        If Not Intersect(Target, hidrometrica.rgoHidrometicas) Is Nothing Then   'Rango editable
            hidrometrica.modificadoHid (Target.Row)
            btnCapturar.Enabled = True
        ElseIf Not Intersect(Target, Range("B5:H5")) Is Nothing Then
            f = Range("B5").Value
            If IsDate(f) Then
                f = Format(f, "yyyy/mm/dd")
                CapturaMatutino.fecha = f
            End If
            CapturaMatutino.actualizaHojas
        End If
    End If
End Sub
