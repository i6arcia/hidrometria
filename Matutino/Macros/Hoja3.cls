VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnActualizarClmN_Click()
    CapturaMatutino.actualizaHojas
    btnCapturarClmN.Enabled = False
End Sub

Private Sub btnCapturarClmN_Click()
    climaNorte.tieneEditadosClmN
    btnCapturarClmN.Enabled = False
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not CapturaMatutino.editando Then
        On Error Resume Next
        If Not Intersect(Target, climaNorte.rgoClimaNorte) Is Nothing Then   'Rango editable
            'MsgBox "Área editable"
            climaNorte.modificadoClmN Target.Row, Target.Column
            btnCapturarClmN.Enabled = True
        ElseIf Not Intersect(Target, Range("B5:K5")) Is Nothing Then
            f = Range("B5").Value
            If IsDate(f) Then
                f = Format(f, "yyyy/mm/dd")
                CapturaMatutino.fecha = f
            End If
            CapturaMatutino.actualizaHojas
        End If
    End If
End Sub
