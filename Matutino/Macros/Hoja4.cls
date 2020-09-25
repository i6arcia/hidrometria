VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnActualizarClmS_Click()
    CapturaMatutino.actualizaHojas
    btnCapturarClmS.Enabled = False
End Sub

Private Sub btnCapturarClmS_Click()
    climaSur.tieneEditadosClmS
    btnCapturarClmS.Enabled = False
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not CapturaMatutino.editando Then
        On Error Resume Next
        If Not Intersect(Target, climaSur.rgoClimaSur) Is Nothing Then   'Rango editable
            'MsgBox "Área editable"
            climaSur.modificadoClmS Target.Row, Target.Column
            btnCapturarClmS.Enabled = True
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
