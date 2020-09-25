VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnActualizarPl_Click()
    CapturaMatutino.actualizaHojas
    btnCapturarPl.Enabled = False
End Sub

Private Sub btnCapturarPl_Click()
    pluviometros.tieneEditadosPluvios
    btnCapturarPl.Enabled = False
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Not CapturaMatutino.editando Then
        If Not Intersect(Target, pluviometros.rgoPluvios) Is Nothing Then   'Rango editable
            'MsgBox "Área editable"
            pluviometros.modificadoPluvios Target.Row
            btnCapturarPl.Enabled = True
        ElseIf Not Intersect(Target, Range("B5:H5")) Is Nothing Then
            f = Range("B5").Value
            If IsDate(f) Then
                f = Format(f, "yyyy/mm/dd")
                CapturaMatutino.fecha = f
            End If
            CapturaMatutino.actualizaHojas
        ElseIf Target.Column = 3 And Target.Row >= 8 Then
            pluviometros.iniciaPluvios
            pluviometros.limpiaPluvios
            pluviometros.obtienePluvios
        End If
    End If
End Sub
