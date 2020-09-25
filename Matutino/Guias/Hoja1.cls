VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub btnActualizar_Click()
    'Bandera control de cambios
    Seguimiento.bandera = True
    'Inicia variables
    Seguimiento.iniciaSeg
    'Obtiene estaciones climatologicas
    Seguimiento.obtenerEst
    'Obtiene datos de nivel de las estaciones
    Seguimiento.obtenerNiveles
    'Obtener valores de lluvia
    Seguimiento.obtenerLluvia
    'Bandera control de cambios
    Seguimiento.bandera = False
End Sub

Private Sub btnGuardar_Click()
    Seguimiento.guardarNiveles
    Seguimiento.guardarLluvia
    btnGuardar.Enabled = False
    Hoja2.btnGuardar2.Visible = False
End Sub

Private Sub lblNext_Click()
    Dim f As String
    Dim fD As String
    'Control de cambios
    Seguimiento.bandera = True
    
    f = Seguimiento.getFecha

    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If

    fD = Format(DateAdd("d", 1, f), "yyyy/mm/dd")

    Range("B5").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    Hoja2.Range("B5").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("B5").Interior.color = RGB(255, 242, 204)
        Hoja2.Range("B5").Interior.color = RGB(189, 215, 238)
    Else
        Range("B5").Interior.color = xlNone
        Hoja2.Range("B5").Interior.color = xlNone
    End If

    Seguimiento.setFecha fD
    Seguimiento.obtenerNiveles
    'Obtener valores de lluvia
    Seguimiento.obtenerLluvia
    'Control de cambios
    Seguimiento.bandera = False
End Sub

Private Sub lblNow_Click()
    'Bandera control de cambios
    Seguimiento.bandera = True
    'Inicia variables
    Seguimiento.iniciaSeg
    'Obtiene estaciones climatologicas
    Seguimiento.obtenerEst
    'Obtiene datos de nivel de las estaciones
    Seguimiento.obtenerNiveles
    'Obtener valores de lluvia
    Seguimiento.obtenerLluvia
    'Bandera control de cambios
    Seguimiento.bandera = False
End Sub

Private Sub lblPrev_Click()
    Dim f As String
    Dim fD As String
    'Control de cambios
    Seguimiento.bandera = True
    
    f = Seguimiento.getFecha

    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If

    fD = Format(DateDiff("d", 1, f), "yyyy/mm/dd")

    Range("B5").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    Hoja2.Range("B5").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("B5").Interior.color = RGB(255, 242, 204)
        Hoja2.Range("B5").Interior.color = RGB(189, 215, 238)
    Else
        Range("B5").Interior.color = xlNone
        Hoja2.Range("B5").Interior.color = xlNone
    End If

    Seguimiento.setFecha fD
    Seguimiento.obtenerNiveles
    'Obtener valores de lluvia
    Seguimiento.obtenerLluvia
    'Control de cambios
    Seguimiento.bandera = False
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim fi As Integer, ff As Integer
    Dim ci As Integer, cf As Integer
    Dim f As String
    'Filtro si esta editando documento
    If Not Seguimiento.bandera Then
        On Error GoTo inicia
        'Verifica rango de edicion
        If Target.Column = 1 Then
            MsgBox "Debes modificar el grupo para agregar nuevas estaciones", vbCritical
            btnActualizar_Click
        ElseIf Target.Row = 1 Then
            'MsgBox "Colums"
            btnActualizar_Click
        ElseIf Not Intersect(Target, Seguimiento.getRanEdit) Is Nothing Then
            'If ctrlCambios.restFil <> 0 Then
                ctrlCambios.modificado Target.Row, Target.Column
                btnGuardar.Enabled = True
                Hoja2.btnGuardar2.Visible = True
            'End If
        ElseIf Target.Row = 7 Then
            btnActualizar_Click
        ElseIf Not Intersect(Target, Range("B5:Y5")) Is Nothing Then
            f = Range("B5").Value
            If IsDate(f) Then
                f = Format(f, "yyyy/mm/dd")
                'Control de cambios
                Seguimiento.bandera = True
                Range("B5").Value = "Xalapa, Ver. -- " & Format(f, "dddd") & " " & Format(f, "dd") & " de " & Format(f, "mmmm") & " de " & Format(f, "yyyy") & " --"
                Hoja2.Range("B5").Value = "Xalapa, Ver. -- " & Format(f, "dddd") & " " & Format(f, "dd") & " de " & Format(f, "mmmm") & " de " & Format(f, "yyyy") & " --"
                If (f = Format(Now, "yyyy/mm/dd")) Then
                    Range("B5").Interior.color = RGB(255, 242, 204)
                    Hoja2.Range("B5").Interior.color = RGB(189, 215, 238)
                Else
                    Range("B5").Interior.color = xlNone
                    Hoja2.Range("B5").Interior.color = xlNone
                End If

                Seguimiento.setFecha f
                Seguimiento.obtenerNiveles
                Seguimiento.obtenerLluvia
                'Bandera control de cambios
                Seguimiento.bandera = False
            Else
                lblNow_Click
            End If
        End If
    End If
    Exit Sub
inicia:
    lblNow_Click
End Sub
