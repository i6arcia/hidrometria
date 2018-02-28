VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub lblNext_Click()
    'Variables
    Dim f As String 'Fecha
    Dim fD As String 'Fecha diferida
    'Obtiene la fecha global
    f = Observatorios.getFecha
    'Si no existe asigna el dia de hoy
    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If
    
    'Agrega 1 dia
    fD = Format(DateAdd("d", 1, f), "yyyy/mm/dd")
    'Escribe la fecha en la hoja
    Range("E7").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    'Si la fecha es igual al día actual lo rellena color azul
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("E7").Interior.Color = RGB(221, 235, 247)
    Else
        Range("E7").Interior.Color = xlNone
    End If
    'Asigna la nueva fecha a la variable global
    Observatorios.setFecha fD
    'Actualiza la informacion
    Observatorios.actualizar
End Sub

Private Sub lblNow_Click()
    'Restablece la informacion a la fecha actual
    Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    Range("E7").Interior.Color = RGB(221, 235, 247)
    Observatorios.setFecha Format(Now, "yyyy/mm/dd")
    Observatorios.actualizar
End Sub

Private Sub lblPrev_Click()
    'Variables
    Dim f As String     'Fecha
    Dim fD As String    'Fecha diferida
    'Obtiene la fecha global
    f = Observatorios.getFecha
    'Si no existe asigna el dia de hoy
    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If
    'Resta un dia a la fecha
    fD = Format(DateAdd("d", -1, f), "yyyy/mm/dd")
    'Escribe la fecha en la hoja
    Range("E7").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    'Si la fecha es igual al dia de hoy, cambia a color azul de fondo
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("E7").Interior.Color = RGB(221, 235, 247)
    Else
        Range("E7").Interior.Color = xlNone
    End If
    'Actualiza la información
    Observatorios.setFecha fD
    Observatorios.actualizar
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim x As Integer, y As Integer
    Dim fil As Integer, col As Integer
    
     '**Control de cambios**
    
    'Verificar variables globales iniciadas
    If Observatorios.iniFil > 0 Then
        'Pregunta si está editando archivo "inicio"
        If Not Observatorios.flagEdit Then
            'Verificar rango de edición
            If Not Intersect(Target, Range(Cells(Observatorios.iniFil, 2), Cells(Observatorios.ultFil + 1, Observatorios.x * 3))) Is Nothing Then
                fil = Target.Row
                col = Target.Column
                If col = 1 Then
                    Observatorios.ultFil = 0
                    'Observatorios.inicio
                    'contar
                    
                    Observatorios.contarFilas
                ElseIf col > 1 And col < Observatorios.x * 3 Then
                    If (col Mod 3 = 0) Then
                        x = (col / 3) - 1
                        y = fil - Observatorios.iniFil
                        Observatorios.edit(x, y) = -1
                        MsgBox "Es valor de lluvia"
                    ElseIf ((col + 1) Mod 3 = 0) Then
                        'Observatorios.getInfoSIH ((col + 1) / 3)
                        MsgBox "Es valor de hora"
                    Else
                        MsgBox "Dentro del rango pero no es valor"
                    End If
                End If
            End If
        End If
    Else
        'Inicia variables
        Observatorios.inicio
    End If
End Sub


'If Not Observatorios.flagEdit Then
    '    If Observatorios.ultFil > 0 Then
    '        If Not Intersect(Target, Range("B11:R" & Observatorios.ultFil)) Is Nothing Then
    '            fil = Target.Row
    '            col = Target.Column
    '            If (col Mod 3 = 0) Then
    '                x = (col / 3) - 1
    '                y = fil - Observatorios.iniFil
    '                Observatorios.edit(x, y) = -1
    '            End If
    '        End If
    '    Else
    '        Observatorios.inicio
    '    End If
    'End If
