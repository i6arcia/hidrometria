Attribute VB_Name = "ctrlCambios"
Option Explicit

Public mCambios() As Integer
Public lCambios() As Integer
'mCambios (Columna, Fila)
'mCambios (Hora, Estacion)

Sub modificado(fil As Integer, col As Integer)
    fil = fil - Seguimiento.getFilIni
    col = col - Seguimiento.getColIni
    On Error GoTo inicia
    If (mCambios(col, fil) = 0) Then
        'Celda vacia
        mCambios(col, fil) = 2
        'MsgBox "La celda esta vacia, se agregara registro"
        gris col, fil
    ElseIf (mCambios(col, fil) = 1) Then
        mCambios(col, fil) = 3
        'MsgBox "La celda esta ocupada, Se modificara el registro"
        gris col, fil
    End If
    Exit Sub
inicia:
    'Bandera control de cambios
    Seguimiento.bandera = True
    'Inicia variables
    Seguimiento.iniciaSeg
    'Obtiene estaciones climatologicas
    Seguimiento.obtenerEst
    'Obtiene datos de nivel de las estaciones
    Seguimiento.obtenerNiveles
    'Obtiene datos de lluvia de las estaciones
    Seguimiento.obtenerLluvia
    'Bandera control de cambios
    Seguimiento.bandera = False
End Sub

Sub modificadoLluvia(fil As Integer, col As Integer)
    fil = fil - Seguimiento.getFilIni
    col = col - Seguimiento.getColIni
    On Error GoTo inicia
    If (lCambios(col, fil) = 0) Then
        'Celda vacia
        lCambios(col, fil) = 2
        'MsgBox "La celda esta vacia, se agregara registro"
        gris2 col, fil
    ElseIf (lCambios(col, fil) = 1) Then
        lCambios(col, fil) = 3
        'MsgBox "La celda esta ocupada, Se modificara el registro"
        gris2 col, fil
    End If
    Exit Sub
inicia:
    'Bandera control de cambios
    Seguimiento.bandera = True
    'Inicia variables
    Seguimiento.iniciaSeg
    'Obtiene estaciones climatologicas
    Seguimiento.obtenerEst
    'Obtiene datos de nivel de las estaciones
    Seguimiento.obtenerNiveles
    'Obtiene datos de lluvia de las estaciones
    Seguimiento.obtenerLluvia
    'Bandera control de cambios
    Seguimiento.bandera = False
End Sub


Private Sub gris(col As Integer, rows As Integer)
    Dim hj As Excel.Worksheet
    Set hj = Worksheets("Niveles")
    hj.Cells(rows + Seguimiento.getFilIni, col + Seguimiento.getColIni).Interior.color = RGB(242, 242, 242)
End Sub

Private Sub gris2(col As Integer, rows As Integer)
    Dim hj As Excel.Worksheet
    Set hj = Worksheets("Lluvia")
    hj.Cells(rows + Seguimiento.getFilIni, col + Seguimiento.getColIni).Interior.color = RGB(242, 242, 242)
End Sub

