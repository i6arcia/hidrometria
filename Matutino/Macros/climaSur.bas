Attribute VB_Name = "climaSur"
Option Explicit

'***************************************
'***************************************
'***********   CLIMA SUR  **************
'***************************************
'***************************************

'Hojas de excel
Dim hjClmSur As Excel.Worksheet

'Control de manejo de hoja
Private filIni As Integer
Private filFin As Integer

Private colClv As Integer
Private colLluvia As Integer
Private colAmb As Integer
Private colMax As Integer
Private colMin As Integer
Private colEvap As Integer
Private colPrs As Integer
Private colHum As Integer

Private valLluvia As String
Private valAmb As String
Private valMax As String
Private valMin As String
Private valEvap As String
Private valPres As String
Private valHum As String

'Contadores y banderas
Private i As Integer

'Variable control de presas
Dim mClmS(100, 11) As String
Dim imClmS As Integer
'Rango
Public rgoClimaSur As Range

Sub iniciaClmSur()
    Dim clave As String
    'Asigna hojas a variables
    Set hjClmSur = Worksheets("Sur")
    'Control de hoja
    filIni = 8
    filFin = hjClmSur.Range("A" & Rows.Count).End(xlUp).Row
    colClv = 1
    colLluvia = 7
    colAmb = 8
    colMax = 9
    colMin = 10
    colEvap = 11
    colPrs = 5
    colHum = 6
    
    'Contador matriz control
    imClmS = 0
    
'   Matriz para el control de datos en la hoja
'*****************************************************
'| clave | Estado | Fila  | Lluvia | Ambiente | Máxima | Minima | Evaporación | Presión | Humedad | AmbienteAyer |
'*****************************************************
    For i = filIni To filFin
        clave = hjClmSur.Cells(i, colClv).Value
        If clave <> "" Then
            setMClimaS clave, 0, i, "0", "0", "0", "0", "0", "0", "0", "0", "0"
        End If
    Next i
    
    rangoClmSur
    
End Sub

Sub limpiaClmSur()
    If filIni = 0 Then iniciaClmSur
    
    CapturaMatutino.editando = True
    
    '****** TITULO DE HOJA CLIMA NORTE ********
    hjClmSur.Range("B5").Value = "Xalapa, Ver. -- --"
    
    hjClmSur.Range(hjClmSur.Cells(filIni, colPrs), hjClmSur.Cells(filFin, colEvap)).ClearContents
    
    CapturaMatutino.editando = False
End Sub

Sub obtieneClmSur()
    If filIni = 0 Then iniciaClmSur
    
    CapturaMatutino.editando = True
    
    ClmSurLluvia
    ClmSurTemp
    ClmSurEvap
    ClmSurPresion
    ClmSurHumedad
    
    'Ambiente ayer
    
    CapturaMatutino.editando = False
End Sub

Private Sub ClmSurLluvia()
    Dim valLluvia As String
    If filIni = 0 Then iniciaClmSur
    For i = 0 To imClmS - 1
        valLluvia = dataBase.getLluvia(mClmS(i, 0), fecha)
        If valLluvia <> "" Then
            hjClmSur.Cells(mClmS(i, 2), colLluvia).Value = valLluvia
            mClmS(i, 1) = 1
            mClmS(i, 3) = 1
        End If
    Next i
End Sub

Private Sub ClmSurTemp()
    Dim valAmb As String, valMax As String, valMin As String
    
    If filIni = 0 Then iniciaClmSur
    
    For i = 0 To imClmS - 1
        dataBase.getTemp mClmS(i, 0), fecha
        If dataBase.temperaturas(1) <> "" Then
            hjClmSur.Cells(mClmS(i, 2), colAmb).Value = dataBase.temperaturas(1)
            mClmS(i, 1) = 1
            mClmS(i, 4) = 1
        End If
        If dataBase.temperaturas(2) <> "" Then
            hjClmSur.Cells(mClmS(i, 2), colMax).Value = dataBase.temperaturas(2)
            mClmS(i, 1) = 1
            mClmS(i, 5) = 1
        End If
        If dataBase.temperaturas(3) <> "" Then
            hjClmSur.Cells(mClmS(i, 2), colMin).Value = dataBase.temperaturas(3)
            mClmS(i, 1) = 1
            mClmS(i, 6) = 1
        End If
    Next i
End Sub

Private Sub ClmSurEvap()
    If filIni = 0 Then iniciaClmSur

    Dim valEvap As String
    For i = 0 To imClmS - 1
        valEvap = dataBase.getEvaporacion(mClmS(i, 0), fecha)
        If valEvap <> "" Then
            hjClmSur.Cells(mClmS(i, 2), colEvap).Value = valEvap
            mClmS(i, 1) = 1
            mClmS(i, 7) = 1
        End If
    Next i
End Sub

Private Sub ClmSurPresion()
    If filIni = 0 Then iniciaClmSur
    
    Dim valPres As String
    For i = 0 To imClmS - 1
        If mClmS(i, 0) = "VERVC" Or mClmS(i, 0) = "ORZVC" Or mClmS(i, 0) = "COTVC" Then
            valPres = dataBase.getPresion(mClmS(i, 0), fecha)
            hjClmSur.Cells(mClmS(i, 2), colPrs).Value = valPres
            mClmS(i, 1) = 1
            mClmS(i, 8) = 1
        End If
    Next i
End Sub

Private Sub ClmSurHumedad()
    If filIni = 0 Then iniciaClmSur
    
    Dim valHum As String
    For i = 0 To imClmS - 1
        If mClmS(i, 0) = "VERVC" Or mClmS(i, 0) = "ORZVC" Or mClmS(i, 0) = "COTVC" Or mClmS(i, 0) = "SJNVC" Then
            valHum = dataBase.getHumedad(mClmS(i, 0), fecha)
            hjClmSur.Cells(mClmS(i, 2), colHum).Value = valHum
            mClmS(i, 1) = 1
            mClmS(i, 9) = 1
        End If
    Next i
End Sub

Private Sub rangoClmSur()
    If filIni = 0 Then iniciaClmSur
    
    Set rgoClimaSur = hjClmSur.Range(hjClmSur.Cells(filIni, colLluvia), hjClmSur.Cells(filFin, colEvap))
    For i = 0 To imClmS
        If mClmS(i, 0) = "VERVC" Or mClmS(i, 0) = "ORZVC" Or mClmS(i, 0) = "COTVC" Then
            Set rgoClimaSur = Union(rgoClimaSur, hjClmSur.Cells(mClmS(i, 2), colPrs), hjClmSur.Cells(mClmS(i, 2), colHum))
        ElseIf mClmS(i, 0) = "SJNVC" Then
            Set rgoClimaSur = Union(rgoClimaSur, hjClmSur.Cells(mClmS(i, 2), colHum))
        End If
    Next i
End Sub

Public Sub modificadoClmS(fil As Integer, col As Integer)
Dim colorCelda As String
Dim posicion As Integer

If imClmS = 0 Then CapturaMatutino.actualizaHojas

    For i = 0 To imClmS - 1
        If mClmS(i, 2) = fil Then
            Select Case col
                Case colLluvia
                    posicion = 3
                Case colAmb
                    posicion = 4
                Case colMax
                    posicion = 5
                Case colMin
                    posicion = 6
                Case colEvap
                    posicion = 7
                Case colPrs
                    posicion = 8
                Case colHum
                    posicion = 9
            End Select
            
            If mClmS(i, posicion) = 1 Then
                mClmS(i, posicion) = 3
                mClmS(i, 1) = 5
                'MsgBox "Modificar"
            ElseIf mClmS(i, posicion) = 0 Then
                mClmS(i, posicion) = 2
                mClmS(i, 1) = 5
                'MsgBox "Agregar"
            End If
            If mClmS(i, 4) = 2 And mClmS(i, 5) = 2 And mClmS(i, 6) = 2 Then mClmS(i, 11) = 2
            If mClmS(i, 4) = 3 And mClmS(i, 5) = 3 And mClmS(i, 6) = 3 Then mClmS(i, 11) = 3
            gris fil, col
            Exit For
        End If
    Next i
End Sub

Public Sub tieneEditadosClmS()
    Dim valor As String
    Dim max As String, min As String, amb As String
    Dim respuesta As Boolean
    Dim hayErrores As Boolean
    Dim j As Integer
    Dim pCol As Integer
    
    hayErrores = False
    
    For i = 0 To imClmS - 1
        If mClmS(i, 1) = 5 Then
            For j = 3 To 11
                    If mClmS(i, j) = 3 Or mClmS(i, j) = 2 Then
                        Select Case j
                            Case 3
                                valor = hjClmSur.Cells(mClmS(i, 2), colLluvia).Value
                                respuesta = validacion.validaLluvia(valor, mClmS(i, 0), fecha, mClmS(i, j))
                                pCol = colLluvia
                            Case 4 To 6       '**********TEMPERATURAS***************
                                If mClmS(i, 11) = 2 Or mClmS(i, 11) = 3 Then
                                    max = hjClmSur.Cells(mClmS(i, 2), colMax).Value
                                    min = hjClmSur.Cells(mClmS(i, 2), colMin).Value
                                    amb = hjClmSur.Cells(mClmS(i, 2), colAmb).Value
                                    respuesta = validacion.validaTemps(amb, max, min, mClmS(i, 0), fecha, 3)
                                    pCol = colAmb
                                    j = 6
                                Else
                                    respuesta = False
                                    pCol = colAmb
                                    j = 6
                                End If
                                
                                If respuesta Then
                                    blanco CInt(mClmS(i, 2)), colMax
                                    blanco CInt(mClmS(i, 2)), colMin
                                Else
                                    rojo CInt(mClmS(i, 2)), colMax
                                    rojo CInt(mClmS(i, 2)), colMin
                                End If
                                mClmS(i, 11) = 1
                                mClmS(i, 4) = 1
                                mClmS(i, 5) = 1
                                mClmS(i, 6) = 1
                            Case 7      '**********EVAPORACION***************
                                valor = hjClmSur.Cells(mClmS(i, 2), colEvap).Value
                                respuesta = validacion.validaEvap(valor, mClmS(i, 0), fecha, mClmS(i, j))
                                pCol = colEvap
                            Case 8       '**********PRESION***************
                                valor = hjClmSur.Cells(mClmS(i, 2), colPrs).Value
                                respuesta = validacion.validaPresion(valor, mClmS(i, 0), fecha, mClmS(i, j))
                                pCol = colPrs
                            Case 9       '**********HUMEDAD***************
                                valor = hjClmSur.Cells(mClmS(i, 2), colHum).Value
                                respuesta = validacion.validaHumedad(valor, mClmS(i, 0), fecha, mClmS(i, j))
                                pCol = colHum
                        End Select
                        'Respuesta de validacion de datos
                        If respuesta Then
                            blanco CInt(mClmS(i, 2)), pCol
                        Else
                            rojo CInt(mClmS(i, 2)), pCol
                            hayErrores = True
                        End If
                        mClmS(i, j) = 1
                        mClmS(i, 1) = 1
                    End If
            Next j
        End If
    Next i
    
    If hayErrores Then
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
    Else
        CapturaMatutino.actualizaHojas
    End If

End Sub


'*********************
' **    GET//SET    **
'*********************
Private Sub setMClimaS(clv As String, edo As Integer, fil As Integer, lluv As String, amb As String, max As String, min As String, evp As String, prs As String, hum As String, ambA As String, edoTemp As String)
    'Matriz para clima 1
    'mClmS{
    '   0 | Clave
    '   1 | Estado
    '   2 | Fila
    '   3 | Lluvia
    '   4 | Ambiente
    '   5 | Máxima
    '   6 | Mínima
    '   7 | Evaporación
    '   8 | Presión
    '   9 | Humedad
    '   10| Ambiente de Ayer
    '}
    mClmS(imClmS, 0) = clv
    mClmS(imClmS, 1) = edo
    mClmS(imClmS, 2) = fil
    mClmS(imClmS, 3) = lluv
    mClmS(imClmS, 4) = amb
    mClmS(imClmS, 5) = max
    mClmS(imClmS, 6) = min
    mClmS(imClmS, 7) = evp
    mClmS(imClmS, 8) = prs
    mClmS(imClmS, 9) = hum
    mClmS(imClmS, 10) = ambA
    mClmS(imClmS, 11) = edoTemp
    
    imClmS = imClmS + 1
End Sub

'***************************************
'//////////////////////////////////////
'           ENFASIS EN HOJA
'//////////////////////////////////////
'***************************************

Private Sub gris(fil As Integer, Optional col As Integer)
    hjClmSur.Cells(fil, col).Interior.Color = RGB(242, 242, 242)
    hjClmSur.Cells(fil, col).Font.Color = vbBlack
    hjClmSur.Cells(fil, col).Font.Bold = False
End Sub

Private Sub blanco(fil As Integer, Optional col As Integer)
    hjClmSur.Cells(fil, col).Interior.Color = vbWhite
    hjClmSur.Cells(fil, col).Font.Color = vbBlack
    hjClmSur.Cells(fil, col).Font.Bold = False
End Sub

Private Sub rojo(fil As Integer, Optional col As Integer)
    hjClmSur.Cells(fil, col).Interior.Color = vbRed
    hjClmSur.Cells(fil, col).Font.Color = vbBlack
    hjClmSur.Cells(fil, col).Font.Bold = False
End Sub

Private Sub amarillo(fil As Integer, Optional col As Integer)
    hjClmSur.Cells(fil, col).Interior.Color = vbYellow
    hjClmSur.Cells(fil, col).Font.Color = vbBlack
    hjClmSur.Cells(fil, col).Font.Bold = False
End Sub
Private Sub naranja(fil As Integer, Optional col As Integer)
    hjClmSur.Cells(fil, col).Interior.Color = RGB(255, 192, 0)
    hjClmSur.Cells(fil, col).Font.Color = vbBlack
    hjClmSur.Cells(fil, col).Font.Bold = False
End Sub
Private Sub textoRojo(fil As Integer, Optional col As Integer)
    hjClmSur.Cells(fil, col).Interior.Color = RGB(255, 192, 0)
    hjClmSur.Cells(fil, col).Font.Color = RGB(192, 0, 0)
    hjClmSur.Cells(fil, col).Font.Bold = True
End Sub

