Attribute VB_Name = "climaNorte"
Option Explicit

'***************************************
'***************************************
'***********  CLIMA NORTE **************
'***************************************
'***************************************

'Hojas de excel
Dim hjClmNrt As Excel.Worksheet

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
Dim mClmN(100, 11) As String
Dim imClmN As Integer
'Rango
Public rgoClimaNorte As Range

Sub iniciaClmNrt()
    Dim clave As String
    'Asigna hojas a variables
    Set hjClmNrt = Worksheets("Norte")
    'Control de hoja
    filIni = 8
    filFin = hjClmNrt.Range("A" & Rows.Count).End(xlUp).Row
    colClv = 1
    colLluvia = 7
    colAmb = 8
    colMax = 9
    colMin = 10
    colEvap = 11
    colPrs = 5
    colHum = 6
    
    'Contador matriz control
    imClmN = 0
    
'   Matriz para el control de datos en la hoja
'*****************************************************
'| clave | Estado | Fila  | Lluvia | Ambiente | Máxima | Minima | Evaporación | Presión | Humedad | AmbienteAyer | edoTemp
'*****************************************************
    For i = filIni To filFin
        clave = hjClmNrt.Cells(i, colClv).Value
        If clave <> "" Then
            setMClimaN clave, 0, i, "0", "0", "0", "0", "0", "0", "0", "0", "0"
        End If
    Next i
    
    rangoClmNrt
    
End Sub

Sub limpiaClmNrt()
    If filIni = 0 Then iniciaClmNrt
    
    CapturaMatutino.editando = True
    
    '****** TITULO DE HOJA CLIMA NORTE ********
    hjClmNrt.Range("B5").Value = "Xalapa, Ver. -- --"
    
    hjClmNrt.Range(hjClmNrt.Cells(filIni, colPrs), hjClmNrt.Cells(filFin, colEvap)).ClearContents
    
    CapturaMatutino.editando = False
End Sub

Sub obtieneClmNrt()
    If filIni = 0 Then iniciaClmNrt
    
    CapturaMatutino.editando = True
    
    ClmNrtLluvia
    ClmNrtTemp
    ClmNrtEvap
    ClmNrtPresion
    ClmNrtHumedad
    
    'Ambiente ayer
    
    CapturaMatutino.editando = False
End Sub

'Obtiene lluvia desde la base de datos
Private Sub ClmNrtLluvia()
    For i = 0 To imClmN - 1
        valLluvia = dataBase.getLluvia(mClmN(i, 0), fecha)
        If valLluvia <> "" Then
            hjClmNrt.Cells(mClmN(i, 2), colLluvia).Value = valLluvia
            mClmN(i, 1) = 1
            mClmN(i, 3) = 1
        End If
    Next i
End Sub
'Obtiene temperaturas desde la base de datos
Private Sub ClmNrtTemp()
    For i = 0 To imClmN - 1
        dataBase.getTemp mClmN(i, 0), fecha
        If dataBase.temperaturas(1) <> "" Then
            hjClmNrt.Cells(mClmN(i, 2), colAmb).Value = dataBase.temperaturas(1)
            mClmN(i, 1) = 1
            mClmN(i, 4) = 1
        End If
        If dataBase.temperaturas(2) <> "" Then
            hjClmNrt.Cells(mClmN(i, 2), colMax).Value = dataBase.temperaturas(2)
            mClmN(i, 1) = 1
            mClmN(i, 5) = 1
        End If
        If dataBase.temperaturas(3) <> "" Then
            hjClmNrt.Cells(mClmN(i, 2), colMin).Value = dataBase.temperaturas(3)
            mClmN(i, 1) = 1
            mClmN(i, 6) = 1
        End If
        If mClmN(i, 4) = 1 And mClmN(i, 5) = 1 And mClmN(i, 6) = 1 Then mClmN(i, 11) = 1
    Next i
End Sub
'Obtiene Evaporacion desde la base de datos
Private Sub ClmNrtEvap()
    For i = 0 To imClmN - 1
        valEvap = dataBase.getEvaporacion(mClmN(i, 0), fecha)
        If valEvap <> "" Then
            hjClmNrt.Cells(mClmN(i, 2), colEvap).Value = valEvap
            mClmN(i, 1) = 1
            mClmN(i, 7) = 1
        End If
    Next i
End Sub
'Obtiene Presion de la estacion de la base de datos
Private Sub ClmNrtPresion()
    For i = 0 To imClmN - 1
        If mClmN(i, 0) = "TXPVC" Or mClmN(i, 0) = "XOBVC" Then
            valPres = dataBase.getPresion(mClmN(i, 0), fecha)
            hjClmNrt.Cells(mClmN(i, 2), colPrs).Value = valPres
            mClmN(i, 1) = 1
            mClmN(i, 8) = 1
        End If
    Next i
End Sub
'Obtiene humedad desde la base de datos
Private Sub ClmNrtHumedad()
    For i = 0 To imClmN - 1
        If mClmN(i, 0) = "TXPVC" Or mClmN(i, 0) = "XOBVC" Then
            valHum = dataBase.getHumedad(mClmN(i, 0), fecha)
            hjClmNrt.Cells(mClmN(i, 2), colHum).Value = valHum
            mClmN(i, 1) = 1
            mClmN(i, 9) = 1
        End If
    Next i
End Sub

Private Sub rangoClmNrt()
    If filIni = 0 Then iniciaClmNrt
    
    Set rgoClimaNorte = hjClmNrt.Range(hjClmNrt.Cells(filIni, colLluvia), hjClmNrt.Cells(filFin, colEvap))
    For i = 0 To imClmN - 1
        If mClmN(i, 0) = "TXPVC" Or mClmN(i, 0) = "XOBVC" Then
            Set rgoClimaNorte = Union(rgoClimaNorte, hjClmNrt.Cells(mClmN(i, 2), colPrs), hjClmNrt.Cells(mClmN(i, 2), colHum))
        End If
    Next i
End Sub

Public Sub modificadoClmN(fil As Integer, col As Integer)
Dim colorCelda As String
Dim posicion As Integer

If imClmN = 0 Then CapturaMatutino.actualizaHojas

    For i = 0 To imClmN - 1
        If mClmN(i, 2) = fil Then
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
            
            If mClmN(i, posicion) = 1 Then
                mClmN(i, posicion) = 3
                mClmN(i, 1) = 5
                'MsgBox "Modificar"
            ElseIf mClmN(i, posicion) = 0 Then
                mClmN(i, posicion) = 2
                mClmN(i, 1) = 5
                'MsgBox "Agregar"
            End If
            If mClmN(i, 4) = 2 And mClmN(i, 5) = 2 And mClmN(i, 6) = 2 Then mClmN(i, 11) = 2
            If mClmN(i, 4) = 3 And mClmN(i, 5) = 3 And mClmN(i, 6) = 3 Then mClmN(i, 11) = 3
            gris fil, col
            Exit For
        End If
    Next i
End Sub

Public Sub tieneEditadosClmN()
    Dim valor As String
    Dim max As String, min As String, amb As String
    Dim respuesta As Boolean
    Dim hayErrores As Boolean
    Dim j As Integer
    Dim pCol As Integer
    
    hayErrores = False
    
    For i = 0 To imClmN - 1
        If mClmN(i, 1) = 5 Then
            For j = 3 To 10
                    If mClmN(i, j) = 3 Or mClmN(i, j) = 2 Then
                        Select Case j
                            Case 3      '**********LLUVIA***************
                                valor = hjClmNrt.Cells(mClmN(i, 2), colLluvia).Value
                                respuesta = validacion.validaLluvia(valor, mClmN(i, 0), fecha, mClmN(i, j))
                                pCol = colLluvia
                            Case 4 To 6        '**********TEMPERATURAS***************
                                If mClmN(i, 11) = 2 Or mClmN(i, 11) = 3 Then
                                    max = hjClmNrt.Cells(mClmN(i, 2), colMax).Value
                                    min = hjClmNrt.Cells(mClmN(i, 2), colMin).Value
                                    amb = hjClmNrt.Cells(mClmN(i, 2), colAmb).Value
                                    respuesta = validacion.validaTemps(amb, max, min, mClmN(i, 0), fecha, mClmN(i, 11))
                                    pCol = colAmb
                                    j = 6
                                Else
                                    respuesta = False
                                    pCol = colAmb
                                    j = 6
                                End If
                                
                                If respuesta Then
                                    blanco CInt(mClmN(i, 2)), colMax
                                    blanco CInt(mClmN(i, 2)), colMin
                                Else
                                    rojo CInt(mClmN(i, 2)), colMax
                                    rojo CInt(mClmN(i, 2)), colMin
                                End If
                                mClmN(i, 11) = 1
                                mClmN(i, 4) = 1
                                mClmN(i, 5) = 1
                                mClmN(i, 6) = 1
                            Case 7      '**********EVAPORACION***************
                                valor = hjClmNrt.Cells(mClmN(i, 2), colEvap).Value
                                respuesta = validacion.validaEvap(valor, mClmN(i, 0), fecha, mClmN(i, j))
                                pCol = colEvap
                            Case 8       '**********PRESION***************
                                valor = hjClmNrt.Cells(mClmN(i, 2), colPrs).Value
                                respuesta = validacion.validaPresion(valor, mClmN(i, 0), fecha, mClmN(i, j))
                                pCol = colPrs
                            Case 9       '**********HUMEDAD***************
                                valor = hjClmNrt.Cells(mClmN(i, 2), colHum).Value
                                respuesta = validacion.validaHumedad(valor, mClmN(i, 0), fecha, mClmN(i, j))
                                pCol = colHum
                        End Select
                        'Respuesta de validacion de datos
                        If respuesta Then
                            blanco CInt(mClmN(i, 2)), pCol
                        Else
                            rojo CInt(mClmN(i, 2)), pCol
                            hayErrores = True
                        End If
                        mClmN(i, j) = 1
                        mClmN(i, 1) = 1
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
Private Sub setMClimaN(clv As String, edo As Integer, fil As Integer, lluv As String, amb As String, max As String, min As String, evp As String, prs As String, hum As String, ambA As String, edoTemp As String)
    'Matriz para clima 1
    'mClmN{
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
    '   11| edoTemp
    '}
    mClmN(imClmN, 0) = clv
    mClmN(imClmN, 1) = edo
    mClmN(imClmN, 2) = fil
    mClmN(imClmN, 3) = lluv
    mClmN(imClmN, 4) = amb
    mClmN(imClmN, 5) = max
    mClmN(imClmN, 6) = min
    mClmN(imClmN, 7) = evp
    mClmN(imClmN, 8) = prs
    mClmN(imClmN, 9) = hum
    mClmN(imClmN, 10) = ambA
    mClmN(imClmN, 11) = edoTemp
    
    imClmN = imClmN + 1
End Sub
'***************************************
'//////////////////////////////////////
'           ENFASIS EN HOJA
'//////////////////////////////////////
'***************************************

Private Sub gris(fil As Integer, Optional col As Integer)
    hjClmNrt.Cells(fil, col).Interior.Color = RGB(242, 242, 242)
    hjClmNrt.Cells(fil, col).Font.Color = vbBlack
    hjClmNrt.Cells(fil, col).Font.Bold = False
End Sub

Private Sub blanco(fil As Integer, Optional col As Integer)
    hjClmNrt.Cells(fil, col).Interior.Color = vbWhite
    hjClmNrt.Cells(fil, col).Font.Color = vbBlack
    hjClmNrt.Cells(fil, col).Font.Bold = False
End Sub

Private Sub rojo(fil As Integer, Optional col As Integer)
    hjClmNrt.Cells(fil, col).Interior.Color = vbRed
    hjClmNrt.Cells(fil, col).Font.Color = vbBlack
    hjClmNrt.Cells(fil, col).Font.Bold = False
End Sub

Private Sub amarillo(fil As Integer, Optional col As Integer)
    hjClmNrt.Cells(fil, col).Interior.Color = vbYellow
    hjClmNrt.Cells(fil, col).Font.Color = vbBlack
    hjClmNrt.Cells(fil, col).Font.Bold = False
End Sub
Private Sub naranja(fil As Integer, Optional col As Integer)
    hjClmNrt.Cells(fil, col).Interior.Color = RGB(255, 192, 0)
    hjClmNrt.Cells(fil, col).Font.Color = vbBlack
    hjClmNrt.Cells(fil, col).Font.Bold = False
End Sub
Private Sub textoRojo(fil As Integer, Optional col As Integer)
    hjClmNrt.Cells(fil, col).Interior.Color = RGB(255, 192, 0)
    hjClmNrt.Cells(fil, col).Font.Color = RGB(192, 0, 0)
    hjClmNrt.Cells(fil, col).Font.Bold = True
End Sub
