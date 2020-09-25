Attribute VB_Name = "hidrometrica"
Option Explicit

'***************************************
'***************************************
'*********** HIDROMETRICA **************
'***************************************
'***************************************

'Hojas de excel
Dim hjHdr As Excel.Worksheet

'Control de manejo de hoja
Private filIni As Integer
Private filFin As Integer
Private colClv As Integer
Private colNivAyer As Integer
Private colNiv As Integer
Private colGasto As Integer
Private colTendencia As Integer
Private colNamo As Integer
Private colDStd As Integer

'Contadores y banderas
Private i As Integer

'Variable control de presas
Dim mHid(100, 6) As String
Dim imHid As Integer
'Rango
Public rgoHidrometicas As Range

'Inicia variables
Sub iniciaHidro()
    Dim clave As String
    'Asigna hojas a variables
    Set hjHdr = Worksheets("HIDROMETRICA")
    'Control de hoja
    filIni = 17
    colClv = 1
    colNiv = 7
    colNivAyer = 6
    colGasto = 8
    colTendencia = 9
    colNamo = 4
    colDStd = 10
    filFin = hjHdr.Range("A" & Rows.Count).End(xlUp).Row
    'Contador matriz control
    imHid = 0
    
'   Matriz para el control de datos en la hoja
'******************************************************************
'| clave     | Nivel | NivAy | Namo  | Estado | Fila  | Des. Std
'******************************************************************
    For i = filIni To filFin
        clave = hjHdr.Cells(i, colClv).Value
        If clave <> "" Then
            setMHid clave, "", "", "", 0, i, ""
        End If
    Next i
    rangoHid
End Sub

Sub limpiaHidro()
    If filIni = 0 Then iniciaHidro
    
    CapturaMatutino.editando = True
    
    '****** TITULO DE HOJA HIDROMETRICA ********
    hjHdr.Range("B5").Value = "Xalapa, Ver. -- --"
    
    hjHdr.Range(hjHdr.Cells(filIni, colNivAyer), hjHdr.Cells(filFin, colTendencia)).ClearContents
    'Limpia enfasis
    hjHdr.Range(hjHdr.Cells(filIni, colNiv), hjHdr.Cells(filFin, colNiv)).Font.Bold = False
    hjHdr.Range(hjHdr.Cells(filIni, colNiv), hjHdr.Cells(filFin, colNiv)).Font.Color = vbBlack
    For i = filIni To filFin
        If hjHdr.Cells(i, colNiv).Interior.Color <> RGB(166, 166, 166) Then
            hjHdr.Cells(i, colNiv).Interior.Color = vbWhite
        End If
    Next i
    CapturaMatutino.editando = False
End Sub

'Crea el rango de edicion de la hoja
Private Sub rangoHid()
    If filIni = 0 Then
        iniciaHidro
    Else
        Set rgoHidrometicas = hjHdr.Range(hjHdr.Cells(mHid(0, 5), colNiv), hjHdr.Cells(mHid(imHid - 1, 5), colNiv))
    End If
End Sub

'/////// ///////////
'/////// OBTIENE VALORES DE HIDRO  ///////////
'/////// ///////////
Sub obtieneHidro()
    If filIni = 0 Then iniciaHidro
    
    CapturaMatutino.editando = True
    
    namo
    nivelesAyer
    nivelesHoy
    desviacionStd
    
    tendencias
    enfasis
    
    CapturaMatutino.editando = False
End Sub
'*********************
' **    GET//SET    **
'*********************
Sub setMHid(clv As String, niv As String, nivAyer As String, namo As String, edo As Integer, fila As Integer, deS As String)
    'mHid   {
    '   0 | Clave de la estación
    '   1 | Nivel
    '   2 | Nivel del dia anterior
    '   3 | Namo de la estación
    '   4 | Estado del valor Vacio 0|1 Con dato
    '   5 | Numero de fila en la hoja
    '   6 | Desviacion estandar del nivel
    '}
    mHid(imHid, 0) = clv
    mHid(imHid, 1) = niv
    mHid(imHid, 2) = nivAyer
    mHid(imHid, 3) = namo
    mHid(imHid, 4) = edo
    mHid(imHid, 5) = fila
    mHid(imHid, 6) = deS
    imHid = imHid + 1
End Sub

'/////// Obtiene valores de NAMO  ///////////
Private Sub namo()
    Dim valNamo As String
    For i = 0 To imHid - 1
        valNamo = dataBase.getNamo(mHid(i, 0))
        hjHdr.Cells(mHid(i, 5), colNamo).Value = Format(valNamo, "0.00")
        mHid(i, 3) = valNamo
    Next i
End Sub
'/////// Obtiene valores de NIVEL DIA ANTERIOR ///////////
Private Sub nivelesAyer()
    dataBase.getNivelGrp "GRGCHIDRO", ayer
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            buscaEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), colNivAyer
        Next i
    End If
End Sub
'/////// Obtiene valores de NIVEL ACTUAL  ///////////
Private Sub nivelesHoy()
    dataBase.getNivelGrp "GRGCHIDRO", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            buscaEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), colNiv
        Next i
    End If
End Sub
'/////// Obtiene valores de DESVIACIÓN ESTÁNDAR ///////////
Private Sub desviacionStd()
Dim st As String
    For i = 0 To imHid - 1
        st = dataBase.getDesviacionStd(mHid(i, 0), fecha)
        mHid(i, 6) = st
        hjHdr.Cells(mHid(i, 5), colDStd).Value = st
    Next i
End Sub

'/////// ///////////
'/////// MODIFICACIONES EN HOJA  ///////////
'/////// ///////////

'/////// CALCULA TENDENCIA DEL RIO  ///////////
Private Sub tendencias()
Dim nivActual As String
Dim nivUltimo As String
    'Hay 2 niveles
    For i = 0 To imHid - 1
        If mHid(i, 1) <> "" Then
            If dataBase.hayDosNiveles(mHid(i, 0), fecha) Then
                'Obtiene 2 últimos niveles
                dataBase.getDosUNivls mHid(i, 0), fecha
                nivActual = dataBase.dosNiveles(0)
                nivUltimo = dataBase.dosNiveles(1)
                If nivActual > nivUltimo Then
                    '1 -> En ascenso
                    hjHdr.Cells(mHid(i, 5), colTendencia).Value = 1
                ElseIf nivActual = nivUltimo Then
                    '0 -> Se mantiene
                    hjHdr.Cells(mHid(i, 5), colTendencia).Value = 0
                Else
                    '-1 -> En descenso
                    hjHdr.Cells(mHid(i, 5), colTendencia).Value = -1
                End If
            End If
        End If
    Next i
End Sub

'/////// MARCA ENFASIS EN LA HOJA ///////////
Private Sub enfasis()
    For i = 0 To imHid - 1
        'Si no vacios Desviacion estandar || namo || nivel
        If mHid(i, 6) <> "" And mHid(i, 3) <> "" And mHid(i, 1) <> "" Then
            If CDbl(mHid(i, 1)) >= CDbl(mHid(i, 3)) Then
                textoRojo CInt(mHid(i, 5))
            ElseIf CDbl(mHid(i, 1)) >= (CDbl(mHid(i, 3)) - CDbl(mHid(i, 6))) Then
                naranja CInt(mHid(i, 5))
            Else
                blanco CInt(mHid(i, 5))
            End If
        End If
    Next i
End Sub

'/////// Busca estacion y escribe en hoja ///////////
Private Sub buscaEscribe(clv As String, val As String, col As Integer)
    Dim j As Integer
    For j = 0 To imHid
        If mHid(j, 0) = clv Then
            hjHdr.Cells(mHid(j, 5), col).Value = val
            If col = colNivAyer Then mHid(j, 2) = val
            If col = colNiv Then
                mHid(j, 1) = val
                mHid(j, 4) = 1
            End If
            
            Exit For
        End If
    Next j
End Sub

'/////// **************************** ///////////
'/////// Cambia estado de MODIFICADOS ///////////
'/////// **************************** ///////////
Public Sub modificadoHid(fil As Integer)
    'Recorre arreglo de datos
    For i = 0 To imHid - 1
        If mHid(i, 5) = fil Then
            If mHid(i, 4) = 1 Then
                mHid(i, 4) = 3
                'MsgBox "Modificar"
            ElseIf mHid(i, 4) = 0 Then
                mHid(i, 4) = 2
                'MsgBox "Agregar"
            End If
            gris fil
            Exit For
        End If
    Next i
End Sub

'/////// ******************************* ///////////
'/////// Analisa si hay valores editados ///////////
'/////// ******************************* ///////////
Public Sub tieneEditadosHid()
    Dim valNiv As String
    Dim respuesta As Boolean
    Dim hayErrores As Boolean
    
    'Valida que esten iniciadas las variables
    If filIni = 0 Then iniciaHidro
    
    hayErrores = False
    
    'Recorre el arreglo de estaciones
    For i = 0 To imHid - 1
        'Si hay valores para 2|Agregar o 3|Remplazar
        If mHid(i, 4) = 3 Or mHid(i, 4) = 2 Then
            'Obtiene el valor de nivel capturado en la celda
            valNiv = hjHdr.Cells(mHid(i, 5), colNiv).Value
            'Valida el valor de nivel
            respuesta = validaNivel3(valNiv, mHid(i, 0), fecha)
            
            If respuesta Then
                'Si el valor es para eliminar dato
                If valNiv = "" Or valNiv = "ddd" Or valNiv = "DDD" Then
                    dataBase.eliminarNiv mHid(i, 0), fecha
                Else
                    If mHid(i, 4) = 2 Then
                        dataBase.addNivel valNiv, mHid(i, 0), fecha
                    ElseIf mHid(i, 4) = 3 Then
                        dataBase.repNivel valNiv, mHid(i, 0), fecha
                    End If
                End If
                'Cambia el fondo de celda a blanco
                blanco CInt(mHid(i, 5))
            Else
                rojo CInt(mHid(i, 5))
                hayErrores = True
            End If
            mHid(i, 4) = 1
        End If
    Next i
    
    If hayErrores Then
        MsgBox "Algunos valores capturados de NIVEL no son correctos", vbCritical, "Error en captura"
    Else
        CapturaMatutino.actualizaHojas
    End If
End Sub

'************************************
' **    COLORES/ENFASIS EN HOJA    **
'************************************

Private Sub gris(fil As Integer, Optional col As Integer)
    hjHdr.Cells(fil, colNiv).Interior.Color = RGB(242, 242, 242)
    hjHdr.Cells(fil, colNiv).Font.Color = vbBlack
    hjHdr.Cells(fil, colNiv).Font.Bold = False
End Sub

Private Sub blanco(fil As Integer, Optional col As Integer)
    hjHdr.Cells(fil, colNiv).Interior.Color = vbWhite
    hjHdr.Cells(fil, colNiv).Font.Color = vbBlack
    hjHdr.Cells(fil, colNiv).Font.Bold = False
End Sub

Private Sub rojo(fil As Integer, Optional col As Integer)
    hjHdr.Cells(fil, colNiv).Interior.Color = vbRed
    hjHdr.Cells(fil, colNiv).Font.Color = vbBlack
    hjHdr.Cells(fil, colNiv).Font.Bold = False
End Sub
Private Sub naranja(fil As Integer, Optional col As Integer)
    hjHdr.Cells(fil, colNiv).Interior.Color = RGB(255, 192, 0)
    hjHdr.Cells(fil, colNiv).Font.Color = vbBlack
    hjHdr.Cells(fil, colNiv).Font.Bold = False
End Sub
Private Sub textoRojo(fil As Integer, Optional col As Integer)
    hjHdr.Cells(fil, colNiv).Interior.Color = RGB(255, 192, 0)
    hjHdr.Cells(fil, colNiv).Font.Color = RGB(192, 0, 0)
    hjHdr.Cells(fil, colNiv).Font.Bold = True
End Sub

