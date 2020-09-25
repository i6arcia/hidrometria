Attribute VB_Name = "pluviometros"
Option Explicit

'***************************************
'***************************************
'*********** PLUVIOMETROS **************
'***************************************
'***************************************

'Hojas de excel
Dim hjPluvios As Excel.Worksheet

'Control de manejo de hoja
Private filIni As Integer
Private filFin As Integer

Private colClv As Integer
Private colEstacion As Integer
Private colEdo As Integer
Private colCuenca As Integer
Private colLluvia As Integer

Private valLluvia As String

'Contadores y banderas
Private i As Integer

'Variable control de presas
Dim mPluvios(100, 3) As String
Dim imPluvios As Integer
'Rango
Public rgoPluvios As Range

Sub iniciaPluvios()
Dim clv As String
    'Asigna hojas a variables
    Set hjPluvios = Worksheets("Pluviometros")
    'Control de hoja
    filIni = 8
    filFin = hjPluvios.Range("C" & Rows.Count).End(xlUp).Row
    colClv = 3
    colEstacion = 4
    colEdo = 5
    colCuenca = 6
    colLluvia = 7
    
    'Contador matriz control
    imPluvios = 0
    
'   Matriz para el control de datos en la hoja
'*****************************************************
'| clave | Fila  | Lluvia | edoVariable
'*****************************************************
    For i = filIni To filFin
        clv = hjPluvios.Cells(i, colClv).Value
        If clv <> "" Then
            setMPluvios clv, i, "", "0"
        End If
    Next i
    'Genera el rango control de edición de la hoja
    rangoPluvios
End Sub

Sub limpiaPluvios()
    If filIni = 0 Then iniciaPluvios
    
    CapturaMatutino.editando = True
    
    '****** TITULO DE HOJA CLIMA NORTE ********
    hjPluvios.Range("B5").Value = "Xalapa, Ver. -- --"
    
    hjPluvios.Range(hjPluvios.Cells(filIni, colLluvia), hjPluvios.Cells(filFin, colLluvia)).ClearContents
    
    CapturaMatutino.editando = False
End Sub

Private Sub rangoPluvios()
    'Verifica que esten iniciadas las variables
    If filIni = 0 Then iniciaPluvios
    '
    Set rgoPluvios = hjPluvios.Range(hjPluvios.Cells(filIni, colLluvia), hjPluvios.Cells(filFin, colLluvia))
End Sub

Sub obtienePluvios()
    If filIni = 0 Then iniciaPluvios
    
    CapturaMatutino.editando = True
    
    infEstaciones
    
    PluviosLluvia
    
    CapturaMatutino.editando = False
End Sub

Private Sub infEstaciones()

    Dim cadena As String
    
    If filIni = 0 Then iniciaPluvios
    
    For i = 0 To imPluvios - 1
        cadena = dataBase.getNombreEstacion(mPluvios(i, 0))
        hjPluvios.Cells(mPluvios(i, 1), colEstacion).Value = cadena
        cadena = dataBase.getNombreEstado(mPluvios(i, 0))
        hjPluvios.Cells(mPluvios(i, 1), colEdo).Value = cadena
        cadena = dataBase.getNombreCuenca(mPluvios(i, 0))
        hjPluvios.Cells(mPluvios(i, 1), colCuenca).Value = cadena
    Next i
    
End Sub

Private Sub PluviosLluvia()
    If filIni = 0 Then iniciaPluvios
    
    For i = 0 To imPluvios - 1
        valLluvia = dataBase.getLluvia(mPluvios(i, 0), fecha)
        If valLluvia <> "" Then
            hjPluvios.Cells(mPluvios(i, 1), colLluvia).Value = valLluvia
            mPluvios(i, 3) = 1
        End If
    Next i
End Sub

Public Sub modificadoPluvios(fil As Integer)
    
    Dim colorCelda As String
    Dim posicion As Integer

    If imPluvios = 0 Then CapturaMatutino.actualizaHojas

    For i = 0 To imPluvios - 1
        If mPluvios(i, 1) = fil Then
            If mPluvios(i, 3) = 1 Then
                mPluvios(i, 3) = 3
                'MsgBox "Modificar"
            ElseIf mPluvios(i, 3) = 0 Then
                mPluvios(i, 3) = 2
                'MsgBox "Agregar"
            End If
            gris fil, colLluvia
            Exit For
        End If
    Next i

End Sub

Public Sub tieneEditadosPluvios()
    Dim respuesta As Boolean
    Dim hayErrores As Boolean
    
    'Valida que esten iniciadas las variables
    If filIni = 0 Then iniciaPluvios
    
    hayErrores = False
    
    'Recorre el arreglo de estaciones
    For i = 0 To imPluvios - 1
        'Si hay valores para 2|Agregar o 3|Remplazar
        If mPluvios(i, 3) = 3 Or mPluvios(i, 3) = 2 Then
            'Obtiene el valor de lluvia capturado en la celda
            valLluvia = hjPluvios.Cells(mPluvios(i, 1), colLluvia).Value
            'Valida el valor de lluvia
            respuesta = validacion.validaLluvia2(valLluvia)
            
            If respuesta Then
                If IsNumeric(valLluvia) Then
                    If CDbl(valLluvia) <= 0.01 Then
                        valLluvia = "0.01"
                        If mPluvios(i, 3) = 2 Then
                            dataBase.addLluvia valLluvia, mPluvios(i, 0), fecha
                        ElseIf mPluvios(i, 3) = 3 Then
                            dataBase.repLluvia valLluvia, mPluvios(i, 0), fecha
                        End If
                    Else
                        If mPluvios(i, 3) = 2 Then
                            dataBase.addLluvia valLluvia, mPluvios(i, 0), fecha
                        ElseIf mPluvios(i, 3) = 3 Then
                            dataBase.repLluvia valLluvia, mPluvios(i, 0), fecha
                        End If
                    End If
                    'Cambia el fondo de celda a blanco
                    blanco CInt(mPluvios(i, 1))
                ElseIf valLluvia = "inap" Or valLluvia = "INAP" Or valLluvia = "Inap" Then
                    valLluvia = "0.01"
                    If mPluvios(i, 3) = 2 Then
                        dataBase.addLluvia valLluvia, mPluvios(i, 0), fecha
                    ElseIf mPluvios(i, 3) = 3 Then
                        dataBase.repLluvia valLluvia, mPluvios(i, 0), fecha
                    End If
                    'Cambia el fondo de celda a blanco
                    blanco CInt(mPluvios(i, 1))
                ElseIf valLluvia = "" Or valLluvia = "ddd" Or valLluvia = "DDD" Then
                    If mPluvios(i, 3) = 3 Then
                        dataBase.eliminarLluvia mPluvios(i, 0), fecha
                    End If
                    'Cambia el fondo de celda a blanco
                    blanco CInt(mPluvios(i, 1))
                Else
                    'ERROR
                    rojo CInt(mPluvios(i, 1))
                    hayErrores = True
                End If
            Else
                rojo CInt(mPluvios(i, 1))
                hayErrores = True
            End If
            mPluvios(i, 3) = 1
        End If
    Next i
    
    If hayErrores Then
        MsgBox "Algunos valores capturados de lluvia no son correctos", vbCritical, "Error en captura"
    Else
        CapturaMatutino.actualizaHojas
    End If
End Sub

'*********************
' **    GET//SET    **
'*********************
Private Sub setMPluvios(clv As String, fil As Integer, lluv As String, edoVariable As String)
    'Matriz para Pluvios
    'mPluvios{
    '   0 | Clave
    '   1 | Fila
    '   2 | Lluvia
    '   3 | Estado de la variable
    '}
    
    mPluvios(imPluvios, 0) = clv
    mPluvios(imPluvios, 1) = fil
    mPluvios(imPluvios, 2) = lluv
    mPluvios(imPluvios, 3) = edoVariable
    
    imPluvios = imPluvios + 1
End Sub

'***************************************
'//////////////////////////////////////
'           ENFASIS EN HOJA
'//////////////////////////////////////
'***************************************

Private Sub gris(fil As Integer, Optional col As Integer)
    hjPluvios.Cells(fil, colLluvia).Interior.Color = RGB(242, 242, 242)
    hjPluvios.Cells(fil, colLluvia).Font.Color = vbBlack
    hjPluvios.Cells(fil, colLluvia).Font.Bold = False
End Sub

Private Sub blanco(fil As Integer, Optional col As Integer)
    hjPluvios.Cells(fil, colLluvia).Interior.Color = vbWhite
    hjPluvios.Cells(fil, colLluvia).Font.Color = vbBlack
    hjPluvios.Cells(fil, colLluvia).Font.Bold = False
End Sub

Private Sub rojo(fil As Integer, Optional col As Integer)
    hjPluvios.Cells(fil, colLluvia).Interior.Color = vbRed
    hjPluvios.Cells(fil, colLluvia).Font.Color = vbBlack
    hjPluvios.Cells(fil, colLluvia).Font.Bold = False
End Sub

Private Sub amarillo(fil As Integer, Optional col As Integer)
    hjPluvios.Cells(fil, colLluvia).Interior.Color = vbYellow
    hjPluvios.Cells(fil, colLluvia).Font.Color = vbBlack
    hjPluvios.Cells(fil, colLluvia).Font.Bold = False
End Sub
Private Sub naranja(fil As Integer, Optional col As Integer)
    hjPluvios.Cells(fil, colLluvia).Interior.Color = RGB(255, 192, 0)
    hjPluvios.Cells(fil, colLluvia).Font.Color = vbBlack
    hjPluvios.Cells(fil, colLluvia).Font.Bold = False
End Sub
Private Sub textoRojo(fil As Integer, Optional col As Integer)
    hjPluvios.Cells(fil, colLluvia).Interior.Color = RGB(255, 192, 0)
    hjPluvios.Cells(fil, colLluvia).Font.Color = RGB(192, 0, 0)
    hjPluvios.Cells(fil, colLluvia).Font.Bold = True
End Sub
