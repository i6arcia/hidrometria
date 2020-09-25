Attribute VB_Name = "ctrlCambios"
Option Explicit

Public mCambios(100, 6)
Public iMtz As Integer
Public rgo As Range
Public rgoEdos As Range
Public rgoDeri As Range
Public rgoTomata As Range

Public rgoHidro As Range

Dim prs As Excel.Worksheet
Dim i As Integer

Public Sub llenaMatriz()
    iMtz = 0
    ' *** mCambios {
    '       0 clvEst |
    '       1 Fila |
    '       2 Columna |
    '       3 tipoD |
    '       4 edo  |
    '       5 GeneraDerivada  |
    '       6 clvDerivada)
    '   }***

    Set prs = Worksheets("PRESAS")
    
    setValMatriz "PTLPB", 2, 4, 4, 0        'PATLA
    setValMatriz "TEPPB", 3, 4, 4, 0        'TEPEXI
    'Los reyes
    setValMatriz "LREHD", 13, 7, 1, 0, 1, "LREPBELEVA"      'Nivel
    setValMatriz "LREHD", 12, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "LREHD", 12, 10, 3, 0                      'Gasto
    'La Laguna
    setValMatriz "LLAHD", 15, 7, 1, 0, 1, "LLAPBELEVA"      'Nivel
    setValMatriz "LLAHD", 14, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "LLAHD", 14, 10, 3, 0                      'Gasto
    'Necaxa
    setValMatriz "NNEPB", 17, 7, 1, 0, 1, "NNEPBELEVA"      'Nivel
    setValMatriz "NNEPB", 16, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "NNEPB", 16, 14, 4, 0                      'Lluvia
    'Nexapa
    setValMatriz "NEXPB", 19, 7, 1, 0, 1, "NEXPBELEVA"      'Nivel
    setValMatriz "NEXPB", 18, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "NEXPB", 19, 12, 3, 0, 1                   'Gasto Deriva
    setValMatriz "NEXPB", 18, 12, 6, 0, 1                   'Vertedor
    setValMatriz "NEXPB", 18, 10, 3, 0, 2                   'Gasto derivado
    'Tenango
    setValMatriz "TENPB", 21, 7, 1, 0, 1, "TENPBELEVA"      'Nivel
    setValMatriz "TENPB", 20, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "TENPB", 20, 10, 3, 0                      'Gasto
    'La soledad
    setValMatriz "SOLPB", 23, 7, 1, 0, 1, "SOLPBELEVA"      'Nivel
    setValMatriz "SOLPB", 22, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "SOLPB", 22, 10, 3, 0, 2                   'Gasto
    setValMatriz "SOLPB", 22, 14, 4, 0                      'Lluvia
    setValMatriz "SOLPB", 23, 14, 5, 0                      'Evaporación
    setValMatriz "SOLPB", 22, 11, 7, 0, 1                   'O.T.
    'MIGUEL ALEMAN (TEMASCAL)
    setValMatriz "TEMOX", 28, 7, 1, 0, 1, "TEMOXELEVA"      'Nivel
    setValMatriz "TEMOX", 27, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "TEMOX", 27, 10, 3, 0                      'Gasto
    setValMatriz "TEMOX", 1, 1, 0, 0, 2, "TEMOXAREAE"       'Área
    'MIGUEL DE LA MADRID (CERRO DE ORO)
    setValMatriz "CDOOX", 30, 7, 1, 0, 1, "CDOOXELEVA"      'Nivel
    setValMatriz "CDOOX", 29, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "CDOOX", 29, 10, 6, 0                      'Vertedor
    setValMatriz "CDOOX", 1, 1, 0, 0, 2, "CDOOXAREAE"       'Área
    'Chicayan
    'El moralillo
    setValMatriz "PEMVC", 37, 7, 1, 0, 1, "PEMVCELEVA"      'Nivel
    setValMatriz "PEMVC", 36, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "PEMVC", 36, 10, 3, 0                      'Gasto
    setValMatriz "PEMVC", 1, 1, 0, 0, 2, "PEMVCAREAE"       'Área
    'El encanto
    setValMatriz "PECVC", 38, 14, 8, 0, 2                   'Tomata Gasto
    setValMatriz "PECVC", 39, 14, 8, 0, 2                   'Tirante
    setValMatriz "PECVC", 39, 7, 1, 0, 2                    'Nivel
    setValMatriz "PECVC", 38, 10, 3, 0, 2                   'Gasto
    setValMatriz "PECVC", 38, 11, 7, 0, 1                   'O.T.
    'Camelpo
    setValMatriz "DCAVC", 41, 7, 1, 0                       'Nivel
    'Tuxpango
    setValMatriz "PTXVC", 43, 7, 1, 0                       'Nivel
    setValMatriz "PTXVC", 42, 10, 3, 0, 2                   'Gasto
    setValMatriz "PTXVC", 42, 11, 7, 0, 1                   'O.T.
    'Laguna de catemaco
    setValMatriz "LDCVC", 45, 7, 1, 0, 1, "LDCVCELEVA"      'Nivel
    setValMatriz "LDCVC", 44, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "LDCVC", 44, 10, 3, 0, 2                   'Gasto
    setValMatriz "LDCVC", 44, 11, 7, 0, 1                   'O.T.
    setValMatriz "LDCVC", 44, 14, 4, 0                      'Lluvia
    setValMatriz "LDCVC", 1, 1, 0, 0, 2, "LDCVCNUVAREA"     'Área
    'La cangrejera
    setValMatriz "LCAVC", 47, 7, 1, 0, 1, "LCAVCELEVA"      'Nivel
    setValMatriz "LCAVC", 46, 7, 2, 0, 2                    'Almacenamiento
    setValMatriz "LCAVC", 46, 10, 3, 0, 2                   'Gasto
    setValMatriz "LCAVC", 1, 1, 0, 0, 2, "LCAVCAREAE"         'Área
    'Plantas La cangrejeta
    ''PCNVC', 'CB2VC', 'PB3VC'
    setValMatriz "PCNVC", 52, 6, 1, 0       'Nivel
    setValMatriz "PCNVC", 51, 6, 4, 0       'Lluvia
    setValMatriz "CB2VC", 51, 7, 4, 0       'Lluvia
    setValMatriz "PB3VC", 52, 8, 1, 0       'Nivel
    setValMatriz "PB3VC", 51, 8, 4, 0       'Lluvia
    'Gastos Presa la cangrejera
    setValMatriz "LCAVC", 50, 12, 9, 0, 1                   'Gasto Pemex
    setValMatriz "LCAVC", 51, 12, 9, 0, 1                   'Gasto CNA
    setValMatriz "LCAVC", 52, 12, 9, 0, 1                   'Gasto Morelos
    setValMatriz "LCAVC", 53, 12, 3, 0, 2                   'Gasto TOTAL
    
    '******Presas parte inferior*********
    'Buenos Aires
    setValMatriz "BAIPB", 56, 4, 1, 0       'Nivel
    setValMatriz "BAIPB", 56, 5, 8, 0       'Gasto
    setValMatriz "BAIPB", 56, 6, 4, 0       'Lluvia
    'Sontalaco
    setValMatriz "STLPB", 57, 4, 1, 0       'Nivel
    setValMatriz "STLPB", 57, 5, 8, 0       'Gasto
    setValMatriz "STLPB", 57, 6, 4, 0       'Lluvia
    'Tomata
    setValMatriz "TOMVC", 58, 4, 1, 0       'Nivel
    setValMatriz "TOMVC", 58, 5, 8, 0, 1       'Gasto
    setValMatriz "TOMVC", 58, 6, 4, 0       'Lluvia
    setValMatriz "TOMVC", 58, 7, 5, 0       'Evaporación
    'Usila
    setValMatriz "UCFOX", 59, 4, 1, 0       'Nivel
    setValMatriz "UCFOX", 59, 5, 8, 0       'Gasto
    setValMatriz "UCFOX", 59, 6, 4, 0       'Lluvia
    'Stp. Domingo
    setValMatriz "STDOX", 60, 4, 1, 0       'Nivel
    setValMatriz "STDOX", 60, 5, 8, 0       'Gasto
    setValMatriz "STDOX", 60, 6, 4, 0       'Lluvia
    'Pecaditos
    setValMatriz "DPSOX", 61, 4, 1, 0       'Nivel
    setValMatriz "DPSOX", 61, 5, 8, 0       'Gasto
    setValMatriz "DPSOX", 61, 6, 4, 0       'Lluvia
    setValMatriz "DPSOX", 61, 7, 5, 0       'Evaporación
    'Naranjastitlan
    setValMatriz "NARPB", 56, 10, 1, 0      'Nivel
    setValMatriz "NARPB", 56, 11, 8, 0      'Gasto
    setValMatriz "NARPB", 56, 12, 4, 0      'Lluvia
    'Chicomapa
    setValMatriz "CHMVC", 57, 10, 1, 0      'Nivel
    setValMatriz "CHMVC", 57, 11, 8, 0      'Gasto
    setValMatriz "CHMVC", 57, 12, 4, 0      'Lluvia
    'Tepeyac
    setValMatriz "TYAPB", 58, 10, 1, 0      'Nivel
    setValMatriz "TYAPB", 58, 11, 8, 0      'Gasto
    setValMatriz "TYAPB", 58, 12, 4, 0      'Lluiva
    'Cosalapa
    setValMatriz "CSPOX", 59, 10, 1, 0      'Nivel
    setValMatriz "CSPOX", 59, 11, 8, 0      'Gasto
    setValMatriz "CSPOX", 59, 12, 4, 0      'Lluvia
End Sub

Public Sub setValMatriz(clv As String, fil As Integer, col As Integer, tipo As Integer, edo As Integer, Optional deriva As Integer, Optional clvDer As String)
    ' *** mCambios (0 clvEst | 1 Fila | 2 Columna | 3 tipoD | 4 edo  | 5 GeneraDerivada  | 6 clvDerivada) ***
    mCambios(iMtz, 0) = clv         'Clave
    mCambios(iMtz, 1) = fil         'Fila
    mCambios(iMtz, 2) = col         'Columna
    mCambios(iMtz, 3) = tipo        'Tipo de dato
    mCambios(iMtz, 4) = edo         'Estado
    mCambios(iMtz, 5) = deriva      'GeneraDerivada
    mCambios(iMtz, 6) = clvDer      'clvDerivada
    iMtz = iMtz + 1
    '****************************
    'Tipo dato{
        ' 1|  Nivel
        ' 2|  Almacenamiento
        ' 3|  Gasto
        ' 4|  Lluvia
        ' 5|  Evaporación
        ' 6|  Vertedor
        ' 7|  O.T.
        ' 8|  GastoRio (Para estacion CFE)
        ' 9|  Gasto PB's La cangrejera
        ' 0|  Área de la presa
    '}
    'Estado dato{
        '0|  Vacio
        '1|  Con dato
        '-----------------
        '2|  Add
        '3|  Modificar
    '}
    '****************************
End Sub

Public Sub generaRangos()
    If mCambios(0, 0) = "" Then
        llenaMatriz
    End If
    
    Set rgo = prs.Cells(mCambios(0, 1), mCambios(0, 2))
    Set rgoDeri = prs.Cells(1, 1)
    Set rgoTomata = prs.Cells(58, 5)
    For i = 0 To iMtz - 1
        If mCambios(i, 5) = 2 Then
            If mCambios(i, 3) <> 8 Then
                Set rgoDeri = Union(rgoDeri, prs.Cells(mCambios(i, 1), mCambios(i, 2)))
            ElseIf mCambios(i, 3) = 8 Then
                Set rgoTomata = Union(rgoTomata, prs.Cells(mCambios(i, 1), mCambios(i, 2)))
            End If
        Else
            Set rgo = Union(rgo, prs.Cells(mCambios(i, 1), mCambios(i, 2)))
        End If
    Next i
End Sub
Sub modificado(fil As Integer, col As Integer)
' *** mCambios {
'            | 0 clvEst     |
'            | 1 Fila       |
'            | 2 Columna    |
'            | 3 tipoD      |
'            | 4 edo        |
'            | 5 GeneraDerivada  |
'            | 6 clvDerivada|
'       }***

    Dim deri As String
    
    For i = 0 To iMtz
        If mCambios(i, 1) = fil Then
            If mCambios(i, 2) = col Then
                If mCambios(i, 4) = 1 Then
                    mCambios(i, 4) = 3
                ElseIf mCambios(i, 4) = 0 Then
                    mCambios(i, 4) = 2
                End If
                
                gris fil, col
                
                'Si el valor genera derivadas
                If mCambios(i, 5) = 1 Then
                    Select Case mCambios(i, 3)
                        Case 1 'Nivel
                            deri = dataBase.getDerAlmace(CStr(mCambios(i, 6)), prs.Cells(fil, col).Value)
                            escribeAlmaDeri CStr(mCambios(i, 0)), deri
                            Presas.siCompletar CStr(mCambios(i, 0)), prs.Cells(fil, col).Value
                        Case 7  'O.T -> Gasto
                            getGasto CStr(mCambios(i, 0)), CDbl(prs.Cells(fil, col).Value)
                        Case 9
                            gastoSumado CStr(mCambios(i, 0))
                        Case 6
                            CapturaMatutino.editando = True
                            prs.Range("J18").Formula = "=L18+L19"
                            CapturaMatutino.editando = False
                        Case 8
                            tomata
                            modificado 39, 7
                    End Select
                End If
'                MsgBox "Variable ENCONTRADA en: " & CStr(i) & vbLf _
'                        & "Estacion: " & mCambios(i, 0) & vbLf _
'                        & "Fil: " & CStr(mCambios(i, 1)) & " Col: " & CStr(mCambios(i, 2)) & vbLf _
'                        & "Tipo_ " & CStr(mCambios(i, 3)) & " Edo: " & CStr(mCambios(i, 4))
                Exit For
            End If
        End If
    Next i
End Sub

Sub tomata()
    CapturaMatutino.editando = True
    'Gasto de tomata
    prs.Range("N38").Formula = "=E58"
    'prs.Range("N38").NumberFormat = "0.000 m³/s"
    'Tirante
    prs.Range("N39").Formula = "=(N38/(2*36))^(2/3)"
    prs.Range("N39").NumberFormat = "0.00"
    'Nivel de EL ENCANTO
    If prs.Range("E58").Value <> "" Then
        prs.Range("G39").Formula = "=N39+323"
        prs.Range("G39").NumberFormat = "0.00"
    Else
        prs.Range("G39").Formula = ""
    End If
    CapturaMatutino.editando = False
End Sub

'Busca si existe un valor editado en la matriz
Sub tieneEditados()
    Dim dato As String
    Dim respuesta As Boolean
    Dim errores As Boolean
'****************************
'Tipo dato{
    '1|  Nivel
    '2|  Almacenamiento
    '3|  Gasto
    '4|  Lluvia
    '5|  Evaporación
    '6|  Vertedor
    '7|  O.T. 2
    '8|  Gasto en Río
'}
    'Recorre todo el arreglo
    For i = 0 To iMtz
        'Valida que existen valores modificados
        If mCambios(i, 4) = 2 Or mCambios(i, 4) = 3 Then
            'Recupera el valor de la celda
            dato = prs.Cells(mCambios(i, 1), mCambios(i, 2)).Value
            Select Case mCambios(i, 3)  'Tipo de dato
                Case 1      'Nivel
                    respuesta = validacion.validaNivel(dato, CStr(mCambios(i, 0)), fecha, CStr(mCambios(i, 4)))
                Case 2      'Almacenamiento
                    respuesta = validacion.validaAlmacenamiento(dato, CStr(mCambios(i, 0)), fecha, CStr(mCambios(i, 4)))
                Case 3      'Gasto
                    validacion.validaGasto i
                    respuesta = True
                Case 4      'Lluvia
                    validacion.validaLluviaPresas i
                    respuesta = True
                Case 5      'Evaporación
                    validacion.validaEvapPresas i
                    respuesta = True
                Case 6      'Vertedor
                    validacion.validaVertedor i
                    respuesta = True
                Case 7      'Obra toma 2
                    'validacion.validaOt2 i
                Case 8      'Gasto en rio
                    validacion.validaGastoRio i
                    respuesta = True
            End Select
            If Not respuesta Then
                rojo CInt(mCambios(i, 1)), CInt(mCambios(i, 2))
                mCambios(i, 4) = 3
                MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
                errores = True
            End If
        End If
    Next i
    If Not errores Then CapturaMatutino.actualizaHojas
End Sub

Sub escribeAlmaDeri(clv As String, val As String)
    CapturaMatutino.editando = True
    For i = 0 To iMtz
        If mCambios(i, 0) = clv Then
            If mCambios(i, 3) = 2 Then
                prs.Cells(mCambios(i, 1), mCambios(i, 2)).Value = val
                If mCambios(i, 4) = 1 Then
                    mCambios(i, 4) = 3
                Else
                    mCambios(i, 4) = 2
                End If
                Exit For
            End If
        End If
    Next i
    CapturaMatutino.editando = False
End Sub
Sub escribeNivDeri(clv As String, val As String)
    CapturaMatutino.editando = True
    For i = 0 To iMtz
        If mCambios(i, 0) = clv Then
            If mCambios(i, 3) = 1 Then
                prs.Cells(mCambios(i, 1), mCambios(i, 2)).Value = val
                Exit For
            End If
        End If
    Next i
    CapturaMatutino.editando = False
End Sub
Sub buscayEscribe(clv As String, val As String, tipoDato As Integer)
'Tipo dato{
    '1|  Nivel
    '2|  Almacenamiento
    '3|  Gasto
    '4|  Lluvia
    '5|  Evaporación
    '6|  Vertedor
    '7|  O.T. 2
    '8|  Gasto en Río
'}
    If val <> "" Then
        CapturaMatutino.editando = True
        For i = 0 To iMtz
            If mCambios(i, 0) = clv Then
                If mCambios(i, 3) = tipoDato Then
                    prs.Cells(mCambios(i, 1), mCambios(i, 2)).Value = val
                    'Cambia el estado a modificado
                    mCambios(i, 4) = 1
                    Exit For
                End If
            End If
        Next i
        CapturaMatutino.editando = False
    End If
End Sub
Sub buscaEscribeAyer(clv As String, val As String, tipoDato As Integer, col As Integer)
'Tipo dato{
    '1|  Nivel
    '2|  Almacenamiento
    '3|  Gasto
    '4|  Lluvia
    '5|  Evaporación
    '6|  Vertedor
    '7|  O.T. 2
    '8|  Gasto en Río
'}
    If val <> "" Then
        CapturaMatutino.editando = True
        For i = 0 To iMtz
            If mCambios(i, 0) = clv Then
                If mCambios(i, 3) = tipoDato Then
                    prs.Cells(mCambios(i, 1), col).Value = val
                    Exit For
                End If
            End If
        Next i
        CapturaMatutino.editando = False
    End If
End Sub
Sub getGasto(clv As String, val As Double)
    CapturaMatutino.editando = True
    For i = 0 To iMtz
        If mCambios(i, 0) = clv And mCambios(i, 3) = 3 Then
            prs.Cells(mCambios(i, 1), mCambios(i, 2)).Value = Round(val / 86400, 3)
            If mCambios(i, 4) = 1 Then
                mCambios(i, 4) = 3
            Else
                mCambios(i, 4) = 2
            End If
            Exit For
        End If
    Next i
    CapturaMatutino.editando = False
End Sub
Sub gastoSumado(clv As String)
    CapturaMatutino.editando = True
    If clv = "NEXPB" Then
        prs.Range("J18").Formula = "=L18+L19"
    ElseIf clv = "LCAVC" Then
        'presas.Range("L53").Formula = "=SUM(L50:L52)"
        prs.Range("L53").Formula = "=SUM(L50:L52)"
        prs.Range("J46").Value = prs.Range("L53").Value
        For i = 0 To iMtz
            If mCambios(i, 0) = clv And mCambios(i, 3) = 3 And mCambios(i, 5) = 2 Then
                If mCambios(i, 4) = 1 Then
                    mCambios(i, 4) = 3
                Else
                    mCambios(i, 4) = 2
                End If
                'MsgBox "Ejemplo " & CStr(prs.Range("L53").Value)
                Exit For
            End If
        Next i
    End If
    CapturaMatutino.editando = False
End Sub
Sub conDato(clv As String, tipo As Integer)
    For i = 0 To iMtz
        If mCambios(i, 0) = clv Then
            If mCambios(i, 3) = tipo Then
                mCambios(i, 4) = 1
                Exit For
            End If
        End If
    Next i
End Sub


Private Sub gris(fil As Integer, col As Integer)
    Dim hj As Excel.Worksheet
    Set hj = Worksheets("PRESAS")
    hj.Cells(fil, col).Interior.Color = RGB(242, 242, 242)
End Sub

Private Sub rojo(fil As Integer, col As Integer)
    Dim hj As Excel.Worksheet
    Set hj = Worksheets("PRESAS")
    hj.Cells(fil, col).Interior.Color = vbRed
End Sub

