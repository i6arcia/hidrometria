Attribute VB_Name = "Presas"
'****************************************
'****************************************
'Control y Captura de informaciÓn PRESAS
'****************************************
'****************************************

Option Explicit

Private hjPrs As Excel.Worksheet

'Control de manejo de hoja
Private filIni As Integer
Private filFin As Integer
Private colIni As Integer
Private colFin As Integer
Private colClv As Integer
Private colNivAyer As Integer
Private colNiv As Integer
Private colGasto As Integer
Private colOt As Integer

'Variables contadores y banderas
Private i As Integer

'variable control de presas
Dim derivadasPrs(20, 8)
Dim imDP As Integer     'Contador matriz Prs
Dim hdCFE(10, 2) As String
Dim imCFE As Integer

Sub iniciaPresas()
    'Asigna hojas a variables
    Set hjPrs = Worksheets("PRESAS")
    'Control de hoja
    filIni = 12
    filFin = 65
    colClv = 1
    colNiv = 7
    colNivAyer = 6
    colGasto = 10
    colOt = 11
    
    imDP = 0
    imCFE = 0
    
'Matriz de datos derivados de presas
'Tabla para completar presas (DERIVADAS)
'Se activará junto con el calculo automatico de almacenamiento

'   clv Estacion |   Nivel  | clvDerArea   |  clvDerAlmac  |ObraToma|Vertedor|Derrame | AutoCompletar |   Fila  |  Nombre de la estación
' 0    "LREHD"   |    Val   |     x        | "LREPBELEVA"  |  val   |   d0   |   d0   |      0        |         |  LOS REYES
' 1    "LLAHD"   |    Val   |     x        | "LLAPBELEVA"  |  val   |   d0   |   d0   |      0        |         |  LA LAGUNA
' 2    "NNEPB"   |    Val   |     x        | "NNEPBELEVA"  |   x    |   d0   |   d0   |      0        |         |  NECAXA
' 3    "NEXPB"   |    Val   |     x        | "NEXPBELEVA"  |  val   |   d0   |   d0   |      0        |         |  NEXAPA
' 4    "TENPB"   |    Val   |     x        | "TENPBELEVA"  |  val   |   d0   |   d0   |      0        |         |  TENANGO
' 5    "SOLPB"   |    Val   |     x        | "SOLPBELEVA"  |  val   |   d0   |   d0   |      0        |         |  LA SOLEDAD
' 6    "TEMOX"   |    Val   |"TEMOXAREAE"  | "TEMOXELEVA"  |  val   |   d0   |   d0   |      0        |         |  MIGUEL ALEMAN (TEMASCAL)
' 7    "CDOOX"   |    Val   |"CDOOXAREAE"  | "CDOOXELEVA"  |  d0    |   val  |   d0   |      0        |         |  MIGUEL DE LA MADRID (CERRO DE ORO)
' 8    "PEMVC"   |    Val   |"PEMVCAREAE"  | "PEMVCELEVA"  |  val   |    x   |   x0   |      0        |         |  EL MORALILLO
' 9    "LDCVC"   |    Val   |"LDCVCNIVAREA"| "LDCVCELEVA"  |  val   |    x   |   x0   |      0        |         |  LAGUNA DE CATEMACO (CANSECO)
'10    "LCAVC"   |    Val   |"LCAVCAREAE"  | "LCAVCELEVA"  |  val   |   d0   |   d0   |      0        |         |  LA CANGREJERA

'x -> Vacio     'Val -> Valor Capturado en celda       'd0 -> Derivado (0)
'AutoCompletar -> Indice que lanzara la accion de llenado

'Envia datos a procedimiento para introducir valores
setDerivadasPrs "LREHD", "", "x", "LREPBELEVA", "v", "d", "d", 0, 12              'LOS REYES
setDerivadasPrs "LLAHD", "", "x", "LLAPBELEVA", "v", "d", "d", 0, 14               'LA LAGUNA
setDerivadasPrs "NNEPB", "", "x", "NNEPBELEVA", "x", "d", "d", 0, 16               'NECAXA
setDerivadasPrs "NEXPB", "", "x", "NEXPBELEVA", "v", "v", "d", 0, 18               'NEXAPA
setDerivadasPrs "TENPB", "", "x", "TENPBELEVA", "v", "d", "d", 0, 20               'TENANGO
setDerivadasPrs "SOLPB", "", "x", "SOLPBELEVA", "v", "d", "d", 0, 22               'LA SOLEDAD
setDerivadasPrs "TEMOX", "", "TEMOXAREAE", "TEMOXELEVA", "v", "d", "d", 0, 27      'MIGUEL ALEMAN (TEMASCAL)
setDerivadasPrs "CDOOX", "", "CDOOXAREAE", "CDOOXELEVA", "d", "v", "d", 0, 29      'MIGUEL DE LA MADRID (CERRO DE ORO)
setDerivadasPrs "PEMVC", "", "PEMVCAREAE", "PEMVCELEVA", "v", "x", "x", 0, 36      'EL MORALILLO
setDerivadasPrs "LDCVC", "", "LDCVCNIVAREA", "LDCVCELEVA", "v", "x", "x", 0, 44   'LAGUNA DE CATEMACO (CANSECO)
setDerivadasPrs "LCAVC", "", "LCAVCAREAE", "LCAVCELEVA", "v", "d", "d", 0, 46      'LA CANGREJERA
End Sub

Sub limpiaHojaPresas()
    If filIni = 0 Then iniciaPresas
    
    CapturaMatutino.editando = True
    '******HOJA PRESAS********
    hjPrs.Range("B5").Value = "Xalapa, Ver. -- --"
    'Lluvias Patla y Tepexi
    hjPrs.Range("D2:D3").ClearContents
    'Puebla
    hjPrs.Range("F12:J23,L18:L19,K22:L23,N16,N22:N23").ClearContents
    hjPrs.Range("J18").Formula = "=L18+L19"
    hjPrs.Range("J22").Formula = "=ROUND(K22/86400,3)"
    hjPrs.Range("H12").Formula = "=G12/D12"
    hjPrs.Range("H14").Formula = "=G14/D14"
    hjPrs.Range("H16").Formula = "=G16/D16"
    hjPrs.Range("H18").Formula = "=G18/D18"
    hjPrs.Range("H20").Formula = "=G20/D20"
    hjPrs.Range("H22").Formula = "=G22/D22"
    'Oaxaca
    hjPrs.Range("F27:J30").ClearContents
    hjPrs.Range("H27").Formula = "=G27/D27"
    hjPrs.Range("H29").Formula = "=G29/D29"
    'Veracruz
    hjPrs.Range("F34:L47,N44").ClearContents
    hjPrs.Range("J38").Formula = "=ROUND(K38/86400,3)"
    hjPrs.Range("J42").Formula = "=ROUND(K42/86400,3)"
    hjPrs.Range("J44").Formula = "=ROUND(K44/86400,3)"
    hjPrs.Range("N38").Formula = "=E58"                'Formulas para El Encanto
    hjPrs.Range("N39").Formula = "=(N38/(2*36))^(2/3)" 'Formulas para El Encanto
    If hjPrs.Range("E58").Value <> "" Then             'Escala el encanto
        'hjPrs.Range("G39").Formula = "=+N39+323"
    Else
        hjPrs.Range("G39").Formula = ""
    End If
    hjPrs.Range("H36").Formula = "=G36/D36"
    hjPrs.Range("H46").Formula = "=G46/D46"
    hjPrs.Range("G39").Interior.Color = xlNone
    hjPrs.Range("J46").Formula = "=L53"
    'Parte inferior
    hjPrs.Range("F51:H52").ClearContents
    hjPrs.Range("L50:L53").ClearContents
    hjPrs.Range("L53").Formula = "=SUM(L50:L52)"
    hjPrs.Range("D56:G61,J56:L59").ClearContents
    
    'Area editable (AZUL)
    'Patlas
    hjPrs.Range("D2:D3").Interior.Color = RGB(197, 217, 241)
    'Tecolutla
    hjPrs.Range("G13,G15,G17,G19,G21,G23,J12:J15,J20:J21,L18:L19,K22:L23,N16,N22:N23").Interior.Color = RGB(197, 217, 241)
    'Papaloapan
    hjPrs.Range("G28,G30,J27:J30").Interior.Color = RGB(197, 217, 241)
    'Panuco-coatza
    hjPrs.Range("G37,G41,G43,G45,G47,J36:J37,K38:L39, K42:L45,N44").Interior.Color = RGB(197, 217, 241)
    'Plantas Cangrejera
    hjPrs.Range("F51:H51,F52,H52,L50:L53").Interior.Color = RGB(197, 217, 241)
    'Hidrometricas CFE
    hjPrs.Range("D56:F61,G58,G61,J56:L59").Interior.Color = RGB(197, 217, 241)
    CapturaMatutino.editando = False
End Sub

'###Obtener datos de presas###
Sub obtenerPresas()
    If filIni = 0 Then iniciaPresas
    
    CapturaMatutino.editando = True
    lluviasEvaporacion
    nivAlmAyer
    nivAlmHoy
    gastos
    PBs
    hidroCFE
    CapturaMatutino.editando = False
End Sub

'#############################################
'#############################################
'#############################################

Private Sub lluviasEvaporacion()
    Dim valLluvia As String
    Dim valEvaporacion As String
    
    If filIni = 0 Then iniciaPresas
    
    'Consulta lluvia de estacion PATLA
    valLluvia = dataBase.getLluvia("PTLPB", CapturaMatutino.fecha)
    ctrlCambios.buscayEscribe "PTLPB", valLluvia, 4
    'Consulta lluvia de estacion TEPEXI
    valLluvia = dataBase.getLluvia("TEPPB", CapturaMatutino.fecha)
    ctrlCambios.buscayEscribe "TEPPB", valLluvia, 4
    'Consulta lluvia de presa Necaxa
    valLluvia = dataBase.getLluvia("NNEPB", CapturaMatutino.fecha)
    ctrlCambios.buscayEscribe "NNEPB", valLluvia, 4
    'Consulta lluvia de presa La Soledad
    valLluvia = dataBase.getLluvia("SOLPB", CapturaMatutino.fecha)
    ctrlCambios.buscayEscribe "SOLPB", valLluvia, 4
    'Consulta Evaporación de presa La Soledad
    valEvaporacion = dataBase.getEvaporacion("SOLPB", CapturaMatutino.fecha)
    ctrlCambios.buscayEscribe "SOLPB", valEvaporacion, 5
    'Consulta lluvia de presa Laguna de Catemaco
    valLluvia = dataBase.getLluvia("LDCVC", CapturaMatutino.fecha)
    ctrlCambios.buscayEscribe "LDCVC", valLluvia, 4
End Sub

Private Sub nivAlmAyer()
    If filIni = 0 Then iniciaPresas
    
    'Consulta valores de nivel día Ayer
    dataBase.getNivelGrp "GRGCPRESAS", ayer
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscaEscribeAyer dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 1, colNivAyer
        Next i
    End If
    'Consulta valores de almacenamietos día de Ayer
    dataBase.getVolAlmGrp "GRGCPRESAS", ayer
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscaEscribeAyer dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 2, colNivAyer
        Next i
    End If
End Sub

Private Sub nivAlmHoy()
    If filIni = 0 Then iniciaPresas
    
    'Consulta valores de nivel día Hoy
    dataBase.getNivelGrp "GRGCPRESAS", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 1
        Next i
    End If
    'Consulta valores de almacenamietos día de Ayer
    dataBase.getVolAlmGrp "GRGCPRESAS", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 2
        Next i
    End If
End Sub
Private Sub gastos()
    Dim vert As String
    
    If filIni = 0 Then iniciaPresas

    'Consulta valores de nivel día Hoy
    dataBase.getGastoGrp "GRGCPRESAS", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            'Excepcion // replica valor de gasto en la cangrejera
            If dataBase.mValGrp(i, 0) = "LCAVC" Then
                CapturaMatutino.editando = True
                hjPrs.Range("L53").Value = dataBase.mValGrp(i, 1)
                CapturaMatutino.editando = False
            End If
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 3
        Next i
    End If
    'Consulta VERTEDOR de Nexapa
    vert = dataBase.getVertedor("NEXPB", fecha)
    ctrlCambios.buscayEscribe "NEXPB", vert, 6
    
    'Consulta VERTEDOR de cerro de oro
    vert = dataBase.getVertedor("CDOOX", fecha)
    ctrlCambios.buscayEscribe "CDOOX", vert, 6
End Sub

Private Sub PBs()
    If filIni = 0 Then iniciaPresas
    
    Dim valLluvia As String
    Dim valNiv As String
    'Lluvia de PB1
    valLluvia = dataBase.getLluvia("PCNVC", fecha)
    ctrlCambios.buscayEscribe "PCNVC", valLluvia, 4
    'Lluvia de PB2
    valLluvia = dataBase.getLluvia("CB2VC", fecha)
    ctrlCambios.buscayEscribe "CB2VC", valLluvia, 4
    'Lluvia de PB3
    valLluvia = dataBase.getLluvia("PB3VC", fecha)
    ctrlCambios.buscayEscribe "PB3VC", valLluvia, 4
    
    'Consulta niveles de las plantas
    'Nivel de PB1
    valNiv = dataBase.getNivel("PCNVC", fecha)
    ctrlCambios.buscayEscribe "PCNVC", valNiv, 1
    'Nivel de PB3
    valNiv = dataBase.getNivel("PB3VC", fecha)
    ctrlCambios.buscayEscribe "PB3VC", valNiv, 1
End Sub

Private Sub hidroCFE()
    If filIni = 0 Then iniciaPresas
    
    'Consulta Niveles
    dataBase.getNivelGrp "hidrocfe", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 1
        Next i
    End If
    'Consulta gastos
    dataBase.getGasRioGrp "hidrocfe", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 8
        Next i
    End If
    'Consulta lluvia
    dataBase.getLluviaGrp "hidrocfe", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 4
        Next i
    End If
    'Consulta Evaporacion
    dataBase.getEvapGrp "hidrocfe", fecha
    If dataBase.iValG > 0 Then
        For i = 0 To dataBase.iValG - 1
            ctrlCambios.buscayEscribe dataBase.mValGrp(i, 0), dataBase.mValGrp(i, 1), 5
        Next i
    End If
End Sub

'Indica que la estacion se va a Auto completar
Sub siCompletar(clv As String, niv As String)
    'Recorre matriz
    For i = 0 To imDP - 1
        'Busca la posicion en la matriz con la clave
        If derivadasPrs(i, 0) = clv Then
            derivadasPrs(i, 7) = 1      'Cambia valor a positivo
            derivadasPrs(i, 1) = niv    'Asigna valor Nivel
            Exit For
        End If
    Next i
End Sub
'Auto completa datos de las presas
Sub autoCompletar()
    Dim area As String
    Dim almacenamiento As String
    
    'Recorre matriz derivadas Presas
    For i = 0 To imDP - 1
        'Verifica que el campo auto completar sea positivo
        If derivadasPrs(i, 7) = 1 Then
            'Valida que exista un nivel para calcular variables derivadas
            If derivadasPrs(i, 1) <> "" Then
                '----Campo AREA----
                If derivadasPrs(i, 2) <> "x" Then   'Clave para obtener area de presa derivado
                    area = dataBase.getDerArea(CStr(derivadasPrs(i, 2)), CStr(derivadasPrs(i, 1)))
                    validacion.validaArea CStr(derivadasPrs(i, 0)), area
                End If
                '----Campo Almacenamiento----
                If derivadasPrs(i, 3) <> "" Then   'Clave para obtener Almacenamiento derivado
                End If
                '----Campo Obra Toma (GASTO)----
                If derivadasPrs(i, 4) = "d" Then   'Valida capturar derivada = 0
                    dataBase.repGasto "0", CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
                '----Campo Vertedor----
                If derivadasPrs(i, 5) = "d" Then   'Valida capturar derivada = 0
                    dataBase.repVertedor "0", CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
                '----Campo Derrame----
                If derivadasPrs(i, 6) = "d" Then   'Valida capturar derivada = 0
                    dataBase.repDerrame "0", CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
            Else        'ELMINA VALORES AUTOCOMPLETADOS
                MsgBox "Se eliminaran datos autocompletados"
                If derivadasPrs(i, 2) <> "x" Then   'Clave para obtener area de presa derivado
                    dataBase.eliminaArea CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
                '----Campo Almacenamiento----
                If derivadasPrs(i, 3) <> "" Then   'Clave para obtener Almacenamiento derivado
                End If
                '----Campo Obra Toma (GASTO)----
                If derivadasPrs(i, 4) = "d" Then   'Valida capturar derivada = 0
                    dataBase.eliminarGasto CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
                '----Campo Vertedor----
                If derivadasPrs(i, 5) = "d" Then   'Valida capturar derivada = 0
                    dataBase.eliminarVertedor CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
                '----Campo Derrame----
                If derivadasPrs(i, 6) = "d" Then   'Valida capturar derivada = 0
                    dataBase.eliminarDerrame CStr(derivadasPrs(i, 0)), CapturaMatutino.fecha
                End If
            End If
        End If
    Next i
End Sub



'---------------------------------------------------------------------------------------
'                               SET / GET
'---------------------------------------------------------------------------------------

'Procedimiento que llena la matriz derivadaPrs
Private Sub setDerivadasPrs(clv As String, valNiv As String, clvArea As String, clvAlm As String, ot As String, vertedor As String, derrame As String, autoComp As String, fila As String)
    'derivadasPrs{
    '   0| clv Estacion
    '   1| Valor Nivel
    '   2| Clave Derivada Area
    '   3| Clave Derivada Almacenamiento
    '   4| Obra Toma
    '   5| Vertedor
    '   6| Derrame
    '   7| Auto completar
    '   8| Fila en la hoja de Excel
    '}
    derivadasPrs(imDP, 0) = clv
    derivadasPrs(imDP, 1) = valNiv
    derivadasPrs(imDP, 2) = clvArea
    derivadasPrs(imDP, 3) = clvAlm
    derivadasPrs(imDP, 4) = ot
    derivadasPrs(imDP, 5) = vertedor
    derivadasPrs(imDP, 6) = derrame
    derivadasPrs(imDP, 7) = autoComp
    derivadasPrs(imDP, 8) = fila
    imDP = imDP + 1
End Sub

Private Sub setHdCFE(clv As String, fila As String, columna As String)
    'derivadasPrs{
    '   0| clv Estacion
    '   1| Fila en la hoja de Excel
    '   2| Columna en la hoja de Excel
    '}
    hdCFE(imCFE, 0) = clv
    hdCFE(imCFE, 1) = fila
    hdCFE(imCFE, 2) = columna
    imCFE = imCFE + 1
End Sub

''Obtiene todas las variables derivadas de la informacion en hoja
'Sub varDerivadas()
'    ' Variables
'    Dim nivel As String
'
'
'    inicia
'
'    'Conexión a la base de datos
'    dbSIH.ConnectionString = dns
'    'Abre conexión a BD
'    dbSIH.Open
'
'    For i = 0 To 10
'        nivel = Presas.Cells(prs(i, 2) + 1, colNiv).Value
'        If nivel <> "" Then
'            'Consulta almacenamietos
'            '        SELECT * FROM transformtables where xvalue= '57.47' and transformtableid = 'cdooxeleva'
'            query = "select * from transformtables where xvalue= '" & nivel & "' and transformtableid = '" & prs(i, 1) & "'"
'            'Ejecuta consulta
'            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
'            If Not adoRs.EOF Then
'                Presas.Cells(prs(i, 2), colNiv).Value = Format(adoRs!Yvalue, "0.000")
'            End If
'            'Cierra resultado de consulta
'            adoRs.Close
'        End If
'    Next i
'
'    'Fin de la conexión
'    dbSIH.Close
'
'End Sub
'Sub capturaPresas()
'
'    inicia
'
'    'Variables
'    Dim niv As String
'    Dim alm As String
'
'    niv = Presas.Cells(prs(0, 2) + 1, colNiv)
'    If niv <> "" Then
'        dataBase.repNivel niv, prs(0, 0), fecha
'    End If
'End Sub
