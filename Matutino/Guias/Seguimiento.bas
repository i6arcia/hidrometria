Attribute VB_Name = "Seguimiento"
Option Explicit
'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String
'Variables fecha
Private fecha As String
Private hora As String
'Constantes
Public dns As String
Public grp As String
'Contadores
Private i As Integer
Private j As Integer
'Variables de control de hoja
Private filIni As Integer
Private filFin As Integer
Private colIni As Integer
Private colFin As Integer
Private filFin2 As Integer
Private colFin2 As Integer
'Variable hoja de excel
Private seg As Excel.Worksheet
Private llu As Excel.Worksheet
'Bandera de control de cambios
Public bandera As Boolean

'###INICIA VARIABLES GLOBALES DEL MODULO###
Sub iniciaSeg()
    'Inicia variables de control de hoja
    varCtrlHoja
    
    'Escribe fecha hoja NIVELES
    seg.Range("B5").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & _
                            Format(Now, "dd") & " de " & Format(Now, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
    'Color de fondo en la celda fecha
    seg.Range("B5").Interior.color = RGB(255, 242, 204)
    
    'Escribe fecha hoja LLUVIA
    llu.Range("B5").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & _
                            Format(Now, "dd") & " de " & Format(Now, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
    'Color de fondo en la celda fecha
    llu.Range("B5").Interior.color = RGB(189, 215, 238)
    
    'Asigna fecha actual en una variable
    fecha = Format(Now, "yyyy/mm/dd")
    '*****Resguarda fecha en una celda
    seg.Range("AA1").Value = fecha
    
    'Almacena información de constantes
    If Hoja1.txbxDNS.Text <> "" Then
        'DNS para la base de datos
        dns = Hoja1.txbxDNS.Text
        Hoja1.txbxDNS.BackColor = vbWhite
    Else
        'Resalta en rojo
        Hoja1.txbxDNS.BackColor = vbRed
        'Termina proceso
        End
    End If
    If Hoja1.txbxGrp.Text <> "" Then
        'Grupo de estaciones para consultar en BD
        grp = Hoja1.txbxGrp.Text
        Hoja1.txbxGrp.BackColor = vbWhite
    Else
        'Resalta en rojo
        Hoja1.txbxGrp.BackColor = vbRed
        'Termina proceso
        End
    End If
    
End Sub

'Variables para el control de hoja
Private Sub varCtrlHoja()
    'Asigna hoja a la variable
    Set seg = Worksheets("Niveles")
    Set llu = Worksheets("Lluvia")
    'Inicia variables de control de hoja
    filIni = 8
    colIni = 2
    filFin = seg.Range("A" & rows.Count).End(xlUp).Row
    colFin = seg.Cells(filIni - 1, Columns.Count).End(xlToLeft).Column - 2
    
    filFin2 = llu.Range("A" & rows.Count).End(xlUp).Row
    colFin2 = llu.Cells(filIni - 1, Columns.Count).End(xlToLeft).Column - 3
End Sub

'### OBTIENE ESTACIONES DE LA BASE DE DATOS ###
Sub obtenerEst()
    'Verifica que las varibales globales esten iniciadas
    If filIni < 1 Then
        'Si no estan iniciadas la variables las inicia
        iniciaSeg
    End If
    '*Limpia contenido de la hoja NIVELES*
    If filFin >= filIni Then
        With seg.Range(seg.Cells(filIni, colIni - 1), seg.Cells(filFin + 1, colFin + 3))
            .ClearContents
            .Interior.color = xlNone
            .BorderAround xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
    End If

    '*Limpia contenido de la hoja LLUVIA*
    If filFin2 >= filIni Then
        With llu.Range(llu.Cells(filIni, colIni - 1), llu.Cells(filFin2 + 1, colFin2 + 3))
            .ClearContents
            .Interior.color = xlNone
            .BorderAround xlNone
            .Borders(xlInsideHorizontal).LineStyle = xlNone
            .Borders(xlInsideVertical).LineStyle = xlNone
        End With
    End If

    'Conecta con modulo ESTACIONES, obtiene la informacion de las estaciones
    'Obtiene nombre de las estaciones
    Estaciones.obtenerInf dns, grp
    'Escribe nombre de las estaciones en hoja NIVELES
    Estaciones.escribeNombres "Niveles", filIni, colIni - 1
    
    'Escribe nombre de las estaciones en hoja LLUVIA
    Estaciones.escribeNombres "Lluvia", filIni, colIni - 1
    
    'Actualiza fila fin
    filFin = seg.Range("A" & rows.Count).End(xlUp).Row
    filFin2 = llu.Range("A" & rows.Count).End(xlUp).Row
    
    'Formato para columna NAMO
    With seg.Range(seg.Cells(filIni, colFin + 2), seg.Cells(filFin, colFin + 2))
        .BorderAround xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.00"
        .Interior.color = RGB(255, 242, 204)
    End With
    'Escribir NAMO
    For i = filIni To filFin
        With seg.Cells(i, colFin + 2)
            .Value = Estaciones.clvEst(2, i - filIni)
            If Estaciones.clvEst(2, i - filIni) = "" Then
                .Interior.color = RGB(166, 166, 166)
            End If
        End With
    Next i
End Sub

'###OBTIENE NIVELES DE LAS ESTACIONES###
Sub obtenerNiveles()
    'Verifica que variables globales esten iniciadas
    If filIni < 1 Then
        iniciaSeg
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    'Abre conexión a BD
    dbSIH.Open
    'Limpia niveles en la hoja
    If filIni < filFin Then
        With seg.Range(seg.Cells(filIni, colIni), seg.Cells(filFin, colFin))
            .ClearContents
            .Interior.color = xlNone
            .Font.color = vbBlack
            .Font.Bold = False
        End With
    End If
    'Inicia en dimensiones matriz Control de cambios
    ReDim ctrlCambios.mCambios(colFin - colIni, filFin - filIni)
    'Ciclo escribe niveles por cada hora (Columna)
    For i = colIni To colFin
        'Almacena la hora
        hora = Format(seg.Cells(filIni - 1, i).Value, "hh:mm")
        'Verifica que sea un valor hora
        If (IsDate(hora)) Then
            'Asigna color de fondo gris a la celda hora
            gris i, filIni - 1, seg
            'Consulta niveles de la hora
            query = "SELECT t1.station, t1.valuee FROM DTNivel t1, stationgroups t2 WHERE t2.stationgroup = '" & _
                    grp & "'  and t1.station = t2.station AND datee = '" & fecha & " " & hora & "'"
            'Ejecuta consulta
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Inicia contador en fila de inicio
            j = filIni
                'Escribe mientras no este vacio respuesta consulta
                Do While Not adoRs.EOF
                    'Hace cruce de nombre de estaciones, para escribir nivel según corresponda
                    If adoRs!station = Estaciones.clvEst(0, j - filIni) Then
                        'Escribe nivel
                        seg.Cells(j, i).Value = Format(adoRs!Valuee, "0.00")
                        validarNiv.validaNamo Format(adoRs!Valuee, "0.00"), j - filIni
                        'Cambia el puntero al siguiente resultado de la consulta
                        Select Case validarNiv.edoValidacion
                            Case 2
                                proximo i, j, seg
                            Case 3
                                supero i, j, seg
                        End Select
                        adoRs.MoveNext
                        'Modifica matriz control de cambios
                        ctrlCambios.mCambios(i - colIni, j - filIni) = 1
                    End If
                    'Incrementa contador (Fila)
                    j = j + 1
                Loop
            'Cierra resultado de consulta
            adoRs.Close
        Else
            'Resalta en rojo
            rojo i, filIni - 1, seg
        End If
    Next i
    
    'Asigna formato de celdas y texto
    With seg.Range(seg.Cells(filIni, colIni), seg.Cells(filFin, colFin))
        .BorderAround xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.00"
    End With
    'Fin de la conexión
    dbSIH.Close
    
    'Obtener tendencia de los niveles
    For i = 0 To Estaciones.nEst - 1
        seg.Cells(i + filIni, colFin + 1).Value = validarNiv.validarTenGrl(i, fecha)
    Next i
    
End Sub

'###OBTIENE VALORES LLUVIA DE LAS ESTACIONES###
Sub obtenerLluvia()
    'Verifica que variables globales esten iniciadas
    If filIni < 1 Then
        iniciaSeg
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    'Abre conexión a BD
    dbSIH.Open
    'Limpia valores de lluvia en la hoja
    If filIni < filFin2 Then
        With llu.Range(llu.Cells(filIni, colIni), llu.Cells(filFin2, colFin2 + 3))
            .ClearContents
            .Interior.color = xlNone
            .Font.color = vbBlack
            .Font.Bold = False
        End With
    End If
    'Inicia en dimensiones matriz Control de cambios
    ReDim ctrlCambios.lCambios((colFin2 - colIni) + 3, filFin2 - filIni)
    
    '******************************
    '****Obtiene lluvias horaria***
    '******************************
    
    'Ciclo escribe valores de llvia por cada hora (Columna)
    For i = colIni To colFin2
        'Almacena la hora
        hora = Format(llu.Cells(filIni - 1, i).Value, "hh:mm")
        'Verifica que sea un valor hora
        If (IsDate(hora)) Then
            'Asigna color de fondo gris a la celda hora
            gris i, filIni - 1, llu
            'Cambia color a las celdas. Unicamente por referencia a las acumuladas
            If hora = "08:00" Or hora = "17:00" Then
                llu.Range(llu.Cells(filIni, i), llu.Cells(filFin2, i)).Interior.color = RGB(221, 235, 247)
            End If
            
            'Consulta datos lluvia
            query = "SELECT t1.station, t1.valuee FROM DTprecipitacio t1, stationgroups t2 WHERE t2.stationgroup = '" & _
                    grp & "'  and t1.station = t2.station AND datee = '" & fecha & " " & hora & "'"
            'Ejecuta consulta
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Inicia contador en fila de inicio
            j = filIni
                'Escribe mientras no este vacio respuesta consulta
                Do While Not adoRs.EOF
                    'Hace cruce de nombre de estaciones, para escribir nivel según corresponda
                    If adoRs!station = Estaciones.clvEst(0, j - filIni) Then
                        'Si lluvia es Inapreciable (0.01)
                        If adoRs!Valuee <= 0.01 And adoRs!Valuee > 0 Then
                            'Escribe Inap
                            llu.Cells(j, i).Value = "Inap"
                        Else
                            'Escribe dato
                            llu.Cells(j, i).Value = Format(adoRs!Valuee, "0.0")
                        End If
                        'Cambia el puntero al siguiente resultado de la consulta
                        adoRs.MoveNext
                        'Modifica matriz control de cambios
                        ctrlCambios.lCambios(i - colIni, j - filIni) = 1
                        'Reinicia puntero
                        j = filIni - 1
                    End If
                    'Incrementa contador (Fila)
                    j = j + 1
                Loop
            'Cierra resultado de consulta
            adoRs.Close
        Else
            'Resalta en rojo valor de fecha erroneo
            rojo i, filIni - 1, seg
        End If
    Next i
    
    'Asigna fondo gris a las columnas ACUMULADAS
    llu.Range(llu.Cells(filIni, colFin2 + 2), llu.Cells(filFin2, colFin2 + 3)).Interior.color = RGB(166, 166, 166)
    
    '******************************
    '**Obtiene lluvias acumuladas**
    '******************************
    
    'De 7 a 7 hrs (24 hrs)
    query = "SELECT t1.station, sum(t1.valuee) as suma FROM DTprecipitacio t1, stationgroups t2 WHERE t2.stationgroup = '" & _
            grp & "' and t1.station = t2.station AND datee > '" & Format(DateDiff("d", 1, fecha), "yyyy/mm/dd") & _
            " 08:00' AND datee <= '" & fecha & " 08:00' group by t1.station"
    'Ejecuta consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    'Inicia contador en fila de inicio
    j = filIni
        'Escribe mientras no este vacio respuesta consulta
        Do While Not adoRs.EOF
            'Hace cruce de nombre de estaciones, para escribir lluvia según corresponda
            If adoRs!station = Estaciones.clvEst(0, j - filIni) Then
                'Si lluvia es Inapreciable (0.01)
                If adoRs!suma <= 0.01 And adoRs!suma > 0 Then
                    'Escribe Inap
                    llu.Cells(j, colFin2 + 2).Value = "Inap"
                Else
                    'Escribe dato
                    llu.Cells(j, colFin2 + 2).Value = Format(adoRs!suma, "0.0")
                End If
                'Asigna color de fondo a la celda en AZUL
                llu.Cells(j, colFin2 + 2).Interior.color = RGB(221, 235, 247)
                'Cambia el puntero al siguiente resultado de la consulta
                adoRs.MoveNext
                'Modifica matriz control de cambios
                ctrlCambios.lCambios((colFin2 - colIni) + 2, j - filIni) = 1
                'Reinicia puntero
                j = filIni - 1
            End If
            'Incrementa contador (Fila)
            j = j + 1
        Loop
    'Cierra resultado de consulta
    adoRs.Close
    
    'Consulta lluvia acumulada
    'De 8 a 17 hrs
    query = "SELECT t1.station, sum(t1.valuee) as suma FROM DTprecipitacio t1, stationgroups t2 WHERE t2.stationgroup = '" & _
            grp & "'  and t1.station = t2.station AND datee >= '" & fecha & " 08:00' AND datee <= '" & fecha & " 17:00' group by t1.station"
    'Ejecuta consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    'Inicia contador en fila de inicio
    j = filIni
        'Escribe mientras no este vacio respuesta consulta
        Do While Not adoRs.EOF
            'Hace cruce de nombre de estaciones, para escribir nivel según corresponda
            If adoRs!station = Estaciones.clvEst(0, j - filIni) Then
                'Si lluvia es Inapreciable (0.01)
                If adoRs!suma <= 0.01 And adoRs!suma > 0 Then
                    'Escribe Inap
                    llu.Cells(j, colFin2 + 3).Value = "Inap"
                Else
                    'Escribe dato
                    llu.Cells(j, colFin2 + 3).Value = Format(adoRs!suma, "0.0")
                End If
                'Asigna color de fondo a la celda en AZUL
                llu.Cells(j, colFin2 + 3).Interior.color = RGB(221, 235, 247)
                'Cambia el puntero al siguiente resultado de la consulta
                adoRs.MoveNext
                'Modifica matriz control de cambios
                ctrlCambios.lCambios((colFin2 - colIni) + 3, j - filIni) = 1
                'Reinicia puntero
                j = filIni - 1
            End If
            'Incrementa contador (Fila)
            j = j + 1
        Loop
    'Cierra resultado de consulta
    adoRs.Close
    
    'Asigna formato a celdas tabla general
    With llu.Range(llu.Cells(filIni, colIni), llu.Cells(filFin2, colFin2))
        .BorderAround xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.0"
    End With
    'Asigna formato a celdas de las columnas Acumuladas
    With llu.Range(llu.Cells(filIni, colFin2 + 2), llu.Cells(filFin2, colFin2 + 3))
        .BorderAround xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .NumberFormat = "0.0"
        .Font.Bold = True
    End With
    
    'Fin de la conexión
    dbSIH.Close
End Sub


Sub guardarNiveles()
    'Bandera para validad error en el dato Nivel
    Dim corregirError As Boolean
    'Variable estado en tabla de modificados
    Dim edoDato As Integer
    'Almacena valor de nivel
    Dim niv As String
    'Validar variables globales esten iniciadas
    If fecha = "" Then
        'Inicia variables
        iniciaSeg
    End If
    
    'Recorre matriz modificados para guardar datos
    'Recorre columnas
    For i = colIni To colFin
        'Almacena valor de hora
        hora = Format(seg.Cells(filIni - 1, i).Value, "hh:mm")
        'Verifica que sea un valor de hora correcto
        If (IsDate(hora)) Then
            'Recorre filas
            For j = filIni To filFin
                'Valor ESTADO en la matriz control de cambios
                'ctrlCambios.modificado j - filIni, i - colIni
                edoDato = ctrlCambios.mCambios(i - colIni, j - filIni)
                'Si estado es 2 Agregado o 3 modificado
                If edoDato = 2 Or edoDato = 3 Then
                    'Inicia variable NO HAY ERROR
                    corregirError = False
                    'Almacena valor de niv3el
                    niv = seg.Cells(j, i).Value
                    'Si el valor es vacio o comando para eliminar /DDD/
                    If niv = "" Or niv = "ddd" Or niv = "DDD" Then        'Validar para ELIMINAR valor
                        'Si es el estado es modificado
                        If edoDato = 3 Then
                            'Elimina valor de la base de datos
                            dataBase.eliminarNiv j - filIni, fecha + " " + hora
                        End If
                        'Devuelve el fondo de celda a blanco
                        blanco i, j, seg
                    'Valida que el valor del nivel sea numerico
                    ElseIf (IsNumeric(niv)) Then    'Validar para CAPTURAR valor
                        'Establece formato
                        niv = Format(niv, "0.00")
                        'Devuelve el dondo de celda a blanco
                        blanco i, j, seg
                        'Si el estado es Agregado
                        If edoDato = 2 Then
                            'Agrega nivel a la BD
                            dataBase.addNivel niv, j - filIni, fecha & " " & hora
                        'Si el estado es modificado
                        ElseIf edoDato = 3 Then
                            'Remplaza el valor en la BD
                            dataBase.repNivel niv, j - filIni, fecha & " " & hora
                        End If
                        'Valida si hubo error
                        Select Case validarNiv.edoValidacion
                            Case 0  'SI HUBO ERROR
                                'Cambia fondo de celda a rojo
                                rojo i, j, seg
                                'Bandera ERROR cambia a verdadero
                                corregirError = True
                            Case 1
                                'Cambia fondo de la celda a blanco
                                blanco i, j, seg
                                'Escribe el valor de la tendencia del río
                                bandera = True
                                seg.Cells(j, colFin + 1).Value = validarNiv.tendencia
                                bandera = False
                        End Select
                        'Otros valores del estado de validación provocan error
                        If validarNiv.edoValidacion <> 1 Then corregirError = True
                    Else    'ERROR no es numerico ni comando de eliminar /DDD/
                        'Cambia fondo de celda a rojo
                        rojo i, j, seg
                        'Bandera ERROR cambia a verdadero
                        corregirError = True
                    End If
                End If
            Next j
        Else
            'La celda hora no tiene el formato correcto
            MsgBox "El valor hora no es correcto", vbCritical
            'Cambia fondo de la celda a rojo
            rojo i, filIni - 1, seg
        End If
    Next i
    'En caso de no encontrar errores en la captura
    If Not corregirError Then
        bandera = True
        'Actualiza los niveles
        obtenerNiveles
        bandera = False
    End If
End Sub

Sub guardarLluvia()
    'Bandera control de errores
    Dim corregirError As Boolean
    'Estado del dato en la BD 0Vacio|2Agregar|3Modificar
    Dim edoDato As Integer
    'Almacena valor de lluvia
    Dim lluvia As String
    'Almacena lluvia acumulada
    Dim Acumulada As String
    
    'Validar que variables globales estén iniciadas
    If fecha = "" Then
        iniciaSeg
    End If
    '***********************************
    '*Recorre matriz control de cambios*
    '***********************************
    For i = colIni To colFin2
        'Almacena valor de hora
        hora = Format(llu.Cells(filIni - 1, i).Value, "hh:mm")
        'Valida que sea valor hora correcto
        If (IsDate(hora)) Then
                'Indice para valores de filas
                For j = filIni To filFin2
                    'Obtiene el estado del dato en la BD segun controlDeCambios 0Vacio|2Agregar|3Modificar
                    edoDato = ctrlCambios.lCambios(i - colIni, j - filIni)
                    'Si estado corresponde a 2Agregar|3Modificar
                    If edoDato = 2 Or edoDato = 3 Then
                        'Inicia Bandera NO EXISTE ERROR
                        corregirError = False
                        'Almacena valor de lluvia
                        lluvia = llu.Cells(j, i).Value
                        '*************************
                        '*Valida valor de lluvia*
                        '*************************
                        If lluvia = "" Or lluvia = "ddd" Or lluvia = "DDD" Then        'Comandos para eliminar
                            'Si es el estado es modificado
                            If edoDato = 3 Then
                                'Elimina valor de la base de datos
                                dataBase.eliminarLluvia j - filIni, fecha + " " + hora
                            End If
                            'Cambia color de fondo
                            blanco i, j, llu
                        Else    'Valida valor de lluvia
                            If Not (IsNumeric(lluvia)) Then  'Valida NO es numerico
                                'Valida si el valor es alguna variante de Inapreciable
                                If (lluvia = "inap" Or lluvia = "INAP" Or lluvia = "Inap") Then
                                    'Asigna valor numerico de Inapreciable = 0.01
                                    lluvia = 0.01
                                    'Cambia color fondo de celda
                                    blanco i, j, llu
                                Else
                                    'Cambia color fondo celda
                                    rojo i, j, llu
                                    'Bandera ERROR ES VERDADERO
                                    corregirError = True
                                End If
                            ElseIf (CDbl(lluvia) >= 0) Then 'Valida mayor a Cero 0
                                If (CDbl(lluvia) <> 0.01) Then  'Diferente de inapreciable
                                    'Asigna formato al valor de lluvia
                                    lluvia = Format(lluvia, "0.0")
                                    'Cambia color fondo celda
                                    blanco i, j, llu
                                Else
                                    'Bandera modificar hoja
                                    bandera = True
                                    'Cambia valor numerico a Inap
                                    llu.Cells(j, i).Value = "Inap"
                                    'Cierra bande modificar hoja
                                    bandera = False
                                    'Asigna formato al valor de lluvia (0.01)
                                    lluvia = Format(lluvia, "0.00")
                                End If
                            Else
                                'Cambia color fondo de celda
                                rojo i, j, llu
                                'Vandera ERROR VERDADERO
                                corregirError = True
                            End If
                            '*******************************
                            'Enviar valor a la base de datos
                            '*******************************
                            'Si no existe error
                            If Not corregirError Then
                                If edoDato = 2 Then     'AGREGAR
                                    dataBase.addLluvia lluvia, j - filIni, fecha & " " & hora
                                ElseIf edoDato = 3 Then 'REMPLAZAR
                                    dataBase.repLluvia lluvia, j - filIni, fecha & " " & hora
                                End If
                            End If
                        End If
                    End If
                Next j
        Else
            'La celda hora no tiene el formato correcto
            MsgBox "El valor hora no es correcto", vbCritical
            rojo i, filIni - 1, seg
        End If
    Next i
    
    
    '****************************************
    'RECORRE CONTROL DE CAMBIOS EN ACUMULADAS
    '****************************************
    For i = 2 To 3
        For j = filIni To filFin2
            'Obtiene el estado del dato en la BD segun controlDeCambios 0Vacio|2Agregar|3Modificar
            edoDato = ctrlCambios.lCambios((colFin2 - colIni) + i, j - filIni)
            If edoDato = 2 Or edoDato = 3 Then
                'Inicia Bandera NO EXISTE ERROR
                corregirError = False
                '**********Almacena valor de llvia*********
                '(EN ESTE CASO DEBE SER LA LLUVIA ACUMULADA)
                lluvia = llu.Cells(j, i + colFin2).Value
                'Valida valor de lluvia
                If lluvia = "" Or lluvia = "ddd" Or lluvia = "DDD" Then        'Comando para ELIMINAR dato
                    If edoDato = 3 Then     'Si el estado es modificado 3
                        MsgBox "Únicamente se eliminará el residuo del TOTAL de lluvia menos la lluvia acumulada previa." + vbLf + "Lluvia acumulada previa permanecerá.", vbInformation
                        'Elimina valor de la base de datos
                        If i = 2 Then dataBase.eliminarLluvia j - filIni, fecha + " 08:00"
                        If i = 3 Then dataBase.eliminarLluvia j - filIni, fecha + " 17:00"
                    End If
                    'Cambia color fondo celda
                    azul colFin2 + i, j, llu
                Else    'Valida valor de lluvia
                    If Not (IsNumeric(lluvia)) Then  'Validar No es Numerico
                        'Valida si el valor es alguna variante de Inapreciable
                        If (lluvia = "inap" Or lluvia = "INAP" Or lluvia = "Inap") Then
                            'Asigna valor numerico de Inapreciable (0.01)
                            lluvia = 0.01
                            'Cambia color fondo de celda
                            azul colFin2 + i, j, llu
                        Else
                            'Cambia color fondo celda
                            rojo colFin2 + i, j, llu
                            'Banderra ERROR VERDADERO
                            corregirError = True
                        End If
                    ElseIf (CDbl(lluvia) >= 0) Then 'Valor es mayor a 0
                        'Caso de ser valor inapeciable (0.01)
                        If (CDbl(lluvia) <> 0.01) Then
                            'Asigna formato al valor de lluvia
                            lluvia = Format(lluvia, "0.0")
                            'Cambia color fondo de celda
                            azul colFin2 + i, j, llu
                        Else
                            'Bandera modificar hoja
                            bandera = True
                            'Cambia valor numerico a Inap
                            llu.Cells(j, i + colFin2).Value = "Inap"
                            'Cierra bandera modificar hoja
                            bandera = False
                            'Asigna valor numerico de inap
                            lluvia = 0.01
                        End If
                    Else
                        'Cambia color fondo de celda
                        rojo colFin2 + i, j, llu
                        'Bandera ERROR VERDADERO
                        corregirError = True
                    End If
                    '*******************************
                    '----VALIDA LLUVIA ACUMULADA----
                    '*******************************
                    'Si no existe error
                    If Not corregirError Then
                        'Obtiene lluvia acumulada
                        If i = 2 Then
                            Acumulada = dataBase.lluviaAcumulada(Estaciones.clvEst(0, j - filIni), Format(DateDiff("d", 1, fecha), "yyyy/mm/dd 08:00"), Format(fecha, "yyyy/mm/dd 08:00"))
                        ElseIf i = 3 Then
                            Acumulada = dataBase.lluviaAcumulada(Estaciones.clvEst(0, j - filIni), Format(fecha, "yyyy/mm/dd 08:00"), Format(fecha, "yyyy/mm/dd 17:00"))
                        End If
                        'Valida la lluvia total con respecto a la acumulada
                        If Acumulada = "" Then Acumulada = 0    'Asigna 0 si no existe acumulada
                        'Verifica que acumulada sea menor a la lluvia que se va a capturar
                        If CDbl(lluvia) >= CDbl(Acumulada) Then
                            'Almacena el residuo de lluvia total con la acumulada previa
                            lluvia = Round(CDbl(lluvia) - CDbl(Acumulada), 1)
                        Else    'ERROR
                            'Cambia color fondo de celda
                            rojo colFin2 + i, j, llu
                            'Bandera ERROR VERDADERO
                            corregirError = True
                        End If
                    End If
                    '*******************************
                    'Enviar valor a la base de datos
                    '*******************************
                    If i = 2 Then hora = "08:00"
                    If i = 3 Then hora = "17:00"
                    'Si no existe error
                    If Not corregirError Then
                        If edoDato = 2 Then 'Agrega dato
                            dataBase.addLluvia lluvia, j - filIni, fecha & " " & hora
                        ElseIf edoDato = 3 Then 'Remplaza dato
                            dataBase.repLluvia lluvia, j - filIni, fecha & " " & hora
                        End If
                    End If
                End If
            End If
        Next j
    Next i
    
    '*****************************************************
    'Vuelve a obtener valores de lluvia de la base de datos
    '          En caso de no tener algún error
    '******************************************************
    If Not corregirError Then
        bandera = True
        obtenerLluvia
        bandera = False
    End If
End Sub
Sub setFecha(fch As String)
    'Asigna hoja a la variable
    Set seg = Worksheets("Niveles")
    fecha = Format(fch, "yyyy/mm/dd")
    seg.Range("AA1").Value = fecha
End Sub
Public Function getFecha() As String
    getFecha = fecha
End Function

Private Sub rojo(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.color = vbRed
End Sub
Private Sub blanco(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.color = xlNone
End Sub
Private Sub proximo(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.color = vbYellow
    ws.Cells(rows, col).Font.Bold = True
End Sub
Private Sub supero(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.color = vbYellow
    ws.Cells(rows, col).Font.color = vbRed
    ws.Cells(rows, col).Font.Bold = True
End Sub
Private Sub gris(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.color = RGB(242, 242, 242)
End Sub

Private Sub azul(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.color = RGB(221, 235, 247)
End Sub

Public Function getFilIni() As Integer
    If filIni = 0 Then
        varCtrlHoja
    End If
    getFilIni = filIni
End Function

Public Function getColFin2() As Integer
    If filIni = 0 Then
        varCtrlHoja
    End If
    getColFin2 = colFin2
End Function

Public Function getColIni() As Integer
    If colIni = 0 Then
        varCtrlHoja
    End If
    getColIni = colIni
End Function

Public Function getRanEdit() As Range
    If filIni = 0 Then
        iniciaSeg
    End If
        Set getRanEdit = Range(Cells(filIni, colIni), Cells(filFin, colFin))
End Function

Public Function getRanEditLlivia() As Range
    If filIni = 0 Then
        iniciaSeg
    End If
        Set getRanEditLlivia = Range(Cells(filIni, colIni), Cells(filFin2, colFin2))
End Function

Public Function getRanEditLlivia2() As Range
    If filIni = 0 Then
        iniciaSeg
    End If
        Set getRanEditLlivia2 = Range(Cells(filIni, colFin2 + 2), Cells(filFin2, colFin2 + 3))
End Function
