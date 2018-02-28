Attribute VB_Name = "Presas"
'*********************************************************************
'           Sistema en apoyo a la captura de información de
'           presas de la Dirección Técnica en OCGC, CONAGUA
'
'                               PRESAS
'
'*********************************************************************
Option Explicit

'Variables para conexión a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String

'Otras variables
Private fecha As String
Private fechaD As String
Private ultFil As Integer
Private inicioFil As Integer
Private prs As Excel.Worksheet
'Variables de comparacion (Último nivel y Desviación estandar)
Private varCom(6) As Double

Sub inicio()
    'Asigna hoja Presas a una variable
    Set prs = Worksheets("Presas")
    'Coloca fecha y lugar como titulo de la hoja
    prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    fecha = Format(Now, "yyyy/mm/dd")
    'Recupera datos de SIH
    obtenerDatos
End Sub

Sub obtenerDatos()
    'Variables
    Dim clvPrs(4) As String
    Dim hora As String
    Dim colV As Integer
    Dim colH As Integer
    Dim i As Integer
    'Clave SIh para las presas
    clvPrs(0) = "CDOOX" 'Cerro de Oro
    clvPrs(1) = "LCAVC" 'La Cangrejera
    clvPrs(2) = "PCNVC" 'La Cangrejera PB1
    clvPrs(3) = "CB2VC" 'La Cangrejera PB2
    clvPrs(4) = "PB3VC" 'La Cangrejera PB3
    'Fila donde inicia la captura de datos
    inicioFil = 12
    'Confirma asignación de hoja Presas a la variable
    Set prs = Worksheets("Presas")
    'Obtiene número de la última fila con datos
    ultFil = prs.Range("B" & rows.Count).End(xlUp).Row
    
    If (fecha = "") Then
        'Asigna fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    'Limpia contenido
    prs.Range("C" & inicioFil & ":G" & ultFil).ClearContents
    prs.Range("J" & inicioFil & ":R" & ultFil).ClearContents
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene datos de la presa Cerro de Oro
    colH = 2
    For i = inicioFil To ultFil
        colV = 4
        hora = Format(prs.Cells(i, colH).Value, "hh:mm")
        If (IsDate(hora)) Then
            blanco colH, i
            'Consulta nivel de Presa Cerro de oro
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(0) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Almacenamiento de Presa Cerro de oro
            query = "SELECT valuee FROM DTVolAlmac WHERE station = '" & clvPrs(0) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Gasto de Presa Cerro de oro
            query = "SELECT valuee FROM DTVertedor WHERE station = '" & clvPrs(0) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Lluvia de Presa Cerro de oro
            query = "SELECT valuee FROM DTPrecipitacio WHERE station = '" & clvPrs(0) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                        prs.Cells(i, colV).Value = "Inap"
                    Else
                        prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.0")
                    End If
                End If
            adoRs.Close
        Else
            rojo colH, i
        End If
    Next i
    
    'Obtiene datos de presa La Cangrejera
    colH = 9
    For i = inicioFil To ultFil
        colV = 10
        hora = Format(prs.Cells(i, colH).Value, "hh:mm")
        If (IsDate(hora)) Then
            blanco colH, i
            'Consulta nivel de Presa Cerro de oro
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(1) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Almacenamiento de Presa Cerro de oro
            query = "SELECT valuee FROM DTVolAlmac WHERE station = '" & clvPrs(1) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.000")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Lluvia de Presa Cerro de oro
            query = "SELECT valuee FROM DTPrecipitacio WHERE station = '" & clvPrs(1) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                        prs.Cells(i, colV).Value = "Inap"
                    Else
                        prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.0")
                    End If
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta nivel de PB1
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(2) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta lluvia de PB1
            query = "SELECT valuee FROM DTPrecipitacio WHERE station = '" & clvPrs(2) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                        prs.Cells(i, colV).Value = "Inap"
                    Else
                        prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.0")
                    End If
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta lluvia de PB2
            query = "SELECT valuee FROM DTPrecipitacio WHERE station = '" & clvPrs(3) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                        prs.Cells(i, colV).Value = "Inap"
                    Else
                        prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.0")
                    End If
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta nivel de PB3
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(4) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta lluvia de PB3
            query = "SELECT valuee FROM DTPrecipitacio WHERE station = '" & clvPrs(4) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                        prs.Cells(i, colV).Value = "Inap"
                    Else
                        prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.0")
                    End If
                End If
            adoRs.Close
        Else
            'Error en formato de hora
            rojo colH, i
        End If
    Next i
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub capturarDatos()
    'Variables para datos
    Dim niv As String
    Dim alm As String
    Dim ver As String
    Dim llu As String
    'Variables para datos de plantas
    Dim p1N As String
    Dim p1L As String
    Dim p2L As String
    Dim p3N As String
    Dim p3L As String
    'Otras variables
    Dim clvPrs(4) As String
    Dim hora As String
    Dim colH As Integer
    Dim colV As Integer
    Dim i As Integer
    Dim comp(2) As Double
    'Banderas control de errores
    Dim erf As Boolean 'Error en formato
    Dim erc As Boolean 'Error en calculo

    'Confirma asignacion de hoja a la variable
    Set prs = Worksheets("Presas")
    
    'Obtiene desviacion estandar de las variables para validación de datos
    desStd
    
    'Claves SIH para presas
    clvPrs(0) = "CDOOX" 'Cerro de oro
    clvPrs(1) = "LCAVC" 'La cangrejera
    clvPrs(2) = "PCNVC" 'La cangrejera PB1
    clvPrs(3) = "CB2VC" 'La cangrejera PB2
    clvPrs(4) = "PB3VC" 'La cangrejera PB3
    
    'Número de fila que inicia con datos
    inicioFil = 12
    
    'Obtiene el número de la última fila con datos
    ultFil = prs.Range("B" & rows.Count).End(xlUp).Row
    'Banderas control de errores
    erf = False
    erc = False
    
    If (fecha = "") Then
        'Asigna la fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    'Captura en SIH datos de Cerro de oro
    colH = 2
    colV = 4
    For i = inicioFil To ultFil
        erf = False
        hora = Format(prs.Cells(i, colH).Value, "hh:mm")
        If (IsDate(hora)) Then
            blanco colH, i
            'Almacena los datos de la hora
            niv = prs.Cells(i, colV).Value
            alm = prs.Cells(i, colV + 1).Value
            ver = prs.Cells(i, colV + 2).Value
            llu = prs.Cells(i, colV + 3).Value
            
            'Valida el valor de nivel
            If (niv <> "") Then
                If (IsNumeric(niv)) Then
                    comp(0) = ultNiv(clvPrs(0), hora)
                    comp(1) = comp(0) - varCom(0)
                    comp(2) = comp(0) + varCom(0)
                    If (niv >= comp(1) And niv <= comp(2)) Then
                        niv = Format(niv, "0.00")
                        blanco colV, i
                    Else
                        'Error de dato en calculo
                        rojo colV, i
                        erc = True
                    End If
                Else
                    'ERROR, El valor debe ser numérico
                    rojo colV, i
                    erf = True
                End If
            End If
            'Valida el valor de almacenamiento
            If (alm <> "") Then
                If (IsNumeric(alm)) Then
                    alm = Format(alm, "0.00")
                    blanco colV + 1, i
                Else
                    'ERROR, el valor debe ser numérico
                    rojo colV + 1, i
                    erf = True
                End If
            End If
            'Valida el valor de gasto
            If (ver <> "") Then
                If (IsNumeric(ver)) Then
                    ver = Format(ver, "0.00")
                    blanco colV + 2, i
                Else
                    'ERROR, el valor debe ser numérico
                    rojo colV + 2, i
                    erf = True
                End If
            End If
            'valida el valor de lluvia
            If (llu <> "") Then
                If Not (IsNumeric(llu)) Then
                    If (llu = "inap" Or llu = "INAP" Or llu = "Inap") Then
                        llu = 0.01
                        blanco colV + 3, i
                    Else
                        'ERROR, el valor solo puede ser numerico o cadena correspondiente a Inapreciable
                        rojo colV + 3, i
                        erf = True
                    End If
                ElseIf (CDbl(llu) >= 0) Then
                    If (CDbl(llu) <> 0.01) Then
                        llu = Format(llu, "0.0")
                        blanco colV + 3, i
                    Else
                        prs.Cells(i, colV + 3).Value = "Inap"
                        llu = 0.01
                    End If
                Else
                    'ERROR, el valor no puede ser menos a 0
                    rojo colV + 3, i
                    erf = True
                End If
            End If
            'Captura en SIH
            If Not erf Then
                'Conexión a la base de datos
                dbSIH.ConnectionString = "SIH"
                dbSIH.Open
                If niv <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(0) + "', '" + fecha + " " + hora + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If alm <> "" Then
                    query = "REPLACE INTO DTVolAlmac (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(0) + "', '" + fecha + " " + hora + "', '" + alm + "', '" + alm + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If ver <> "" Then
                    query = "REPLACE INTO DTVertedor (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(0) + "', '" + fecha + " " + hora + "', '" + ver + "', '" + ver + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If llu <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(0) + "', '" + fecha + " " + hora + "', '" + llu + "', '" + llu + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                'Fin de la conexión
                dbSIH.Close
            Else
                'ERROR, Encontró algún dato incorrecto
                erc = True
            End If
        Else
            'Dato hora es incorrecto
            rojo colH, i
            erc = True
        End If
    Next i
    
    'Guarda datos de Presa La Cangrejera
    colH = 9
    colV = 10
    For i = inicioFil To ultFil
        colV = 10
        erf = False
        hora = Format(prs.Cells(i, colH).Value, "hh:mm")
        If (IsDate(hora)) Then
            blanco colH, i
            'Datos Presa la Cangrejera
            niv = prs.Cells(i, colV).Value
            alm = prs.Cells(i, colV + 1).Value
            llu = prs.Cells(i, colV + 2).Value
            'Datos Plantas PB
            p1N = prs.Cells(i, colV + 3).Value
            p1L = prs.Cells(i, colV + 4).Value
            p2L = prs.Cells(i, colV + 5).Value
            p3N = prs.Cells(i, colV + 6).Value
            p3L = prs.Cells(i, colV + 7).Value
            
            'Valida el valor de nivel
            If (niv <> "") Then
                If (IsNumeric(niv)) Then
                   niv = Format(niv, "0.00")
                   blanco colV, i
                Else
                    'ERROR, el nivel solo puede ser numérico
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            'Valida el valor de almacenamiento
            If (alm <> "") Then
                If (IsNumeric(alm)) Then
                    alm = Format(alm, "0.000")
                    blanco colV, i
                Else
                    'ERROR, almacenamiento solo puede ser numérico
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            'valida el valor de lluvia
            If (llu <> "") Then
                If Not (IsNumeric(llu)) Then
                    If (llu = "inap" Or llu = "INAP" Or llu = "Inap") Then
                        llu = 0.01
                        blanco colV, i
                    Else
                        'Lluvia de tipo cadena solo puede tener valor correspondiente a Inapreciable
                        rojo colV, i
                        erf = True
                    End If
                ElseIf (CDbl(llu) >= 0) Then
                    If (CDbl(llu) <> 0.01) Then
                        llu = Format(llu, "0.0")
                        blanco colV, i
                    Else
                        prs.Cells(i, colV).Value = "Inap"
                        llu = 0.01
                    End If
                Else
                    'ERROR, lluvia no puede ser menor a 0
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            
            'Valida el valor de nivel PB1
            If (p1N <> "") Then
                If (IsNumeric(p1N)) Then
                   p1N = Format(p1N, "0.00")
                   blanco colV, i
                Else
                    'ERROR, Nivel solo puede ser numérico
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            'valida el valor de lluvia en PB1
            If (p1L <> "") Then
                If Not (IsNumeric(p1L)) Then
                    If (p1L = "inap" Or p1L = "INAP" Or p1L = "Inap") Then
                        p1L = 0.01
                        blanco colV, i
                    Else
                        'ERROR, Lluvia de tipo cadena solo puede ser correspondiente a Inapreciable
                        rojo colV, i
                        erf = True
                    End If
                ElseIf (CDbl(p1L) >= 0) Then
                    If (CDbl(p1L) <> 0.01) Then
                        p1L = Format(p1L, "0.0")
                        blanco colV, i
                    Else
                        prs.Cells(i, colV).Value = "Inap"
                        p1L = 0.01
                    End If
                Else
                    'ERROR, Lluvia no puede ser menor a 0
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            'valida el valor de lluvia en PB2
            If (p2L <> "") Then
                If Not (IsNumeric(p2L)) Then
                    If (p2L = "inap" Or p2L = "INAP" Or p2L = "Inap") Then
                        p2L = 0.01
                        blanco colV, i
                    Else
                        'ERROR, Lluvia de tipo cadena solo puede ser correspondiente a Inapreciable
                        rojo colV, i
                        erf = True
                    End If
                ElseIf (CDbl(p2L) >= 0) Then
                    If (CDbl(p2L) <> 0.01) Then
                        p2L = Format(p2L, "0.0")
                        blanco colV, i
                    Else
                        prs.Cells(i, colV).Value = "Inap"
                        p2L = 0.01
                    End If
                Else
                    'ERROR, Lluvia no puede ser menor a 0
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            'Valida el valor de nivel de PB1
            If (p3N <> "") Then
                If (IsNumeric(p3N)) Then
                   p3N = Format(p3N, "0.00")
                   blanco colV, i
                Else
                    'ERROR, Nivel solo puede ser numérico
                    rojo colV, i
                    erf = True
                End If
            End If
            colV = colV + 1
            'valida el valor de lluvia en PB2
            If (p3L <> "") Then
                If Not (IsNumeric(p3L)) Then
                    If (p3L = "inap" Or p3L = "INAP" Or p3L = "Inap") Then
                        p3L = 0.01
                        blanco colV, i
                    Else
                        'ERROR, Lluvia de tipo cadena solo puede ser correspondiente a Inapreciable
                        rojo colV, i
                        erf = True
                    End If
                ElseIf (CDbl(p3L) >= 0) Then
                    If (CDbl(p3L) <> 0.01) Then
                        p3L = Format(p3L, "0.0")
                        blanco colV, i
                    Else
                        prs.Cells(i, colV).Value = "Inap"
                        p3L = 0.01
                    End If
                Else
                    'ERROR, Lluvia no puede ser menor a 0
                    rojo colV, i
                    erf = True
                End If
            End If
            'Captura información en SIH
            If Not erf Then
                'Conexión a la base de datos
                dbSIH.ConnectionString = "SIH"
                dbSIH.Open
                'Información de Presa La Cangrejera
                If niv <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If alm <> "" Then
                    query = "REPLACE INTO DTVolAlmac (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + alm + "', '" + alm + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If llu <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + llu + "', '" + llu + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                'Información de plantas
                If p1N <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(2) + "', '" + fecha + " " + hora + "', '" + p1N + "', '" + p1N + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p1L <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(2) + "', '" + fecha + " " + hora + "', '" + p1L + "', '" + p1L + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p2L <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(3) + "', '" + fecha + " " + hora + "', '" + p2L + "', '" + p2L + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p3N <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(4) + "', '" + fecha + " " + hora + "', '" + p3N + "', '" + p3N + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p3L <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(4) + "', '" + fecha + " " + hora + "', '" + p3L + "', '" + p3L + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                'Fin de la conexión
                dbSIH.Close
            Else
                'ERROR, Encontro algún dato incorrecto
                erc = True
            End If
        Else
            'El formato de hora no es correcto
            rojo colH, i
            erc = True
        End If
    Next i
    
    If erc Then
        MsgBox "Se encontraron errores en la información.", vbCritical
    End If

End Sub

Sub desStd()
    'Variables
    Dim clvPrs(3) As String
    Dim diasDif As Integer
    Dim fch As String
    
    'Confirma asignacion de hoja a variable
    Set prs = Worksheets("Presas")
    'Número de dias de diferencia para obtener desviación estándar
    diasDif = -5
    
    If (fecha = "") Then
        'Asigna fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    'Fecha N días anterior
    fechaD = Format(DateAdd("d", diasDif, fecha), "yyyy/mm/dd 00:00")
    
    'Claves SIH para las presas
    clvPrs(0) = "CDOOX" 'Cerro de Oro
    clvPrs(1) = "LCAVC" 'La Cangrejera
    clvPrs(2) = "PCNVC" 'La Cangrejera PB1
    'PB2 unicamente captura datos de lluvia y no requiere calculo de desviación estandar
    clvPrs(3) = "PB3VC" 'La Cangrejera PB3
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtine la desviación estandar del valor nivel en Cerro de Oro
    query = "SELECT STD(valuee) as desEs FROM dtNivel WHERE station = '" & clvPrs(0) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(0) = "SN"
        Else
            varCom(0) = adoRs!desEs
        End If
    End If
    adoRs.Close
    'Obtiene la desviación estandar del valor almacenamiento en Cerro de Oro
    query = "SELECT STD(valuee)as desEs FROM DTVolAlmac WHERE station = '" & clvPrs(0) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(1) = "SN"
        Else
            varCom(1) = adoRs!desEs
        End If
    End If
    adoRs.Close
    'Obtiene la desviación estandar del valor gasto en Cerro de Oro
    query = "SELECT STD(valuee)as desEs FROM DTVertedor WHERE station = '" & clvPrs(0) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(2) = "SN"
        Else
            varCom(2) = adoRs!desEs
        End If
    End If
    adoRs.Close
    
    'Obtine la desviación estandar del valor nivel en La Cangrejera
    query = "SELECT STD(valuee)as desEs FROM dtNivel WHERE station = '" & clvPrs(1) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(3) = "SN"
        Else
            varCom(3) = adoRs!desEs
        End If
    End If
    adoRs.Close
    'Obtiene la desviación estandar del valor almacenamiento en La cangrejera
    query = "SELECT STD(valuee)as desEs FROM DTVolAlmac WHERE station = '" & clvPrs(1) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(4) = "SN"
        Else
            varCom(4) = adoRs!desEs
        End If
    End If
    adoRs.Close
    'Obtine la desviación estandar del nivel en La Cangrejera PB1
    query = "SELECT STD(valuee)as desEs FROM dtNivel WHERE station = '" & clvPrs(2) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(5) = "SN"
        Else
            varCom(5) = adoRs!desEs
        End If
    End If
    adoRs.Close
    'Obtine la desviación estandar del nivel en La Cangrejera PB2
    query = "SELECT STD(valuee)as desEs FROM dtNivel WHERE station = '" & clvPrs(3) & "' AND datee >= '" & fechaD & "' AND datee<= '" & fecha & " 23:59'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            varCom(6) = "SN"
        Else
            varCom(6) = adoRs!desEs
        End If
    End If
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close
    

End Sub

Private Function ultNiv(clvSIH As String, hora As String) As Double
    Dim diasD As Integer
    Dim hra As String
    'Confirma asignacion de hoja a variable
    Set prs = Worksheets("Presas")
    'Número de días para generar rango de fechas en la buqueda de valores
    diasD = -2
    If (fecha = "") Then
        'Asigna fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    'Fecha actual menos 1 minuto
    hra = fecha & " " & hora
    hra = Format(DateAdd("N", -1, hra), "yyyy/mm/dd hh:mm")
    'Fecha diferida a la actual
    fechaD = Format(DateAdd("d", diasD, fecha), "yyyy/mm/dd 00:00")
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Consulta el último valor de nivel capturado en Cerro de Oro
    query = "SELECT valuee AS val FROM dtNivel WHERE station = '" & clvSIH & "' AND datee >= '" & fechaD & "' AND datee <= '" & hra & "' ORDER BY Datee DESC LIMIT 1"
    'MsgBox query
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            ultNiv = Format(adoRs!val, "0.00")
        End If
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close

End Function
Private Function ultAlm(clvSIH As String, fecha As String) As Double
    Dim diasD As Integer
    
    'Confirma asignacion de hoja a variable
    Set prs = Worksheets("Presas")
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd " & hora)
    'Número de días para generar rango de fechas en la buqueda de valores
    diasD = -2
    'Fecha diferida a la actual
    fechaD = Format(DateAdd("d", diasD, fecha), "yyyy/mm/dd 00:00")
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Consulta el último valor de nivel capturado en Cerro de Oro
    query = "SELECT valuee AS val FROM DTVolAlmac WHERE station = '" & clvSIH & "' AND datee >= '" & fechaD & "' AND datee <= '" & fecha & "' ORDER BY Datee DESC LIMIT 1"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            ultAlm = Format(adoRs!val, "0.00")
        End If
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close
End Function
Private Function ultGas(clvSIH As String, fecha As String) As Double
    Dim diasD As Integer
    
    'Confirma asignacion de hoja a variable
    Set prs = Worksheets("Presas")
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd " & hora)
    'Número de días para generar rango de fechas en la buqueda de valores
    diasD = -2
    'Fecha diferida a la actual
    fechaD = Format(DateAdd("d", diasD, fecha), "yyyy/mm/dd 00:00")
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Consulta el último valor de nivel capturado en Cerro de Oro
    query = "SELECT valuee AS val FROM DTVertedor WHERE station = '" & clvSIH & "' AND datee >= '" & fechaD & "' AND datee <= '" & fecha & "' ORDER BY Datee DESC LIMIT 1"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            ultGas = Format(adoRs!val, "0.00")
        End If
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close

End Function

Public Function getFecha() As String
    getFecha = fecha
End Function

Public Function setFecha(f As String)
    fecha = f
End Function

Private Sub rojo(col As Integer, rows As Integer)
    Set prs = Worksheets("Presas")
    prs.Cells(rows, col).Interior.Color = vbRed
End Sub
Private Sub blanco(col As Integer, rows As Integer)
    Set prs = Worksheets("Presas")
    prs.Cells(rows, col).Interior.Color = xlNone
End Sub
