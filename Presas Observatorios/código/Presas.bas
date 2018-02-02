Attribute VB_Name = "Presas"
'*********************************************************************
'           Sistema en apoyo a la captura de Lluvia para los
'  Observatorio Meteorológicos de dirección técnica en OCGC, CONAGUA
'
'                               PRESAS
'
'*********************************************************************
Option Explicit

'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String

'Otras variables
Private fecha As String
Private fechaD As String
Private ultFil As Integer
Private inicioFil As Integer
Private prs As Excel.Worksheet

'Desviacion estandar de las variables
Private desv(7) As Double
'Ultimos niveles
Private ultNiv As Double

Sub inicio()
    Set prs = Worksheets("Presas")
    prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    fecha = Format(Now, "yyyy/mm/dd")
    obtenerDatos
End Sub

Sub desviacionStd()
    Dim clvPrs(5) As String
    Dim fechaDif As Date
    Dim mesDif As Integer
    Dim bandera1 As Boolean
    Dim bandera2 As Boolean

    Set prs = Worksheets("Presas")
    'Obtiene el número de la última fila
    ultFil = prs.Range("B" & rows.Count).End(xlUp).Row
    'Número de meses de diferencia para obtener desviación estándar
    mesDif = -1
    'Obtiene la fecha actual
    fecha = Format(Now, "short date")
    'Fecha N días anterior
    fechaDif = DateAdd("m", mesDif, fecha)
    
    'bandera1 = True
    'bandera2 = True
    
    clvPrs(1) = "CDOOX"
    clvPrs(2) = "LCAVC"
    clvPrs(3) = "PCNVC"
    clvPrs(4) = "CB2VC"
    clvPrs(5) = "PB3VC"
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    query = "SELECT STD(valuee)as desEs FROM dtNivel WHERE station = '" & clvPrs(1) & "' AND datee >= '" & Format(fechaDif, "yyyy/mm/dd hh:mm") & "' AND datee<= '" & Format(fecha, "yyyy/mm/dd hh:mm") & "'"
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If (adoRs!desEs = 0) Then
            sNiv = ""
        Else
            sNiv = adoRs!desEs
        End If
    End If
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close

End Sub


Sub obtenerDatos()
    Dim clvPrs(5) As String
    Dim hora As String
    Dim colV As Integer
    Dim colH As Integer
    Dim i As Integer
    
    clvPrs(1) = "CDOOX"
    clvPrs(2) = "LCAVC"
    clvPrs(3) = "PCNVC"
    clvPrs(4) = "CB2VC"
    clvPrs(5) = "PB3VC"
    
    inicioFil = 12
    
    Set prs = Worksheets("Presas")
    
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
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Gasto de Presa Cerro de oro
            query = "SELECT valuee FROM DTVertedor WHERE station = '" & clvPrs(1) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
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
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(2) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Almacenamiento de Presa Cerro de oro
            query = "SELECT valuee FROM DTVolAlmac WHERE station = '" & clvPrs(2) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.000")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta Lluvia de Presa Cerro de oro
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
            'Consulta nivel de PB1
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(3) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta lluvia de PB1
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
            'Consulta lluvia de PB2
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
            colV = colV + 1
            'Consulta nivel de PB3
            query = "SELECT valuee FROM DTNivel WHERE station = '" & clvPrs(5) & "' AND datee = '" & fecha & " " & hora & "'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    prs.Cells(i, colV).Value = Format(adoRs!valuee, "0.00")
                End If
            adoRs.Close
            colV = colV + 1
            'Consulta lluvia de PB3
            query = "SELECT valuee FROM DTPrecipitacio WHERE station = '" & clvPrs(5) & "' AND datee = '" & fecha & " " & hora & "'"
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
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub capturarDatos()
    Dim niv As String
    Dim alm As String
    Dim ver As String
    Dim llu As String
    
    Dim p1N As String
    Dim p1L As String
    Dim p2L As String
    Dim p3N As String
    Dim p3L As String
    
    Dim clvPrs(5) As String
    Dim hora As String
    Dim colH As Integer
    Dim colV As Integer
    Dim i As Integer
    Dim ers As Boolean
    Dim erc As Boolean
    Dim prs As Excel.Worksheet
    
    Set prs = Worksheets("Presas")
    
    clvPrs(1) = "CDOOX"
    clvPrs(2) = "LCAVC"
    clvPrs(3) = "PCNVC"
    clvPrs(4) = "CB2VC"
    clvPrs(5) = "PB3VC"
    
    inicioFil = 12
    
    'Obtiene el número de la última fila
    ultFil = prs.Range("B" & rows.Count).End(xlUp).Row
    
    ers = True
    erc = True
    
    If (fecha = "") Then
        'Asigna la fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        prs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        prs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    colH = 2
    colV = 4
    For i = inicioFil To ultFil
        ers = True
        hora = Format(prs.Cells(i, colH).Value, "hh:mm")
        If (IsDate(hora)) Then
            blanco colH, i
            niv = prs.Cells(i, colV).Value
            alm = prs.Cells(i, colV + 1).Value
            ver = prs.Cells(i, colV + 2).Value
            llu = prs.Cells(i, colV + 3).Value
            
            'Valida el valor de nivel
            If (niv <> "") Then
                If (IsNumeric(niv)) Then
                   niv = Format(niv, "0.00")
                   blanco colV, i
                Else
                    rojo colV, i
                    ers = False
                End If
            End If
            'Valida el valor de almacenamiento
            If (alm <> "") Then
                If (IsNumeric(alm)) Then
                    alm = Format(alm, "0.00")
                    blanco colV + 1, i
                Else
                    rojo colV + 1, i
                    ers = False
                End If
            End If
            'Valida el valor de gasto
            If (ver <> "") Then
                If (IsNumeric(ver)) Then
                    ver = Format(ver, "0.00")
                    blanco colV + 2, i
                Else
                    rojo colV + 2, i
                    ers = False
                End If
            End If
            'valida el valor de lluvia
            If (llu <> "") Then
                If Not (IsNumeric(llu)) Then
                    If (llu = "inap" Or llu = "INAP" Or llu = "Inap") Then
                        llu = 0.01
                        blanco colV + 3, i
                    Else
                        rojo colV + 3, i
                        ers = False
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
                    rojo colV + 3, i
                    ers = False
                End If
            End If
            
            If ers Then
                If niv <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If alm <> "" Then
                    query = "REPLACE INTO DTVolAlmac (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + alm + "', '" + alm + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If ver <> "" Then
                    query = "REPLACE INTO DTVertedor (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + ver + "', '" + ver + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If llu <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(1) + "', '" + fecha + " " + hora + "', '" + llu + "', '" + llu + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
            Else
                erc = False
            End If
        Else
            rojo colH, i
        End If
    Next i
    
    'Guarda datos de Presa La Cangrejera
    colH = 9
    colV = 10
    For i = inicioFil To ultFil
        colV = 10
        ers = True
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
                    rojo colV, i
                    ers = False
                End If
            End If
            colV = colV + 1
            'Valida el valor de almacenamiento
            If (alm <> "") Then
                If (IsNumeric(alm)) Then
                    alm = Format(alm, "0.000")
                    blanco colV, i
                Else
                    rojo colV, i
                    ers = False
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
                        rojo colV, i
                        ers = False
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
                    rojo colV, i
                    ers = False
                End If
            End If
            colV = colV + 1
            
            
            'Valida el valor de nivel PB1
            If (p1N <> "") Then
                If (IsNumeric(p1N)) Then
                   p1N = Format(p1N, "0.00")
                   blanco colV, i
                Else
                    rojo colV, i
                    ers = False
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
                        rojo colV, i
                        ers = False
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
                    rojo colV, i
                    ers = False
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
                        rojo colV, i
                        ers = False
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
                    rojo colV, i
                    ers = False
                End If
            End If
            colV = colV + 1
            'Valida el valor de nivel de PB1
            If (p3N <> "") Then
                If (IsNumeric(p3N)) Then
                   p3N = Format(p3N, "0.00")
                   blanco colV, i
                Else
                    rojo colV, i
                    ers = False
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
                        rojo colV, i
                        ers = False
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
                    rojo colV, i
                    ers = False
                End If
            End If
            If ers Then
                If niv <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(2) + "', '" + fecha + " " + hora + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If alm <> "" Then
                    query = "REPLACE INTO DTVolAlmac (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(2) + "', '" + fecha + " " + hora + "', '" + alm + "', '" + alm + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If llu <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(2) + "', '" + fecha + " " + hora + "', '" + llu + "', '" + llu + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                
                
                If p1N <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(3) + "', '" + fecha + " " + hora + "', '" + p1N + "', '" + p1N + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p1L <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(3) + "', '" + fecha + " " + hora + "', '" + p1L + "', '" + p1L + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p2L <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(4) + "', '" + fecha + " " + hora + "', '" + p2L + "', '" + p2L + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p3N <> "" Then
                    query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(5) + "', '" + fecha + " " + hora + "', '" + p3N + "', '" + p3N + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                If p3L <> "" Then
                    query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvPrs(5) + "', '" + fecha + " " + hora + "', '" + p3L + "', '" + p3L + "', ' ', 'XL', ' ')"
                    adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                End If
                
            Else
                erc = False
            End If
        Else
            rojo colH, i
        End If
    Next i
    
    'Fin de la conexión
    dbSIH.Close

End Sub

Private Sub getDesviciones()
    

End Sub

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
