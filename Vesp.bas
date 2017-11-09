Attribute VB_Name = "Vesp"
'*********************************************************************
'   Sistema en apoyo a la captura de informacion hidroclimatologica
'          para la dirección técnica en OCGC, CONAGUA
'*********************************************************************
Option Explicit
'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String
'Otras variables
Private clvEst As String
Public fecha As String
Private ultFil As Integer
Private i As Integer

'Obtiene datos capturados a las 17 horas si existen
Sub getDatos()
    Dim bandera1 As Boolean
    Dim bandera2 As Boolean

    'Obtiene el número de la última fila
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd")
    'Limpia contenido
    Range("F9:H" & ultFil).ClearContents
    
    bandera1 = True
    bandera2 = True
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene los valores capturados a las 17 horas si existieran
    For i = 9 To ultFil
        'Almacena la clave de la estación
        clvEst = Range("B" & i).Value
        'Inicia bandera
        bandera1 = esEstacion(clvEst, i)

        If bandera1 Then
            'Consulta temperatura y escribe
            query = "SELECT valuee FROM dttempaire WHERE station = '" & clvEst & "' AND datee = '" & fecha & " 17:00'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("F" & i).CopyFromRecordset adoRs
                End If
            adoRs.Close
            'Consulta Lluvia y escribe
            query = "SELECT valuee FROM dtprecipitacio WHERE station = '" & clvEst & "' AND datee = '" & fecha & " 17:00'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("G" & i).Value = Range("K" & i).Value
                End If
            adoRs.Close
            'Consulta nivel y escribe
            query = "SELECT valuee FROM dtnivel WHERE station = '" & clvEst & "' AND datee = '" & fecha & " 17:00'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("H" & i).CopyFromRecordset adoRs
                End If
            adoRs.Close
        Else
            bandera2 = False
        End If
    Next i
    'Fin de la conexión
    dbSIH.Close
    'Mensaje en caso de haber error
    If (Not bandera2) Then
        MsgBox "Alguna(s) claves no son correctas", vbCritical, "ERROR"
    End If
End Sub

'Obtiene lluvias acumuladas de 8 a 17 horas
Sub acumuladas()
    'Variables hidrométricas
    Dim lluvAcu As String
    'Otras variables
    Dim bandera1 As Boolean
    Dim bandera2 As Boolean

    'Obtiene el número de la última fila
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd")
    'Selecciona contenido y lo limpia
    Range("K9:K" & ultFil).ClearContents
    'Bandera que indica un error en los datos
    bandera1 = True
    bandera2 = True
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    'Obtiene lluvias acumuladas
    For i = 9 To ultFil
        'Almacena la clave de la estación
        clvEst = Range("B" & i).Value
        'Inicia bandera
        bandera1 = esEstacion(clvEst, i)
        
        If (bandera1) Then
            query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & clvEst & "' and datee >= '" & fecha & " 08:00' and datee <= '" & fecha & " 17:00'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    'Cambia color de fuente de acuerdo a la situación (lluvia/sin lluvia)
                    If (adoRs!Acumulado > 0) Then
                        Range("K" & i).Font.Color = vbBlue
                    ElseIf (adoRs!Acumulado = 0) Then
                        Range("K" & i).Font.Color = RGB(198, 89, 17)
                    Else
                        Range("K" & i).Font.Color = vbBlack
                    End If
                    'Escribe lluvia acumulada
                    If (adoRs!Acumulado > 0 And adoRs!Acumulado < 0.1) Then
                        Range("K" & i).Formula = "Inap"
                    Else
                        Range("K" & i).Formula = Format(adoRs!Acumulado, "0.0")
                    End If
                End If
            adoRs.Close
        Else
            bandera2 = False
        End If
    Next i

    'Fin de la conexión
    dbSIH.Close
    'Mensaje en caso de haber error
    If (Not bandera2) Then
        MsgBox "Alguna(s) claves no son correctas", vbCritical, "ERROR"
    End If
End Sub

'Obtiene último nivel de escala capturado en el SIH del día
Sub ultNiv()
    Dim bandera1 As Boolean
    Dim bandera2 As Boolean

    'Obtiene el número de la última fila
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd")
    'Limpia contenido
    Range("L9:L" & ultFil).ClearContents
    
    bandera1 = True
    bandera2 = True
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene el último registro de nivel capturado
    For i = 9 To ultFil
        'Almacena la clave de la estación
        clvEst = Range("B" & i).Value
        'Inicia bandera
        bandera1 = esEstacion(clvEst, i)
        
        If (bandera1) Then
            'Consulta el último nivel de escala registrado al día
            query = "SELECT valuee AS val FROM dtNivel WHERE station = '" & clvEst & "' AND datee >= '" & fecha & " 00:00' AND datee <= '" & fecha & " 16:59' ORDER BY Datee DESC LIMIT 1"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("L" & i).Value = Format(adoRs!Val, "0.00")
                End If
            adoRs.Close
        Else
            bandera2 = False
        End If
    Next i
    
    'Fin de la conexión
    dbSIH.Close
    'Mensaje en caso de haber error
    If (Not bandera2) Then
        MsgBox "Alguna(s) claves no son correctas", vbCritical, "ERROR"
    End If
End Sub
'Obtiene la desviación estándar de los niveles de escala
Sub desviacionStd()
    Dim fechaDif As Date
    Dim diasDif As Integer
    Dim bandera1 As Boolean
    Dim bandera2 As Boolean

    'Obtiene el número de la última fila
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    'Número de días para obtener desviación estándar
    diasDif = 30
    'Obtiene la fecha actual
    fecha = Format(Now, "short date")
    'Fecha N días anterior
    fechaDif = DateDiff("d", diasDif, fecha)
    'Limpia contenido
    Range("M9:M" & ultFil).ClearContents
    
    bandera1 = True
    bandera2 = True
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene la desviación estándar del nivel de cada estacion
    For i = 9 To ultFil
        'Almacena la clave de la estación
        clvEst = Range("B" & i).Value
        'Inicia bandera
        bandera1 = esEstacion(clvEst, i)
        
        If (bandera1) Then
            'Consulta desviación estándar
            query = "SELECT STD(valuee)as desEs FROM dtNivel WHERE station = '" & clvEst & "' AND datee >= '" & Format(fechaDif, "yyyy/mm/dd hh:mm") & "' AND datee<= '" & Format(fecha, "yyyy/mm/dd") & " 17:00'"
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    If (adoRs!desEs = 0) Then
                        Range("M" & i).Value = ""
                    Else
                        Range("M" & i).CopyFromRecordset adoRs
                    End If
                End If
            adoRs.Close
        Else
            bandera2 = False
        End If
    Next i
    'Fin de la conexión
    dbSIH.Close
    'Mensaje en caso de haber error
    If (Not bandera2) Then
        MsgBox "Alguna(s) claves no son correctas", vbCritical, "ERROR"
    End If
End Sub

Sub capturar()
    'Variables hidrométricas
    Dim tmax As String
    Dim lluv As String
    Dim niv As String
    Dim lluvAcum As String
    Dim desA As Double
    Dim desB As Double
    'Otras variables
    Dim bandera1 As Boolean
    Dim bandera2 As Boolean

    'Obtiene el número de la última fila
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd") & " 17:00"
    
    'Banderas determina error en los datos
    bandera1 = True
    bandera2 = True
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Captura datos en SIH
    For i = 9 To ultFil
        'Obtiene datos hidrométricos
        clvEst = Range("B" & i).Value
        tmax = Range("F" & i).Value
        lluv = Range("G" & i).Value
        niv = Range("H" & i).Value
        lluvAcum = Range("K" & i).Value
        desA = Range("L" & i).Value + Range("M" & i).Value
        desB = Range("L" & i).Value - Range("M" & i).Value
        
        bandera1 = esEstacion(clvEst, i)
        
        'Valida valor temperatura máxima
        If (bandera1) Then
            If (tmax <> "") Then
                If Not (IsNumeric(tmax)) Then
                    rojo "F", CStr(i)
                    bandera1 = False
                    bandera2 = False
                ElseIf (tmax < 0 And tmax > 60) Then
                    rojo "F", CStr(i)
                    bandera1 = False
                    bandera2 = False
                Else
                    tmax = Format(tmax, "0.0")
                    blanco "F", CStr(i)
                End If
            End If
        Else
            bandera2 = False
        End If
        'Valida valor lluvia
        If bandera1 Then
            If (lluv <> "") Then
                If Not (IsNumeric(lluv)) Then
                    If (lluv = "inap" Or lluv = "INAP" Or lluv = "Inap") Then
                        lluv = 0.01
                        blanco "G", CStr(i)
                    Else
                        rojo "G", CStr(i)
                        bandera1 = False
                        bandera2 = False
                    End If
                ElseIf (lluv < 0) Then
                    rojo "G", CStr(i)
                    bandera1 = False
                    bandera2 = False
                ElseIf (lluv <> 0.01) Then
                    lluv = Format(lluv, "0.0")
                    blanco "G", CStr(i)
                End If
                
                If (lluvAcum = "Inap") Then
                    If (lluv = 0.01) Then
                        lluv = 0
                    End If
                ElseIf (lluv >= lluvAcum) Then
                    lluv = Format(lluv - lluvAcum, "0.0")
                Else
                    rojo "G", CStr(i)
                    bandera1 = False
                    bandera2 = False
                End If
            End If
        End If
        'Continúa validando valor de escala
        If bandera1 Then
            If (niv <> "") Then
                If Not (IsNumeric(niv)) Then
                    rojo "H", CStr(i)
                    bandera1 = False
                    bandera2 = False
                ElseIf (Range("L" & i).Value = "") Then
                    niv = Format(niv, "0.00")
                    blanco "H", CStr(i)
                ElseIf (niv >= desB And niv <= desA) Then
                    niv = Format(niv, "0.00")
                    blanco "H", CStr(i)
                Else
                    rojo "H", CStr(i)
                    bandera1 = False
                    bandera2 = False
                End If
            End If
        End If
        'Captura datos en SIH
        If (bandera1) Then
            'Captura temperatura máxima
            If (tmax <> "") Then
                query = "REPLACE INTO dttempaire (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvEst + "', '" + fecha + "', '" + tmax + "', '" + tmax + "', ' ', 'XL', ' ')"
                adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
            End If
            'Captura lluvia
            If (lluv <> "") Then
                query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvEst + "', '" + fecha + "', '" + lluv + "', '" + lluv + "', ' ', 'XL', ' ')"
                adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
            End If
            'Captura nivel
            If (niv <> "") Then
                query = "REPLACE INTO dtnivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvEst + "', '" + fecha + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
                adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
            End If
        End If
    Next i
    
    'Fin de la conexión
    dbSIH.Close
    'Mensaje en caso de haber error
    If bandera2 Then
        Range("K6").Value = "Última captura " & Format(Now, "dd/mmm/yyyy hh:mm") & " horas."
    Else
        MsgBox "Algunos datos son incorrectos", vbCritical, "ERROR"
    End If
End Sub
'Cambia color de fondo a la celda
Private Sub rojo(col As String, rows As String)
    Range(col & rows).Interior.Color = vbRed
End Sub
Private Sub blanco(col As String, rows As String)
    Range(col & rows).Interior.Color = xlNone
End Sub
'Valida si la clave de estación es correcta
Private Function esEstacion(est As String, lin As Integer) As Boolean
    If est = "" Then
        rojo "B", CStr(lin)
        esEstacion = False
    ElseIf (Len(est) <> 5) Then
        rojo "B", CStr(lin)
        esEstacion = False
    Else
        If (est = "TXPVC" Or est = "XOBVC" Or est = "VERVC" Or est = "ORZVC" Or est = "COTVC") Then
            Range("B" & lin).Interior.Color = RGB(255, 230, 153)
            esEstacion = True
        Else
            Range("B" & lin).Interior.Color = RGB(255, 242, 204)
            esEstacion = True
        End If
    End If
End Function
