Attribute VB_Name = "Vesp"
'Llena las celdas con los valores capturados a las 17 horas si existen
Sub getDatos()
'Variables para conexión a la Base de Datos
Dim dbSIH As New ADODB.Connection
Dim adoRs As New ADODB.Recordset
Dim qry As String
'Variables hidrométricas
Dim clvEs As String
'Otras variables
Dim fecha As String
Dim lastRow As Integer
Dim Flag1 As Boolean
Dim i As Integer

    'Obtiene el número de la última fila
    lastRow = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd")
    'Limpia contenido
    Range("F9:H" & lastRow).ClearContents
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene los valores capturados a las 17 horas si existieran
    For i = 9 To lastRow
        'Almacena la clave de la estación
        clvEs = Range("B" & i).Value
        'Inicia bandera
        Flag1 = esEstacion(clvEs, i)
        'Valida Clave de estación
        
        If Flag1 Then
            'Consulta temperatura y la escribe
            qry = "SELECT valuee FROM dttempaire WHERE station = '" & clvEs & "' AND datee = '" & fecha & " 17:00'"
            adoRs.Open qry, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("F" & i).CopyFromRecordset adoRs
                End If
            adoRs.Close
            'Consulta Lluvia y la escribe
            qry = "SELECT valuee FROM dtprecipitacio WHERE station = '" & clvEs & "' AND datee = '" & fecha & " 17:00'"
            adoRs.Open qry, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("G" & i).CopyFromRecordset adoRs
                End If
            adoRs.Close
            'Consulta nivel y la escribe
            qry = "SELECT valuee FROM dtnivel WHERE station = '" & clvEs & "' AND datee = '" & fecha & " 17:00'"
            adoRs.Open qry, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    Range("H" & i).CopyFromRecordset adoRs
                End If
            adoRs.Close
        End If
    Next i
    'Fin de la conexión
    dbSIH.Close
End Sub
'Obtiene las lluvias acumuladas de las 8 a 17 horas
Sub acumuladas()
'Variables para conexión a la Base de Datos
Dim dbSIH As New ADODB.Connection
Dim adoRs As New ADODB.Recordset
Dim qry As String
'Variables hidrométricas
Dim clvEs As String
Dim lluvAcu As String
'Otras variables
Dim fecha As String
Dim lastRow As Integer
Dim Flag1 As Boolean
Dim flag2 As Boolean

    'Obtiene el número de la última fila
    lastRow = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd")
    'Selecciona contenido y lo limpia
    Range("K9:K" & lastRow).ClearContents
    'Bandera que indica un error en los datos
    flag2 = True
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    'Obtiene lluvias acumuladas
    For i = 9 To lastRow
        'Almacena la clave de la estación
        clvEs = Range("B" & i).Value
        'Inicia bandera
        Flag1 = True
        'Valida Clave de estación
        If clvEs = "" Then
            rojo "B", CStr(i)
            Flag1 = False
            flag2 = False
        ElseIf (Len(clvEs) <> 5) Then
            rojo "B", CStr(i)
            Flag1 = False
            flag2 = False
        Else
            If (clvEs = "TXPVC" Or clvEs = "XOBVC" Or clvEs = "VERVC" Or clvEs = "ORZVC" Or clvEs = "COTVC") Then
                Range("B" & i).Interior.Color = RGB(255, 230, 153)
            Else
                Range("B" & i).Interior.Color = RGB(255, 242, 204)
            End If
        End If
        
        If (Flag1) Then
            qry = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & clvEs & "' and datee >= '" & fecha & " 08:00' and datee <= '" & fecha & " 17:00'"
            adoRs.Open qry, dbSIH, adOpenStatic, adLockReadOnly
                If Not adoRs.EOF Then
                    'Cambia color de fuente dependiendo con lluvia y sin lluvia
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
                Else
                    Range("K" & i).Formula = "Err"
                    Range("K" & i).Interior.Color = "Red"
                End If
            adoRs.Close
        End If
    Next i

    'Fin de la conexión
    dbSIH.Close
    'Mensaje en caso de haber error
    If (Not flag2) Then
        MsgBox "Alguna(s) claves no son correctas", vbCritical, "ERROR"
    End If

End Sub
'Obtiene el último nivel del escala capturado en el SIH del día
Sub ultNiv()
'Variables para conexión a la Base de Datos
Dim dbSIH As New ADODB.Connection
Dim adoRs As New ADODB.Recordset
Dim qry As String
'Variables hidrométricas
Dim clvEs As String
'Otras variables
Dim fecha As String
Dim lastRow As Integer

    'Obtiene el numero de la ultima fila
    lastRow = Range("B" & rows.Count).End(xlUp).Row
    'Obtiene la fecha actual
    fecha = Format(Now, "yyyy/mm/dd")
    'Limpia contenido
    Range("L9:L" & lastRow).ClearContents
    
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene el ultimo registro de nivel capturado
    For i = 9 To lastRow
        'Almacena la clave de la estación
        clvEs = Range("B" & i).Value
        'Consulta el último nivel de escala registrado al día
        qry = "SELECT valuee AS val FROM dtNivel WHERE station = '" & clvEs & "' AND datee >= '" & fecha & " 00:00' AND datee <= '" & fecha & " 16:59' ORDER BY Datee DESC LIMIT 1"
        adoRs.Open qry, dbSIH, adOpenStatic, adLockReadOnly
            If Not adoRs.EOF Then
                Range("L" & i).Value = Format(adoRs!Val, "0.00")
            End If
        adoRs.Close
    Next i
    'Fin de la conexión
    dbSIH.Close

End Sub
'Obtiene la desviación estándar de los niveles de escala
Sub desviacionStd()
'Variables para conexión a la Base de Datos
Dim dbSIH As New ADODB.Connection
Dim adoRs As New ADODB.Recordset
Dim qry As String
'Variables hidrométricas
Dim clvEs As String
'Otras variables
Dim fecha As Date
Dim fechaDif As Date
Dim lastRow As Integer
Dim diasDif As Integer

    'Obtiene el número de la última fila
    lastRow = Range("B" & rows.Count).End(xlUp).Row
    'Número de días para obtener desviación estándar
    diasDif = 30
    'Obtiene la fecha actual
    fecha = Format(Now, "short date")
    'Fecha N días anterior
    fechaDif = DateDiff("d", diasDif, fecha)
    'Limpia contenido
    Range("M9:M" & lastRow).ClearContents
    
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Obtiene la desviación estándar del nivel de cada estacion
    For i = 9 To lastRow
        'Almacena la clave de la estación
        clvEs = Range("B" & i).Value
        'Consulta desviación estándar
        qry = "SELECT STD(valuee)as desEs FROM dtNivel WHERE station = '" & clvEs & "' AND datee >= '" & Format(fechaDif, "yyyy/mm/dd hh:mm") & "' AND datee<= '" & Format(fecha, "yyyy/mm/dd") & " 17:00'"
        adoRs.Open qry, dbSIH, adOpenStatic, adLockReadOnly
            If Not adoRs.EOF Then
                If (adoRs!desEs = 0) Then
                    Range("M" & i).Value = ""
                Else
                    Range("M" & i).CopyFromRecordset adoRs
                End If
            End If
        adoRs.Close
    Next i
    'Fin de la conexión
    dbSIH.Close

End Sub
Sub capturar()

'Variables para conexión a la base de datos
Dim dbSIH As New ADODB.Connection
Dim adoRs As New ADODB.Recordset
Dim qry As String
'Variables hidrométricas
Dim clvEs As String
Dim tmax As String
Dim lluv As String
Dim niv As String
Dim lluvAcum As String
Dim desA As Double
Dim desB As Double
'Otras variables
Dim numRows As Integer
Dim Flag1 As Boolean
Dim flag2 As Boolean
Dim fecha As String
Dim lastRow As Integer

'Obtiene el número de la última fila
numRows = Range("B" & rows.Count).End(xlUp).Row
'Obtiene la fecha actual
fecha = Format(Now, "yyyy/mm/dd") & " 17:00"
'Bandera determina error en los datos
flag2 = True
'Obtiene el número de la última fila
lastRow = Range("B" & rows.Count).End(xlUp).Row

'Conexión a la base de datos
dbSIH.ConnectionString = "SIH"
dbSIH.Open

'Captura datos en SIH
For i = 9 To lastRow
    'Obtiene datos hidrométricos
    clvEs = Range("B" & i).Value
    tmax = Range("F" & i).Value
    lluv = Range("G" & i).Value
    niv = Range("H" & i).Value
    lluvAcum = Format(Val(Range("K" & i).Value), "0.0")
    desA = Range("L" & i).Value + Range("M" & i).Value
    desB = Range("L" & i).Value - Range("M" & i).Value
    'Iniciamos suponiendo que los datos son correctos
    Flag1 = True
    
    'Valida Clave de estación
    If clvEs = "" Then
        rojo "B", CStr(i)
        Flag1 = False
        flag2 = False
    ElseIf (Len(clvEs) <> 5) Then
        rojo "B", CStr(i)
        Flag1 = False
        flag2 = False
    Else
        If (clvEs = "TXPVC" Or clvEs = "XOBVC" Or clvEs = "VERVC" Or clvEs = "ORZVC" Or clvEs = "COTVC") Then
            Range("B" & i).Interior.Color = RGB(255, 230, 153)
        Else
            Range("B" & i).Interior.Color = RGB(255, 242, 204)
        End If
    End If
    
    'Valida valor temperatura máxima
    If Flag1 Then
        If (tmax <> "") Then
            If Not (IsNumeric(tmax)) Then
                rojo "F", CStr(i)
                Flag1 = False
                flag2 = False
            ElseIf (tmax < 0 And tmax > 60) Then
                rojo "F", CStr(i)
                Flag1 = False
                flag2 = False
            Else
                tmax = Format(tmax, "0.0")
                blanco "F", CStr(i)
            End If
        End If
    End If
    'Valida valor lluvia
    If Flag1 Then
        If (lluv <> "") Then
            If Not (IsNumeric(lluv)) Then
                If (lluv = "inap" Or lluv = "INAP" Or lluv = "Inap") Then
                    lluv = 0.01
                    blanco "G", CStr(i)
                Else
                    rojo "G", CStr(i)
                    Flag1 = False
                    flag2 = False
                End If
            ElseIf (lluv < 0) Then
                rojo "G", CStr(i)
                Flag1 = False
                flag2 = False
            ElseIf (lluv <> 0.01) Then
                lluv = Format(lluv, "0.0")
                blanco "G", CStr(i)
            End If
            
            If (lluv >= lluvAcum) Then
                lluv = lluv - lluvAcum
            Else
                rojo "G", CStr(i)
                Flag1 = False
                flag2 = False
            End If
        End If
    End If
    'Continúa validando valor de escala
    If Flag1 Then
        If (niv <> "") Then
            If Not (IsNumeric(niv)) Then
                rojo "H", CStr(i)
                Flag1 = False
                flag2 = False
            ElseIf (Range("L" & i).Value = "") Then
                niv = Format(niv, "0.00")
                blanco "H", CStr(i)
            ElseIf (niv >= desB And niv <= desA) Then
                niv = Format(niv, "0.00")
                blanco "H", CStr(i)
            Else
                rojo "H", CStr(i)
                Flag1 = False
                flag2 = False
            End If
        End If
    End If
    'Captura datos en SIH
    If (Flag1) Then
        'Captura temperatura máxima
        If (tmax <> "") Then
            qry = "REPLACE INTO dttempaire (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvEs + "', '" + fecha + "', '" + tmax + "', '" + tmax + "', ' ', 'XL', ' ')"
            adoRs.Open qry, dbSIH, adOpenDynamic, adLockOptimistic
        End If
        'Captura lluvia
        If (lluv <> "") Then
            qry = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvEs + "', '" + fecha + "', '" + lluv + "', '" + lluv + "', ' ', 'XL', ' ')"
            adoRs.Open qry, dbSIH, adOpenDynamic, adLockOptimistic
        End If
        'Captura nivel
        If (niv <> "" And Flag1) Then
            qry = "REPLACE INTO dtnivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvEs + "', '" + fecha + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
            adoRs.Open qry, dbSIH, adOpenDynamic, adLockOptimistic
        End If
    End If
Next i

'Fin de la conexión
dbSIH.Close

If flag2 Then
    MsgBox "Captura de datos en SIH terminada", vbOKOnly, "SIH"
Else
    MsgBox "Algunos datos no son correctos", vbCritical, "ERROR"
End If

End Sub

Private Sub rojo(col As String, rows As String)
    Range(col & rows).Interior.Color = vbRed
End Sub
Private Sub blanco(col As String, rows As String)
    Range(col & rows).Interior.Color = xlNone
End Sub

Private Sub err()
    MsgBox "Se encontraron algunos ERRORES", vbCritical, "ERROR"
End Sub

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
