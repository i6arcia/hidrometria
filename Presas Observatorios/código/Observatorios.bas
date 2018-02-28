Attribute VB_Name = "Observatorios"
'*********************************************************************
'           Sistema en apoyo a la captura de Lluvia para los
'  Observatorio Meteorológicos de Dirección Técnica en OCGC, CONAGUA
'
'                              OBSERVATORIOS
'
'*********************************************************************
Option Explicit

'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String

'variable hoja de excel
Private obs As Excel.Worksheet

'Variables globales
Public iniFil As Integer
Public ultFil As Integer
Private fecha As String
Public x As Integer
Public y As Integer
Public clvObs() As String

'Control de modificaciones
Public edit() As Integer
Public flagEdit As Boolean


'Otras variables
Private fechaD As String
Private flagParametros As Boolean


Sub inicio()
    If Not flagEdit Then
        flagEdit = True
        'Asigna a la variable la hoja de calculo
        Set obs = Worksheets("Observatorios")
        'Número de fila que inicia con información
        iniFil = 11
            'Asigna variable fecha
            If fecha = "" Then
                'Asigna fecha actual
                fecha = Format(Now, "yyyy/mm/dd")
                obs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
                obs.Range("E7").Interior.Color = RGB(221, 235, 247)
            End If
        'Valor correspondiente al numero de observatorios
        If x < 1 Then
            x = 6
            ReDim clvObs(x - 1)
            'Claves SIH para los observatorios
            clvObs(0) = "TXPVC" 'Tuxpan
            clvObs(1) = "XOBVC" 'Xalapa
            clvObs(2) = "ORZVC" 'Orizaba
            clvObs(3) = "VERVC" 'Veracruz
            clvObs(4) = "COTVC" 'Coatzacoalcos
            clvObs(5) = "RALVC" 'Radar Alvarado
        End If
        'contar filas
        'Obtiene ultima fila
        contarFilas
        flagEdit = False
    End If
End Sub

Sub actualizar()
    'Variables
    Dim i As Integer
    Dim val As Integer
    Dim j As Integer
    
    If x > 0 Then
        flagEdit = True
        For i = 0 To x - 1
            getInfoSIH (clvObs(i))
        Next i
        'Termina la actualización de la hoja
        flagEdit = False
    Else
        inicio
    End If
    
End Sub

Sub getInfoSIH(estacion As String)
    'Variables
    Dim hora As String
    Dim col As Integer
    Dim i As Integer
    
    Select Case estacion
        Case "TXPVC" 'Tuxpan
            col = 1
        Case "XOBVC" 'Xalapa
            col = 2
        Case "ORZVC" 'Orizaba
            col = 3
        Case "VERVC" 'Veracruz
            col = 4
        Case "COTVC" 'Coatzacoalcos
            col = 5
        Case "RALVC" 'Radar Alvarado
            col = 6
    End Select
    
    If x > 0 Then
        'Limpia contenido
        obs.Range(obs.Cells(iniFil, col * 3), obs.Cells(ultFil, col * 3)).ClearContents
        
        'Establece valores Control de modificaciones como vacios
        For i = 0 To y
            edit(col - 1, i) = 0
        Next i
        
        'Conexión
        dbSIH.ConnectionString = "SIH"
        dbSIH.Open
        
        For i = iniFil To ultFil
            hora = Format(obs.Cells(i, (col * 3) - 1).Value, "hh:mm")
            If (IsDate(hora)) Then
                    blanco (col * 3) - 1, i, obs
                    If (hora = "07:00") Then
                        'Obtiene la lluvia acumulada de 8 am del dia anterior a 7 am del día actual (Dato calculado que requiere comprovación)
                        fechaD = Format(DateAdd("d", -1, fecha), "yyyy/mm/dd")
                        query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & estacion & "' and datee >= '" & fechaD & " 08:00' and datee <= '" & fecha & " 07:00'"
                        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                            If Not adoRs.EOF Then
                                'Escribe lluvia acumulada
                                If (adoRs!Acumulado > 0 And adoRs!Acumulado <= 0.1) Then
                                    obs.Cells(i, col * 3).Formula = "Inap"
                                Else
                                    obs.Cells(i, col * 3).Formula = Format(adoRs!Acumulado, "0.0")
                                End If
                                'Control de modificacion (No editable)
                                edit(col - 1, i - iniFil) = 2
                            Else
                                'Control de modificacion (Vacio)
                                edit(col - 1, i - iniFil) = 0
                            End If
                        adoRs.Close
                    ElseIf (hora = "17:00") Then
                        'Obtiene lluvia acumulada de 8 a 17 hrs.
                        query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & estacion & "' and datee >= '" & fecha & " 08:00' and datee <= '" & fecha & " 17:00'"
                        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                            If Not adoRs.EOF Then
                                'Escribe lluvia acumulada
                                If (adoRs!Acumulado > 0 And adoRs!Acumulado <= 0.1) Then
                                    obs.Cells(i, col * 3).Formula = "Inap"
                                Else
                                    obs.Cells(i, col * 3).Formula = Format(adoRs!Acumulado, "0.0")
                                End If
                                'Control de modificacion (No editable)
                                edit(col - 1, i - iniFil) = 2
                            Else
                                'Control de modificacion (Vacio)
                                edit(col - 1, i - iniFil) = 0
                            End If
                        adoRs.Close
                    Else
                        'Consulta lluvia en la hora especifica y la escribe
                        query = "SELECT valuee FROM dtprecipitacio WHERE station = '" & estacion & "' AND datee = '" & fecha & " " & hora & "'"
                        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                            If Not adoRs.EOF Then
                                If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                                    obs.Cells(i, col * 3).Value = "Inap"
                                Else
                                    obs.Cells(i, col * 3).Value = adoRs!valuee
                                End If
                                obs.Cells(i, col * 3).Interior.Color = xlNone
                                'Control de modificacion (Lleno)
                                edit(col - 1, i - iniFil) = 1
                            Else
                                obs.Cells(i, col * 3).Interior.Color = xlNone
                                'Control de modificacion (Vacio)
                                edit(col - 1, i - iniFil) = 0
                            End If
                        adoRs.Close
                    End If
                Else
                    'El formato de hora es incorrecto y cambia el color de relleno
                    rojo (col * 3) - 1, i, obs
                End If
        Next i
        
        'Fin de la conexión
        dbSIH.Close
    Else
        inicio
    End If

End Sub

Sub capturarDatos()
    'Variables
    Dim lluvia As String
    Dim clvObs(6) As String
    Dim hora As String
    Dim col As Integer
    Dim i As Integer
    Dim j As Integer
    'Variables para el control de errores
    Dim erf As Boolean 'Error en formato
    Dim erc As Boolean 'Error en calculo
    
    
    If ultFil > 0 Then
    Else
        inicio
    End If
    
    'Claves SIH para cada observatorio
    clvObs(1) = "TXPVC" 'Tuxpan
    clvObs(2) = "XOBVC" 'Xalapa
    clvObs(3) = "ORZVC" 'Orizaba
    clvObs(4) = "VERVC" 'Veracruz
    clvObs(5) = "COTVC" 'Coatzacoalcos
    clvObs(6) = "RALVC" 'Radar Alvarado
    'Fila donde inicia captura de información
    iniFil = 11
    'Última fila con información
    ultFil = obs.Range("B" & rows.Count).End(xlUp).Row
    'Variables para control de errores
    erf = False
    erc = False
    
    If (fecha = "") Then
        'Asigna la fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        obs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        obs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    'Captura información en SIH
    For i = 1 To 6
        col = i * 3
        For j = iniFil To ultFil
            'Re inicia variable Error en sitaxis
            erf = False
            hora = Format(obs.Cells(j, col - 1).Value, "hh:mm")
            'Valida la hora
            If (IsDate(hora)) Then
                'Excluye 7 y 17 horas. (Datos calculados)
                If (hora <> "17:00" And hora <> "07:00") Then
                    blanco col - 1, j, obs
                    
                    'Apartado para validar el dato lluvia
                    lluvia = obs.Cells(j, col).Value
                    If (lluvia <> "") Then
                        If Not (IsNumeric(lluvia)) Then
                            If (lluvia = "inap" Or lluvia = "INAP" Or lluvia = "Inap") Then
                                lluvia = 0.01
                                blanco col, j, obs
                            Else
                                'Dato lluvia, solo puede ser numérico o cadena"Inap"
                                rojo col, j, obs
                                erf = True
                            End If
                        ElseIf (CDbl(lluvia) >= 0) Then
                            If (CDbl(lluvia) <> 0.01) Then
                                lluvia = Format(lluvia, "0.0")
                                blanco col, j, obs
                            Else
                                'Asigna celda como lluvia inapreciable
                                obs.Cells(j, col).Value = "Inap"
                                lluvia = 0.01
                            End If
                        Else
                            'Dato lluvia no puede ser menor a 0
                            rojo col, j, obs
                            erf = True
                        End If
                        'Captura dato lluvia en SIH
                        If Not erf Then
                            query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvObs(i) + "', '" + fecha + " " + hora + "', '" + lluvia + "', '" + lluvia + "', ' ', 'XL', ' ')"
                            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                        Else
                            'Bandera informe de Error general
                            erc = True
                        End If
                    End If
                End If
            Else
                'Bandera informe de Error general
                rojo col - 1, j, obs
                erc = True
            End If
        Next j
    Next i
    'Fin de la conexión
    dbSIH.Close
    'Mesaje informando Error
    If erc Then
        MsgBox "Se encontraron errores en la información.", vbCritical
    End If
End Sub

'***************************
'Procedimientos herramientas
'***************************
Sub contarFilas()
    Dim val As Integer, i As Integer
    
    If x > 0 Then
        'Obtiene el valor de la ultima fila
        For i = 1 To x
            val = obs.Cells(rows.Count, (i * 3) - 1).End(xlUp).Row
            If ultFil < val Then
                ultFil = val
            End If
        Next i
        
        If y <> (ultFil - iniFil) Then
            y = ultFil - iniFil
            ReDim edit(x - 1, y)
            actualizar
        End If
    Else
        inicio
    End If
End Sub
Sub printM()
    Dim matriz As String
    Dim i As Integer
    Dim j As Integer
    
    If x > 0 Then
        For i = 0 To y
            For j = 0 To x - 1
                matriz = matriz & edit(j, i) & " "
            Next j
            matriz = matriz & vbCrLf
        Next i
        MsgBox "La matriz esta compuesta por los siguiente valores:" & vbCrLf & matriz
    Else
        inicio
        printM
    End If
End Sub


'Procedimientos para cambiar relleno de celdas en caso de error
Private Sub rojo(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.Color = vbRed
End Sub
Private Sub blanco(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.Color = xlNone
End Sub

'Manejo publico de variables
Public Function setFecha(fec As String)
    fecha = fec
End Function
Public Function getFecha() As String
    getFecha = fecha
End Function
