Attribute VB_Name = "Observatorios"
'*********************************************************************
'           Sistema en apoyo a la captura de Lluvia para los
'  Observatorio Meteorológicos de dirección técnica en OCGC, CONAGUA
'
'       OBSERVATORIOS
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
Private obs As Excel.Worksheet

Sub inicio()
    Set obs = Worksheets("Observatorios")
    obs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    obs.Range("E7").Interior.Color = RGB(221, 235, 247)
    fecha = Format(Now, "yyyy/mm/dd")
    obtenerDatos
End Sub

Sub obtenerDatos()
    Dim clvObs(6) As String
    Dim hora As String
    Dim col As Integer
    Dim i As Integer
    Dim j As Integer
    
    clvObs(1) = "TXPVC"
    clvObs(2) = "XOBVC"
    clvObs(3) = "ORZVC"
    clvObs(4) = "VERVC"
    clvObs(5) = "COTVC"
    clvObs(6) = "RALVC"
    
    inicioFil = 11
    
    Set obs = Worksheets("Observatorios")
    
    ultFil = obs.Range("B" & rows.Count).End(xlUp).Row
    
    If (fecha = "") Then
        'Asigna fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        obs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        obs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    'Limpia contenido
    obs.Range("C" & inicioFil & ":C" & ultFil).ClearContents
    obs.Range("F" & inicioFil & ":F" & ultFil).ClearContents
    obs.Range("I" & inicioFil & ":I" & ultFil).ClearContents
    obs.Range("L" & inicioFil & ":L" & ultFil).ClearContents
    obs.Range("O" & inicioFil & ":O" & ultFil).ClearContents
    obs.Range("R" & inicioFil & ":R" & ultFil).ClearContents
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    'Obtiene los datos
    For i = 1 To 6
        col = i * 3
        For j = inicioFil To ultFil
            hora = Format(obs.Cells(j, col - 1).Value, "hh:mm")
            If (IsDate(hora)) Then
                blanco col - 1, j, obs
                If (hora = "07:00") Then
                    'Obtiene la lluvia acumulada de 8 am del dia anterior a 7 am del día actual
                    fechaD = Format(DateAdd("d", -1, fecha), "yyyy/mm/dd")
                    query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & clvObs(i) & "' and datee >= '" & fechaD & " 08:00' and datee <= '" & fecha & " 07:00'"
                    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                        If Not adoRs.EOF Then
                            'Escribe lluvia acumulada
                            If (adoRs!Acumulado > 0 And adoRs!Acumulado <= 0.1) Then
                                obs.Cells(j, col).Formula = "Inap"
                            Else
                                obs.Cells(j, col).Formula = Format(adoRs!Acumulado, "0.0")
                            End If
                        End If
                    adoRs.Close
                ElseIf (hora = "17:00") Then
                    query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & clvObs(i) & "' and datee >= '" & fecha & " 08:00' and datee <= '" & fecha & " 17:00'"
                    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                        If Not adoRs.EOF Then
                            'Escribe lluvia acumulada
                            If (adoRs!Acumulado > 0 And adoRs!Acumulado <= 0.1) Then
                                obs.Cells(j, col).Formula = "Inap"
                            Else
                                obs.Cells(j, col).Formula = Format(adoRs!Acumulado, "0.0")
                            End If
                        End If
                    adoRs.Close
                Else
                    'Consulta Lluvia y escribe
                    query = "SELECT valuee FROM dtprecipitacio WHERE station = '" & clvObs(i) & "' AND datee = '" & fecha & " " & hora & "'"
                    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                        If Not adoRs.EOF Then
                            If (adoRs!valuee > 0 And adoRs!valuee <= 0.1) Then
                                obs.Cells(j, col).Value = "Inap"
                            Else
                                obs.Cells(j, col).Value = adoRs!valuee
                            End If
                        End If
                    adoRs.Close
                End If
            Else
                rojo col - 1, j, obs
            End If
        Next j
    Next i
    
    'Fin de la conexión
    dbSIH.Close
    
End Sub

Sub capturarDatos()
    Dim lluvia As String
    Dim clvObs(6) As String
    Dim hora As String
    Dim col As Integer
    Dim i As Integer
    Dim j As Integer
    Dim ers As Boolean
    Dim erc As Boolean
    Dim obs As Excel.Worksheet
    
    Set obs = Worksheets("Observatorios")
    'MsgBox "El valor es " & obs.Range("C13").Value
    
    clvObs(1) = "TXPVC"
    clvObs(2) = "XOBVC"
    clvObs(3) = "ORZVC"
    clvObs(4) = "VERVC"
    clvObs(5) = "COTVC"
    clvObs(6) = "RALVC"
    
    inicioFil = 11
    
    'Obtiene el número de la última fila
    ultFil = obs.Range("B" & rows.Count).End(xlUp).Row
    
    ers = True
    erc = True
    
    If (fecha = "") Then
        'Asigna la fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
        obs.Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
        obs.Range("E7").Interior.Color = RGB(221, 235, 247)
    End If
    
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    For i = 1 To 6
        col = i * 3
        For j = inicioFil To ultFil
            ers = True
            hora = Format(obs.Cells(j, col - 1).Value, "hh:mm")
            'Valida la hora
            If (IsDate(hora)) Then
                If (hora <> "17:00" And hora <> "07:00") Then
                    blanco col - 1, j, obs
                    'Valida el dato lluvia
                    lluvia = obs.Cells(j, col).Value
                    If (lluvia <> "") Then
                        If Not (IsNumeric(lluvia)) Then
                            If (lluvia = "inap" Or lluvia = "INAP" Or lluvia = "Inap") Then
                                lluvia = 0.01
                                blanco col, j, obs
                            Else
                                rojo col, j, obs
                                ers = False
                            End If
                        ElseIf (CDbl(lluvia) >= 0) Then
                            If (CDbl(lluvia) <> 0.01) Then
                                lluvia = Format(lluvia, "0.0")
                                blanco col, j, obs
                            Else
                                obs.Cells(j, col).Value = "Inap"
                                lluvia = 0.01
                            End If
                        Else
                            rojo col, j, obs
                            ers = False
                        End If
                        
                        If ers Then
                            query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvObs(i) + "', '" + fecha + " " + hora + "', '" + lluvia + "', '" + lluvia + "', ' ', 'XL', ' ')"
                            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                        Else
                            erc = False
                        End If
                        
                    End If
                End If
            Else
                rojo col - 1, j, obs
                erc = False
            End If
        Next j
    Next i
    
    'Fin de la conexión
    dbSIH.Close

End Sub
Public Function setFecha(fec As String)
    fecha = fec
End Function
Public Function getFecha() As String
    getFecha = fecha
End Function

Private Sub rojo(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.Color = vbRed
End Sub
Private Sub blanco(col As Integer, rows As Integer, ws As Excel.Worksheet)
    ws.Cells(rows, col).Interior.Color = xlNone
End Sub

