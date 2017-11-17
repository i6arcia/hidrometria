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
    
    
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    
    If (fecha = "") Then
        'Obtiene la fecha actual
        fecha = Format(Now, "yyyy/mm/dd")
    End If
    
    'Limpia contenido
    Range("C" & inicioFil & ":C" & ultFil).ClearContents
    Range("F" & inicioFil & ":F" & ultFil).ClearContents
    Range("I" & inicioFil & ":I" & ultFil).ClearContents
    Range("L" & inicioFil & ":L" & ultFil).ClearContents
    Range("O" & inicioFil & ":O" & ultFil).ClearContents
    Range("R" & inicioFil & ":R" & ultFil).ClearContents
    
    'Conexión
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    'Obtiene los datos
    For i = 1 To 6
        col = i * 3
        For j = inicioFil To ultFil
            hora = Format(Cells(j, col - 1).Value, "hh:mm")
            If (IsDate(hora)) Then
                If (hora = "07:00") Then
                    fechaD = Format(DateDiff("d", 1, fecha), "yyyy/mm/dd")
                    query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & clvObs(i) & "' and datee >= '" & fechaD & " 08:00' and datee <= '" & fecha & " 07:00'"
                    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                        If Not adoRs.EOF Then
                            'Escribe lluvia acumulada
                            If (adoRs!Acumulado > 0 And adoRs!Acumulado < 0.1) Then
                                Cells(j, col).Formula = "Inap"
                            Else
                                Cells(j, col).Formula = Format(adoRs!Acumulado, "0.0")
                            End If
                        End If
                    adoRs.Close
                ElseIf (hora = "17:00") Then
                    query = "Select sum(valuee) as Acumulado from dtPrecipitacio where station = '" & clvObs(i) & "' and datee >= '" & fecha & " 08:00' and datee <= '" & fecha & " 17:00'"
                    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                        If Not adoRs.EOF Then
                            'Escribe lluvia acumulada
                            If (adoRs!Acumulado > 0 And adoRs!Acumulado < 0.1) Then
                                Cells(j, col).Formula = "Inap"
                            Else
                                Cells(j, col).Formula = Format(adoRs!Acumulado, "0.0")
                            End If
                        End If
                    adoRs.Close
                Else
                    'Consulta Lluvia y escribe
                    query = "SELECT valuee FROM dtprecipitacio WHERE station = '" & clvObs(i) & "' AND datee = '" & fecha & " " & hora & "'"
                    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
                        If Not adoRs.EOF Then
                            If (adoRs!valuee < 0 And adoRs!valuee < 0.1) Then
                                Cells(j, col).Value = "Inap"
                            Else
                                Cells(j, col).Value = adoRs!valuee
                            End If
                        End If
                    adoRs.Close
                End If
            Else
                rojo col - 1, j
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
    Dim bandera As Boolean
    
    clvObs(1) = "TXPVC"
    clvObs(2) = "XOBVC"
    clvObs(3) = "ORZVC"
    clvObs(4) = "VERVC"
    clvObs(5) = "COTVC"
    clvObs(6) = "RALVC"
    
    inicioFil = 11
    
    'Obtiene el número de la última fila
    ultFil = Range("B" & rows.Count).End(xlUp).Row
    
    bandera = True
    
    If (fecha = "") Then
        'Obtiene la fecha actual
        'fecha = Format(Now, "yyyy/mm/dd")
        bandera = False
    End If
    
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = "SIH"
    dbSIH.Open
    
    For i = 1 To 6
        col = i * 3
        For j = inicioFil To ultFil - 1
            lluvia = Format(Cells(j, col).Value, "0.0")
            If (lluvia <> "") Then
                hora = Format(Cells(j, col - 1).Value, "hh:mm")
                If (IsDate(hora)) Then
                    If (hora <> "17:00" And hora <> "07:00") Then
                        'Valida lluvia
                        If Not (IsNumeric(lluvia)) Then
                            If (lluvia = "inap" Or lluvia = "INAP" Or lluvia = "Inap") Then
                                lluvia = 0.01
                                blanco col, j
                            Else
                                rojo col, j
                                bandera = False
                            End If
                        ElseIf (CDbl(lluvia) >= 0) Then
                            If (CDbl(lluvia) <> 0.01) Then
                                lluvia = Format(lluvia, "0.0")
                                blanco col, j
                            End If
                        Else
                            rojo col, j
                            bandera = False
                        End If
                        If bandera Then
                            query = "REPLACE INTO dtprecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + clvObs(i) + "', '" + fecha + " " + hora + "', '" + lluvia + "', '" + lluvia + "', ' ', 'XL', ' ')"
                            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
                        End If
                    End If
                Else
                    bandera = False
                    rojo col - 1, j
                End If
            End If
        Next j
    Next i
    
    'Fin de la conexión
    dbSIH.Close
    
    If Not bandera Then
        MsgBox "Alguna informacion es incorrecta", vbCritical, "ERROR"
    End If

End Sub
Public Function setFecha(fec As String)
    fecha = fec
End Function
Public Function getFecha() As String
    getFecha = fecha
End Function

Private Sub rojo(col As Integer, rows As Integer)
    Cells(rows, col).Interior.Color = vbRed
End Sub
Private Sub blanco(col As Integer, rows As Integer)
    Cells(rows, col).Interior.Color = xlNone
End Sub

