VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim bandera As Boolean

Private Sub lblNext1_Click()
    Dim f As String
    Dim fD As String
    
    f = Presas.getFecha
        
    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If
    
    fD = Format(DateAdd("d", 1, f), "yyyy/mm/dd")
    
    Range("E7").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("E7").Interior.Color = RGB(221, 235, 247)
    Else
        Range("E7").Interior.Color = xlNone
    End If
    
    Presas.setFecha fD
    Presas.obtenerDatos
End Sub

Private Sub lblNow1_Click()
    Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    Range("E7").Interior.Color = RGB(221, 235, 247)
    Presas.setFecha Format(Now, "yyyy/mm/dd")
    Presas.obtenerDatos
End Sub

Private Sub lblPrev1_Click()
    Dim f As String
    Dim fD As String
    
    f = Presas.getFecha
        
    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If
    
    fD = Format(DateAdd("d", -1, f), "yyyy/mm/dd")
    
    Range("E7").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("E7").Interior.Color = RGB(221, 235, 247)
    Else
        Range("E7").Interior.Color = xlNone
    End If
    
    Presas.setFecha fD
    Presas.obtenerDatos
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ultFil As Integer
    Dim fil As Integer
    Dim col As Integer
    Dim dato As Double
    Dim val As String
    If Not bandera Then
        ultFil = Range("B" & rows.Count).End(xlUp).Row
        If Not Intersect(Target, Range("C12:D" & ultFil)) Is Nothing Then
            fil = Target.Row
            col = Target.Column
            If fil >= 12 And fil <= ultFil Then
                If col = 3 Then
                        val = Cells(fil, col).Value
                        If val <> "" Then
                            If IsNumeric(val) Then
                                dato = Format(val, "0.00")
                                Target.Interior.Color = RGB(255, 244, 204)
                                bandera = True
                                Cells(fil, 4).Value = dato - 2.79
                            Else
                                Target.Interior.Color = vbRed
                            End If
                        Else
                            bandera = True
                            Cells(fil, 4).ClearContents
                        End If
                    
                ElseIf col = 4 Then
                    val = Target.Value
                    If val <> "" Then
                        If IsNumeric(val) Then
                            dato = Format(val, "0.00")
                            Target.Interior.Color = xlNone
                            bandera = True
                            Cells(fil, 3).Value = dato + 2.79
                        Else
                            Target.Interior.Color = vbRed
                        End If
                    Else
                        bandera = True
                        Cells(fil, 3).ClearContents
                    End If
                End If
            End If
        End If
    Else
        bandera = False
    End If
End Sub
