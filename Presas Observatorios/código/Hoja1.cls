VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub lblNext_Click()
    Dim f As String
    Dim fD As String
    
    f = Observatorios.getFecha
    
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
    
    Observatorios.setFecha fD
    Observatorios.obtenerDatos
End Sub

Private Sub lblNow_Click()
    Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    Range("E7").Interior.Color = RGB(221, 235, 247)
    Observatorios.setFecha Format(Now, "yyyy/mm/dd")
    Observatorios.obtenerDatos
End Sub

Private Sub lblPrev_Click()
    Dim f As String
    Dim fD As String
    
    f = Observatorios.getFecha
    
    If (f = "") Then
        f = Format(Now, "yyyy/mm/dd")
    End If
    
    fD = Format(DateDiff("d", 1, f), "yyyy/mm/dd")
    
    Range("E7").Value = "Xalapa, Ver. -- " & Format(fD, "dddd") & " " & Format(fD, "dd") & " de " & Format(fD, "mmmm") & " de " & Format(fD, "yyyy") & " --"
    If (fD = Format(Now, "yyyy/mm/dd")) Then
        Range("E7").Interior.Color = RGB(221, 235, 247)
    Else
        Range("E7").Interior.Color = xlNone
    End If
    
    Observatorios.setFecha fD
    Observatorios.obtenerDatos
End Sub
