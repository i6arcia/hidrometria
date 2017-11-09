VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Capturar los datos en SIH despues de que se ejecute el comando de guardar archivo
Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Vesp.capturar
End Sub
Private Sub Workbook_Open()
    
    Range("C6").Value = "Xalapa, Ver., a " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy")
    
    Vesp.getDatos
    Vesp.acumuladas
    Vesp.ultNiv
    Vesp.desviacionStd
End Sub
