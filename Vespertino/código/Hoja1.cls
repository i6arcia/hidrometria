VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub cmdBtnRefresh_Click()
    Range("C6").Value = "Xalapa, Ver., a " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy")
    Vesp.acumuladas ("17:00")
    Vesp.ultNiv
    Vesp.desviacionStd
    Vesp.getDatos
End Sub
