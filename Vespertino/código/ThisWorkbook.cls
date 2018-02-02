VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'*********************************************************************
'   Sistema en apoyo a la captura de informacion hidroclimatologica
'          para la dirección técnica en OCGC, CONAGUA
'*********************************************************************


Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Vesp.capturar
End Sub

Private Sub Workbook_Open()
    
    Range("C6").Value = "Xalapa, Ver., a " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy")
    
    Vesp.acumuladas ("17:00")
    Vesp.ultNiv
    Vesp.desviacionStd
    Vesp.getDatos
End Sub


