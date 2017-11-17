VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Observatorios.capturarDatos
    Observatorios.obtenerDatos
End Sub

Private Sub Workbook_Open()
    Range("E7").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & Format(Now, "dd") & " de " & Format(Now, "mmmm") & " de " & Format(Now, "yyyy") & " --"
    Observatorios.setFecha Format(Now, "yyyy/mm/dd")
    Observatorios.obtenerDatos
End Sub
