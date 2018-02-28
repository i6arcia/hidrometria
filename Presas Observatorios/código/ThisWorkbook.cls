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
'           Sistema en apoyo a la captura de Lluvia y niveles
'               de la dirección técnica en OCGC, CONAGUA
'
'*********************************************************************

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Observatorios.capturarDatos
    Observatorios.actualizar
    Presas.capturarDatos
    Presas.obtenerDatos
End Sub

Private Sub Workbook_Open()
    Observatorios.inicio
    Presas.inicio
End Sub
