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

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Range("K6").Value = Format(Now, "short date")
    Range("K6").Font.Color = vbWhite
End Sub

Private Sub Workbook_Open()
    Dim fchAct As String
    Dim fchReg As String
    Dim lastRow As Integer
    
    fchAct = Format(Now, "Short Date")
    fchReg = Range("k6").Value
    If (fchAct <> fchReg) Then
        'Obtiene el numero de la ultima fila
        lastRow = Range("B" & rows.Count).End(xlUp).Row
        'Selecciona contenido y lo limpia
        Range("F9:H" & lastRow).ClearContents
    End If
    Vesp.acumuladas
    Vesp.ultNiv
    Vesp.desviacionStd
End Sub
