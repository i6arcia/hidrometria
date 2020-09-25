Attribute VB_Name = "ExportarPDF"
Sub exporPDF()
    Dim dir1 As String
    Dim dir2 As String
    Dim hj As Worksheet
    
    On Error GoTo problemas
    Set hj = Worksheets("CLAVES")
    
    dir1 = "\\gc0105w222\Users\jlunal\Desktop\reportes\reporte diario\PruebaPDF"
    dir2 = "\\gc0105w222\Users\jlunal\Desktop\reportes\reporte diario\PruebaPDF"
    If hj.Range("F5").Value <> "" Then
        dir1 = hj.Range("F5").Value & "\INFORME MATUTINO"
    End If
    If hj.Range("F6").Value <> "" Then
        dir2 = hj.Range("F6").Value & "\REPORTEDIARIO" & hj.Range("G2").Value
    End If
    
    
    'Seleccionar hojas
    ActiveWorkbook.Sheets(Array("Presas", "HIDRO", "CLIMA1", "CLIMA2", "CLIMA3", "RESUMEN")).Select
    
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dir1, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dir2, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Exit Sub
problemas:
MsgBox "No se pudo exportar PDF" & vbCrLf & "Probablemente el documento se encuentre abierto o el directorio no exista", vbCritical, "Problemas para exportar PDF"
End Sub
