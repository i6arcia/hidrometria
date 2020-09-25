Attribute VB_Name = "CapturaMatutino"
Option Explicit
'Variables para hojas de excel
Dim hPresas As Excel.Worksheet
Dim hHidro As Excel.Worksheet
Dim hNorte As Excel.Worksheet
Dim hSur As Excel.Worksheet
Dim hPluvio As Excel.Worksheet

'Control de fechas
Public fecha As String
Public ayer As String

Public respuestaFrm As Boolean

Public editando As Boolean


Sub inicia()
    'Asigna hojas a variables
    Set hPresas = Worksheets("PRESAS")
    Set hHidro = Worksheets("HIDROMETRICA")
    Set hNorte = Worksheets("Norte")
    Set hSur = Worksheets("Sur")
    Set hPluvio = Worksheets("Pluviometros")
    
    'Fecha con la cual se va a trabajar
    fecha = Format(Now, "yyyy/mm/dd")
    'fecha = "2020/09/23"
    ayer = Format(DateAdd("d", -1, fecha), "yyyy/mm/dd")

End Sub

Sub titulos()
    'Bandera para editar hoja
    editando = True
    'Valida que este iniciadas las variables
    If fecha = "" Then inicia
    
    'Titulos de hojas
    hPresas.Range("B5").Value = "Xalapa, Ver. -- " & Format(fecha, "dddd") & " " & _
                            Format(fecha, "dd") & " de " & Format(fecha, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
    hHidro.Range("B5").Value = "Xalapa, Ver. -- " & Format(fecha, "dddd") & " " & _
                            Format(fecha, "dd") & " de " & Format(fecha, "mmmm") & _
                            " de " & Format(fecha, "yyyy") & " --"
    hNorte.Range("B5").Value = "Xalapa, Ver. -- " & Format(fecha, "dddd") & " " & _
                            Format(fecha, "dd") & " de " & Format(fecha, "mmmm") & _
                            " de " & Format(fecha, "yyyy") & " --"
    hSur.Range("B5").Value = "Xalapa, Ver. -- " & Format(fecha, "dddd") & " " & _
                            Format(fecha, "dd") & " de " & Format(fecha, "mmmm") & _
                            " de " & Format(fecha, "yyyy") & " --"
    hPluvio.Range("B5").Value = "Xalapa, Ver. -- " & Format(fecha, "dddd") & " " & _
                            Format(fecha, "dd") & " de " & Format(fecha, "mmmm") & _
                            " de " & Format(fecha, "yyyy") & " --"
    'Enfasis en color de la fecha
    If (fecha = Format(Now, "yyyy/mm/dd")) Then 'AZUL
        hPresas.Range("B5").Interior.Color = RGB(220, 230, 241)
        hHidro.Range("B5").Interior.Color = RGB(220, 230, 241)
        hNorte.Range("B5").Interior.Color = RGB(220, 230, 241)
        hSur.Range("B5").Interior.Color = RGB(220, 230, 241)
        hPluvio.Range("B5").Interior.Color = RGB(220, 230, 241)
    Else    'BLANCO
        hPresas.Range("B5").Interior.Color = vbWhite
        hHidro.Range("B5").Interior.Color = vbWhite
        hNorte.Range("B5").Interior.Color = vbWhite
        hSur.Range("B5").Interior.Color = vbWhite
        hPluvio.Range("B5").Interior.Color = vbWhite
    End If
    'Cierra bandera para editar hoja
    editando = False
End Sub

Sub actualizaHojas()
    If dataBase.pruebaBD Then
        ctrlCambios.llenaMatriz
        ctrlCambios.generaRangos
        
        Presas.limpiaHojaPresas
        
        hidrometrica.iniciaHidro
        hidrometrica.limpiaHidro
        climaNorte.iniciaClmNrt
        climaNorte.limpiaClmNrt
        climaSur.iniciaClmSur
        climaSur.limpiaClmSur
        pluviometros.iniciaPluvios
        pluviometros.limpiaPluvios
        
        'inicia
        titulos
        
        Presas.obtenerPresas
        hidrometrica.obtieneHidro
        climaNorte.obtieneClmNrt
        climaSur.obtieneClmSur
        pluviometros.obtienePluvios
        
    Else
        Presas.limpiaHojaPresas
        hidrometrica.limpiaHidro
        climaNorte.limpiaClmNrt
        climaSur.limpiaClmSur
        pluviometros.limpiaPluvios
    End If
End Sub
