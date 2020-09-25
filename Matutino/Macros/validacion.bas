Attribute VB_Name = "validacion"
Option Explicit

Dim val As String
Dim clv As String
Dim fch As String
Dim respFormulario As Boolean

Dim prs As Excel.Worksheet

Private Sub iniciaValidacion(i As Integer)
    Set prs = Worksheets("PRESAS")
    val = prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2))
    clv = ctrlCambios.mCambios(i, 0)
    fch = CapturaMatutino.fecha
End Sub

'1|  Validación Nivel

'Function validaNivel(i As Integer) As Integer
Function validaNivel(niv As String, clv As String, fch As String, edo As String) As Boolean
    If IsNumeric(niv) Then
        'Agrega o modifica
        niv = Format(niv, "0.00")
        If edo = 2 Then      'AGREGAR
            dataBase.addNivel niv, clv, fch
        ElseIf edo = 3 Then  'MODIFICAR
            dataBase.repNivel niv, clv, fch
        End If
        validaNivel = True
    ElseIf niv = "" Or niv = "ddd" Or niv = "DDD" Then
        'Elimina
        dataBase.eliminarNiv clv, fch
        validaNivel = True
    Else
        'ERROR
        validaNivel = False
    End If
End Function

Function validaNivel2(niv As String, clv As String, Optional nivAnt As String, Optional destd As String) As Boolean
    Dim rMas As Double
    Dim rMenos As Double
    
    If IsNumeric(niv) Then
        If nivAnt <> "" And destd <> "" Then
            rMas = CDbl(nivAnt) + CDbl(destd)
            rMenos = CDbl(nivAnt) - CDbl(destd)
            If CDbl(niv) >= rMenos And CDbl(niv) <= rMas Then
                validaNivel2 = True
            Else
                'Ventana Ignorar diferencia
                frmVerificar.loadInfo clv, niv, nivAnt, destd
                frmVerificar.Show
                validaNivel2 = CapturaMatutino.respuestaFrm
            End If
        Else
            'Agrega o modifica
            validaNivel2 = True
        End If
    ElseIf niv = "" Or niv = "ddd" Or niv = "DDD" Then
        'Elimina
        validaNivel2 = True
    Else
        'ERROR
        validaNivel2 = False
    End If
    
End Function

Function validaNivel3(niv As String, clv As String, fch As String) As Boolean
    Dim rMas As Double
    Dim rMenos As Double
    Dim namo As String
    Dim desStd As String
    Dim ultNiv As String
    'Valida que sea número
    If IsNumeric(niv) Then
        'Recupera variables NAMO | DESVIACION ESTANDAR | ÚLTIMO NIVEL
        namo = dataBase.getNamo(clv)
        desStd = dataBase.getDesviacionStd(clv, fch)
        ultNiv = dataBase.getUltNiv(clv, fch)
        
        If ultNiv <> "" And desStd <> "" Then
            rMas = CDbl(ultNiv) + CDbl(desStd)
            rMenos = CDbl(ultNiv) - CDbl(desStd)
            
            If CDbl(niv) >= rMenos And CDbl(niv) <= rMas Then
                validaNivel3 = True
            Else
                'Ventana Ignorar diferencia
                frmVerificar.loadInfo clv, niv, ultNiv, desStd
                frmVerificar.Show
                validaNivel3 = CapturaMatutino.respuestaFrm
            End If
        Else
            'Agrega o modifica
            validaNivel3 = True
        End If
    ElseIf niv = "" Or niv = "ddd" Or niv = "DDD" Then
        'Elimina
        validaNivel3 = True
    Else 'El valor es incorrecto
        'ERROR
        validaNivel3 = False
    End If
    
End Function


'2|  Validación Almacenamiento
Function validaAlmacenamiento(alm As String, clv As String, fch As String, edo As String) As Boolean
    If IsNumeric(alm) Then
        'Agrega o modifica
        alm = Format(alm, "0.000")
        If edo = 2 Then      'AGREGAR
            dataBase.addAlmacenamiento alm, clv, fch
        ElseIf edo = 3 Then  'MODIFICAR
            dataBase.repAlmacenamiento alm, clv, fch
        End If
        validaAlmacenamiento = True
    ElseIf alm = "" Or alm = "ddd" Or alm = "DDD" Then
        'Elimina
        dataBase.eliminarAlm clv, fch
        validaAlmacenamiento = True
    Else
        'ERROR
        validaAlmacenamiento = False
    End If
End Function
'3|  Validación Gasto
Function validaGasto(i As Integer)
    iniciaValidacion i
    If IsNumeric(val) Then
        val = Format(val, "0.000")
        If ctrlCambios.mCambios(i, 4) = 2 Then      'AGREGAR
            dataBase.addGasto val, clv, fch
        ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'MODIFICAR
            dataBase.repGasto val, clv, fch
        End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then  'EliminaR
            dataBase.eliminarGasto clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If

End Function

Function validaLluvia(val As String, clv As String, fch As String, edo As String) As Boolean
    If IsNumeric(val) Then
        If CDbl(val) >= 0 Then
            If CDbl(val) > 0 And CDbl(val) <= 0.01 Then
                val = "0.01"
            Else
                val = Format(val, "0.0")
            End If
            If edo = 2 Then     'Agregar
                dataBase.addLluvia val, clv, fch
            ElseIf edo = 3 Then  'Modificar
                dataBase.repLluvia val, clv, fch
            End If
            validaLluvia = True
        Else
            'ERROR
            validaLluvia = False
        End If
    ElseIf val = "inap" Or val = "INAP" Or val = "Inap" Then
        val = "0.01"
        If edo = 2 Then      'Agregar
            dataBase.addLluvia val, clv, fch
        ElseIf edo = 3 Then 'Modificar
            dataBase.repLluvia val, clv, fch
        End If
        validaLluvia = True
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        dataBase.eliminarLluvia clv, fecha
        validaLluvia = True
    Else
        'ERROR
        validaLluvia = False
    End If
End Function

Function validaLluvia2(val As String) As Boolean
    
    If IsNumeric(val) Then
        If CDbl(val) >= 0 Then
            validaLluvia2 = True
        Else
            'ERROR
            validaLluvia2 = False
        End If
    ElseIf val = "inap" Or val = "INAP" Or val = "Inap" Then
        validaLluvia2 = True
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        validaLluvia2 = True
    Else
        'ERROR
        validaLluvia2 = False
    End If
End Function

Function validaTemps(amb As String, max As String, min As String, clv As String, fch As String, edo As String) As Boolean
    Dim media As Double
    Dim bandera As Boolean
    bandera = True
    If Not (IsNumeric(amb)) Then bandera = False
    If Not (IsNumeric(max)) Then bandera = False
    If Not (IsNumeric(min)) Then bandera = False
    If bandera Then
        If CDbl(max) >= CDbl(amb) And CDbl(amb) >= CDbl(min) Then
            amb = Format(amb, "0.0")
            max = Format(max, "0.0")
            min = Format(min, "0.0")
            
            media = Round((CDbl(max) + CDbl(min)) / 2, 1)
            
            If edo = 2 Then      'AGREGAR
                dataBase.addTemps amb, max, min, CStr(media), clv, fch
            ElseIf edo = 3 Then  'MODIFICAR
                dataBase.repTemps amb, max, min, CStr(media), clv, fch
            End If
            validaTemps = True
        Else
            'MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
            validaTemps = False
        End If
    ElseIf amb = "" And max = "" And min = "" Then
        dataBase.eliminarTemps clv, fch
        validaTemps = True
    Else
        'MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        validaTemps = False
    End If
End Function

Function validaPresion(valPresion As String, clv As String, fch As String, edo As String) As Boolean
    If IsNumeric(valPresion) Then
        If valPresion > 100 Then
            'Validar valores min max amb
            valPresion = Format(valPresion, "0.0")
            If edo = 2 Then      'AGREGAR
                dataBase.addPresion valPresion, clv, fch
            ElseIf edo = 3 Then  'MODIFICAR
                dataBase.repPresion valPresion, clv, fch
            End If
            validaPresion = True
        Else
            validaPresion = False
        End If
    ElseIf valPresion = "" Or val = "ddd" Or val = "DDD" Then  'EliminaR
            dataBase.eliminarPresion clv, fch
            validaPresion = True
    Else
        'ERROR
        'MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        'Devolver error
        validaPresion = False
    End If
End Function

Function validaHumedad(valHumedad As String, clv As String, fch As String, edo As String) As Boolean
    If IsNumeric(valHumedad) Then
        If valHumedad > 0 Then
            'Validar valores min max amb
            valHumedad = Format(valHumedad, "0")
            If edo = 2 Then      'AGREGAR
                dataBase.addHumedad valHumedad, clv, fch
            ElseIf edo = 3 Then  'MODIFICAR
                dataBase.repHumedad valHumedad, clv, fch
            End If
            validaHumedad = True
        Else
            validaHumedad = False
        End If
    ElseIf valHumedad = "" Or val = "ddd" Or val = "DDD" Then  'EliminaR
            dataBase.eliminarHumedad clv, fch
            validaHumedad = True
    Else
        'ERROR
        'MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        'Devolver error
        validaHumedad = False
    End If
End Function

Function validaEvap(valEvap As String, clv As String, fch As String, edo As String) As Boolean
    If IsNumeric(valEvap) Then
        If valEvap > 0 Then
            'Validar valores min max amb
            valEvap = Format(valEvap, "0.00")
            If edo = 2 Then      'AGREGAR
                dataBase.addEvap valEvap, clv, fch
            ElseIf edo = 3 Then  'MODIFICAR
                dataBase.repEvap valEvap, clv, fch
            End If
            validaEvap = True
        Else
            validaEvap = False
        End If
    ElseIf valEvap = "" Or val = "ddd" Or val = "DDD" Then  'EliminaR
            dataBase.eliminarEvap clv, fch
            validaEvap = True
    Else
        'ERROR
        'MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        'Devolver error
        validaEvap = False
    End If
End Function


'4|  Validación Lluvia
'Sub validaLluvia(i As Integer)
Sub validaLluviaPresas(i As Integer)
    iniciaValidacion i
    If IsNumeric(val) Then
        If CDbl(val) >= 0 Then
            If CDbl(val) > 0 And CDbl(val) <= 0.01 Then
                val = "0.01"
            Else
                val = Format(val, "0.0")
            End If
            If ctrlCambios.mCambios(i, 4) = 2 Then      'Agregar
                dataBase.addLluvia val, clv, fch
            ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'Modificar
                dataBase.repLluvia val, clv, fch
            End If
        Else
            'ERROR
            MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
            prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
        End If
    ElseIf val = "inap" Or val = "INAP" Or val = "Inap" Then
        val = "0.01"
        If ctrlCambios.mCambios(i, 4) = 2 Then      'Agregar
            dataBase.addLluvia val, clv, fch
        ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'Modificar
            dataBase.repLluvia val, clv, fch
        End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        dataBase.eliminarLluvia clv, fecha
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If
End Sub
'5|  Validación Evaporación
Function validaEvapPresas(i As Integer) As Integer
    iniciaValidacion i
    If IsNumeric(val) Then
        'Agrega o modifica
        val = Format(val, "0.00")
        If ctrlCambios.mCambios(i, 4) = 2 Then      'AGREGAR
            dataBase.addEvap val, clv, fch
        ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'MODIFICAR
            dataBase.repEvap val, clv, fch
        End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        'Elimina
            dataBase.eliminarEvap clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If
End Function
'6|  Validación Vertedor
Function validaVertedor(i As Integer) As Integer
    iniciaValidacion i
    If IsNumeric(val) Then
        'Agrega o modifica
        val = Format(val, "0.00")
        If ctrlCambios.mCambios(i, 4) = 2 Then      'AGREGAR
            dataBase.addVertedor val, clv, fch
        ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'MODIFICAR
            dataBase.repVertedor val, clv, fch
        End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        'Elimina
            dataBase.eliminarVertedor clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If
End Function
'7|  Validación O.T.
Function validaOt2(i As Integer) As Integer
    iniciaValidacion i
    If IsNumeric(val) Then
        'Agrega o modifica
        
        val = Format(val, "0")
        If ctrlCambios.mCambios(i, 4) = 2 Then      'AGREGAR
            dataBase.addOT2 val, clv, fch
        ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'MODIFICAR
            dataBase.repOT2 val, clv, fch
        End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        'Elimina
            dataBase.eliminarOT2 clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If
End Function

'8| Validación Gasto Río
Function validaGastoRio(i As Integer) As Integer
    iniciaValidacion i
    If IsNumeric(val) Then
        'Agrega o modifica
        val = Format(val, "0.000")
        If ctrlCambios.mCambios(i, 4) = 2 Then      'AGREGAR
            dataBase.addGastoRio val, clv, fch
        ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'MODIFICAR
            dataBase.repGastoRio val, clv, fch
        End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        'Elimina
            dataBase.eliminarGastoRio clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If
End Function

'9|  Validación Área Elevación
Function validaArea(clav As String, valArea As String) As Integer
    clv = clav
    val = valArea
    fch = CapturaMatutino.fecha
    
    If IsNumeric(val) Then
        'Agrega o modifica
        val = Format(val, "0")
        'If ctrlCambios.mCambios(i, 4) = 2 Then      'AGREGAR
        '    dataBase.addAlmacenamiento val, clv, fch
        'ElseIf ctrlCambios.mCambios(i, 4) = 3 Then  'MODIFICAR
            dataBase.repArea val, clv, fch
        'End If
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        'Elimina
        '    dataBase.eliminaArea clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        'prs.Cells(ctrlCambios.mCambios(i, 1), ctrlCambios.mCambios(i, 2)).Interior.Color = vbRed
    End If
End Function

'------------------------------------------------
'********             TENDENCIAS         ********
'------------------------------------------------

'------------------------------------------------
'********           COLORES           ********
'------------------------------------------------

Function colorLluvia(val As String) As String
    ' Caracteres para establecer color de celda
    ' c -> Correcto
    ' x -> incorrecto
    ' m -> Relevancia media
    ' a -> Relevancia máxima
    
    'Es numerico el valor
    If IsNumeric(val) Then
        If CDbl(val) >= 0 And CDbl(val) < 1000 Then
            'Correcto
            If CDbl(val) < 20 Then
                colorLluvia = "c"
                'Correcto
            ElseIf CDbl(val) >= 20 And CDbl(val) < 50 Then
                'Relevante medio
                colorLluvia = "m"
            ElseIf CDbl(val) >= 50 Then
                'Relevante máximo
                colorLluvia = "a"
            End If
        Else
            'ERROR
            colorLluvia = "x"
        End If
    ElseIf val = "inap" Or val = "INAP" Or val = "Inap" Then
        'Correcto
        colorLluvia = "c"
    ElseIf val = "" Or val = "ddd" Or val = "DDD" Then
        'Correcto
        colorLluvia = "c"
    Else
        'ERROR
        colorLluvia = "x"
    End If
End Function

Function colorEvap(val As String)
    ' Caracteres para establecer color de celda
    ' c -> Correcto
    ' x -> incorrecto
    ' m -> Relevancia media
    ' a -> Relevancia máxima
    
    'Es numerico el valor
    If IsNumeric(valEvap) Then
        If valEvap > 0 Then
            'Validar valores min max amb
            valEvap = Format(valEvap, "0.00")
            If edo = 2 Then      'AGREGAR
                dataBase.addEvap valEvap, clv, fch
            ElseIf edo = 3 Then  'MODIFICAR
                dataBase.repEvap valEvap, clv, fch
            End If
        End If
    ElseIf valEvap = "" Or val = "ddd" Or val = "DDD" Then  'EliminaR
            dataBase.eliminarEvap clv, fch
    Else
        'ERROR
        MsgBox "Algunos campos capturados no son correctos", vbCritical, "Error en captura"
        'Devolver error
    End If
End Function

