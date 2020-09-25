Attribute VB_Name = "dataBase"
Option Explicit

'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String

Public dns As String
Public grp As String

Public dosNiveles() As String


Sub iniDataBase()
    If Seguimiento.dns = "" Then
        Seguimiento.iniciaSeg
    End If
    dns = Seguimiento.dns
End Sub
''****************VALORES DE NIVEL*********************
Sub addNivel(valorNiv As String, idEstacion As Integer, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    validarNiv.validaNivel valorNiv, idEstacion, fecha
    If validarNiv.edoValidacion > 0 Then
        'Conexión a la base de datos
        dbSIH.ConnectionString = dns
        dbSIH.Open
            'Consulta insertar nuevo dato
            query = "INSERT INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" & Estaciones.clvEst(0, idEstacion) & _
                    "', '" & fecha & "', '" & valorNiv & "', '" + valorNiv + "', ' ', 'XL', ' ')"
            'query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + Estaciones.clvEst(i) + "', '" + fecha + " " + hora + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
            'Cierra resultado de consulta
            'adoRs.Close
        'Fin de la conexión
        dbSIH.Close
    End If
End Sub

Sub repNivel(valorNiv As String, idEstacion As Integer, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    validarNiv.validaNivel valorNiv, idEstacion, fecha
    If validarNiv.edoValidacion > 0 Then
        'Conexión a la base de datos
        dbSIH.ConnectionString = dns
        dbSIH.Open
            'Consulta remplazar dato
            query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" & Estaciones.clvEst(0, idEstacion) & _
                    "', '" & fecha & "', '" & valorNiv & "', '" + valorNiv + "', ' ', 'XL', ' ')"
            'query = "REPLACE INTO DTNivel (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" + Estaciones.clvEst(i) + "', '" + fecha + " " + hora + "', '" + niv + "', '" + niv + "', ' ', 'XL', ' ')"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
            'Cierra resultado de consulta
            'adoRs.Close
        'Fin de la conexión
        dbSIH.Close
    End If
End Sub

Sub eliminarNiv(idEstacion As Integer, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        resp = MsgBox("Seguro quiere eliminar el nivel de " & Estaciones.clvEst(0, idEstacion) & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM dtNivel WHERE station = '" & Estaciones.clvEst(0, idEstacion) & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

''****************VALORES DE LLUVIA*********************
Sub addLluvia(valorLlu As String, idEstacion As Integer, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DTPrecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" & Estaciones.clvEst(0, idEstacion) & _
                "', '" & fecha & "', '" & valorLlu & "', '" + valorLlu + "', ' ', 'XL', ' ')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repLluvia(valorLlu As String, idEstacion As Integer, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DTPrecipitacio (station, datee, valuee, corrvalue, msgcode, source, timewidth) VALUES ('" & Estaciones.clvEst(0, idEstacion) & _
                "', '" & fecha & "', '" & valorLlu & "', '" + valorLlu + "', ' ', 'XL', ' ')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarLluvia(idEstacion As Integer, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar la LLUVIA de " & Estaciones.infEst(1, idEstacion) & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DTPrecipitacio WHERE station = '" & Estaciones.clvEst(0, idEstacion) & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Cierra resultado de consulta
            'adoRs.Close
        End If
    
    'Fin de la conexión
    dbSIH.Close
End Sub

Function lluviaAcumulada(idEstacion As String, fechaInicial As String, fechaFinal As String) As String
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        query = "SELECT SUM(Valuee)as Acum FROM dtPrecipitacio WHERE station = '" & idEstacion & "' AND Datee > '" & fechaInicial & "' AND Datee < '" & fechaFinal & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            If Not adoRs.EOF Then
                lluviaAcumulada = Format(adoRs!Acum, "0.0")
            Else
                lluviaAcumulada = 0
            End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function


Sub getNivel()

End Sub

Function getUltNiv(idEstacion As Integer, fecha As String) As String
    'Variable para rango de fechas
    Dim d As String
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Rango de fecha para obtener el último nivel (2 días atras)
    d = Format(DateAdd("d", -2, fecha), "yyyy/mm/dd")
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    'Consulta el último nivel de escala registrado al día
    query = "SELECT valuee AS val FROM dtNivel WHERE station = '" & Estaciones.clvEst(0, idEstacion) & _
            "' AND datee >= '" & d & "' AND datee < '" & Format(fecha, "yyyy/mm/dd hh:mm") & "' ORDER BY Datee DESC LIMIT 1"
            'MsgBox query
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            'Devuelve el resultado de la consulta
            getUltNiv = adoRs!val
        End If
    adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function


Sub getDosUNivls(idEstacion As Integer, fecha As String)
    'Variable para rango de fechas
    Dim d As String
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    ReDim dosNiveles(1)
    'Rango de fecha para obtener el último nivel (2 días atras)
    d = Format(DateAdd("d", -2, fecha), "yyyy/mm/dd")
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    'Consulta el último nivel de escala registrado al día
    query = "SELECT valuee AS val FROM dtNivel WHERE station = '" & Estaciones.clvEst(0, idEstacion) & _
            "' AND datee >= '" & d & "' AND datee < '" & Format(fecha, "yyyy/mm/dd 23:59") & "' ORDER BY Datee DESC LIMIT 2"
            'MsgBox query
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            'Devuelve el resultado de la consulta
            dosNiveles(0) = adoRs!val
            adoRs.MoveNext
            dosNiveles(1) = adoRs!val
        End If
    adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Function hayDosNiveles(idEstacion As Integer, fecha As String) As Boolean
    'Variable para rango de fechas
    Dim d As String
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Rango de fecha para obtener el último nivel (2 días atras)
    d = Format(DateAdd("d", -2, fecha), "yyyy/mm/dd")
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    'Consulta el último nivel de escala registrado al día
    query = "SELECT count(*) AS n FROM dtNivel WHERE station = '" & Estaciones.clvEst(0, idEstacion) & _
            "' AND datee >= '" & d & "' AND datee < '" & Format(fecha, "yyyy/mm/dd 23:59") & "'"
            'MsgBox query
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            If adoRs!n > 1 Then hayDosNiveles = True
            Else: hayDosNiveles = False
        End If
    adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function
