Attribute VB_Name = "dataBase"
Option Explicit

'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String

Public dns As String
Public grp As String

Public dosNiveles() As String

Public mValGrp(100, 1) As String
Public iValG As Integer

Public temperaturas(3) As String

Private Sub iniDataBase()
    dns = "SIh"
End Sub

Function pruebaBD() As Boolean
On Error GoTo msg
    iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    'Fin de la conexión
    dbSIH.Close
    pruebaBD = True
    Exit Function
msg:
    MsgBox "Existen problemas en la conexión a la base de datos", vbCritical
    pruebaBD = False
End Function
''****************VALORES DE NIVEL*********************
Sub addNivel(valorNiv As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'validarNiv.validaNivel valorNiv, idEstacion, fecha
    'If validarNiv.edoValidacion > 0 Then
        'Conexión a la base de datos
        dbSIH.ConnectionString = dns
        dbSIH.Open
            'Consulta insertar nuevo dato
            query = "INSERT INTO DdxNivel (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numdia, cantEstac) " & _
                                  "VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorNiv & "', '', '" & fecha & "', '', '" & valorNiv & "', '" & fecha & "', '', '" & valorNiv & "', '0', '" & Day(fecha) & "', '0')"
            'MsgBox query
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
            'Cierra resultado de consulta
        'Fin de la conexión
        dbSIH.Close
    'End If
End Sub

Sub repNivel(valorNiv As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'validarNiv.validaNivel valorNiv, idEstacion, fecha
    'If validarNiv.edoValidacion > 0 Then
        'Conexión a la base de datos
        dbSIH.ConnectionString = dns
        dbSIH.Open
            'Consulta remplazar dato
            'Consulta insertar nuevo dato
            query = "REPLACE INTO DdxNivel (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numdia, cantEstac)" & _
                                  " VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorNiv & "', '', '" & fecha & "', '', '" & valorNiv & "', '" & fecha & "', '', '" & valorNiv & "', '0', '" & Day(fecha) & "', '0')"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Fin de la conexión
        dbSIH.Close
    'End If
End Sub

Sub eliminarNiv(idEstacion As String, fecha As String)
    Dim resp As Integer
    Dim nomEstacion As String
    'valida que esten iniciadas las variables
    If dns = "" Then iniDataBase
    
    nomEstacion = dataBase.getNombreEstacion(idEstacion)
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        resp = MsgBox("Seguro quiere eliminar el nivel de " & nomEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdxNivel WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Function getNivel(idEstacion As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*********************************
    '   Consulta nivel de la estación
    '*********************************
    getNivel = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue as Niv FROM DdxNivel where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getNivel = Format(adoRs!niv, "0.00")
        Else
            getNivel = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getNivelGrp(idGrupo As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*******************************************
    '   Consulta nivel de un grupo de estacines
    '*******************************************
    iValG = 0
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT t1.station, t1.DailyValue as Niv FROM DdxNivel t1, stationGroups t2 where t2.stationgroup = '" & idGrupo & "' and t1.station = t2.station and t1.datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        Do While Not adoRs.EOF
            'Almacena valor de nivel
            mValGrp(iValG, 0) = adoRs!station
            mValGrp(iValG, 1) = Format(adoRs!niv, "0.00")
            iValG = iValG + 1
            adoRs.MoveNext
        Loop
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

''****************VALORES DE ALMACENAMIENTO*********************
Sub addAlmacenamiento(valorAlm As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO ddxVolAlmac (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numdia, cantEstac) " & _
                              "VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorAlm & "', '', '" & fecha & "', '', '" & valorAlm & "', '" & fecha & "', '', '" & valorAlm & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        'Cierra resultado de consulta
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repAlmacenamiento(valorAlm As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        'Consulta insertar nuevo dato
        query = "REPLACE INTO ddxVolAlmac (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numdia, cantEstac)" & _
                              " VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorAlm & "', '', '" & fecha & "', '', '" & valorAlm & "', '" & fecha & "', '', '" & valorAlm & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarAlm(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    resp = MsgBox("Seguro quiere eliminar el Almacenamiento de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
    If resp = vbOK Then
        query = "DELETE FROM ddxVolAlmac WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    End If
    'Fin de la conexión
    dbSIH.Close
End Sub

Function getVolAlm(idEstacion As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    getVolAlm = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue as Alm FROM ddxVolAlmac where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getVolAlm = Format(adoRs!alm, "0.000")
        Else
            getVolAlm = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getVolAlmGrp(idGrupo As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    iValG = 0
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT t1.station, t1.DailyValue as alm FROM ddxVolAlmac t1, stationGroups t2 where t2.stationgroup = '" & idGrupo & "' and t1.station = t2.station and t1.datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        Do While Not adoRs.EOF
            'Almacena valor de nivel
            mValGrp(iValG, 0) = adoRs!station
            mValGrp(iValG, 1) = Format(adoRs!alm, "0.000")
            iValG = iValG + 1
            adoRs.MoveNext
        Loop
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

''****************VALORES DE GASTO*********************
Sub addGasto(valorGas As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdxObratoma (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorGas & "', '', '" & fecha & "', '', '" & valorGas & "', '" & fecha & "', '', '" & valorGas & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        'Cierra resultado de consulta
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repGasto(valorGas As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        'Consulta insertar nuevo dato
        query = "REPLACE INTO ddxobratoma (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorGas & "', '', '" & fecha & "', '', '" & valorGas & "', '" & fecha & "', '', '" & valorGas & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarGasto(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    resp = MsgBox("Seguro quiere eliminar el Almacenamiento de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
    If resp = vbOK Then
        query = "DELETE FROM ddxobratoma WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    End If
    'Fin de la conexión
    dbSIH.Close
End Sub
Function getGasto(idEstacion As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    getGasto = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue as gasto FROM ddxObraToma where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getGasto = Format(adoRs!gasto, "0.000")
        Else
            getGasto = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getGastoGrp(idGrupo As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    iValG = 0
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT t1.station, t1.DailyValue as gasto FROM ddxObraToma t1, stationGroups t2 where t2.stationgroup = '" & idGrupo & "' and t1.station = t2.station and t1.datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        Do While Not adoRs.EOF
            'Almacena valor de nivel
            mValGrp(iValG, 0) = adoRs!station
            mValGrp(iValG, 1) = Format(adoRs!gasto, "0.000")
            iValG = iValG + 1
            adoRs.MoveNext
        Loop
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function


''****************VALORES DE LLUVIA*********************
Sub addLluvia(valorLlu As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdPrecipitacio (station, datee, valuee, msgcode, acumvalue, numdia, cantestac)" & _
                " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorLlu & "', '', '0', '" & Day(fecha) & "','0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repLluvia(valorLlu As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DdPrecipitacio (station, datee, valuee, msgcode, acumvalue, numdia, cantestac)" & _
                " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorLlu & "', '', '0', '" & Day(fecha) & "','0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarLluvia(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar la LLUVIA de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdPrecipitacio WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Cierra resultado de consulta
            'adoRs.Close
        End If
    
    'Fin de la conexión
    dbSIH.Close
End Sub

Function getLluvia(idEstacion As String, fecha As String) As String
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    
    '*********************************
    '   Consulta lluvia de estacion
    '*********************************
    getLluvia = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, valuee FROM ddprecipitacio where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            If adoRs!valuee = "0.01" Then
                getLluvia = "Inap"
            Else
                getLluvia = Format(adoRs!valuee, "0.0")
            End If
        Else
            getLluvia = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getLluviaGrp(idGrupo As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*******************************************
    '   Consulta Lluvia de un grupo de estacines
    '*******************************************
    iValG = 0
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT t1.station, t1.Valuee as lluvia FROM ddprecipitacio t1, stationGroups t2 where t2.stationgroup = '" & idGrupo & "' and t1.station = t2.station and t1.datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        Do While Not adoRs.EOF
            'Almacena valor de nivel
            mValGrp(iValG, 0) = adoRs!station
            mValGrp(iValG, 1) = Format(adoRs!lluvia, "0.000")
            iValG = iValG + 1
            adoRs.MoveNext
        Loop
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

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

''****************VALORES DE TEMPERATURA*********************
Sub addTemps(valAmb As String, valMax As String, valMin As String, tmpMed As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    Dim tmpMedia As Double
    
    tmpMedia = Round((CDbl(valMax) + CDbl(valMin)) / 2, 1)
    
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdxTempaire (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & CStr(tmpMedia) & "', '" & valAmb & "', '" & fecha & "', '', '" & valMax & "', '" & fecha & "', '', '" & valMin & "', '0', '0', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repTemps(valAmb As String, valMax As String, valMin As String, tmpMed As String, idEstacion As String, fecha As String)
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DdxTempaire (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                  "VALUES ('" & idEstacion & "', '" & fecha & "', '" & tmpMed & "', '" & valAmb & "', '" & fecha & "', '', '" & valMax & "', '" & fecha & "', '', '" & valMin & "', '0', '0', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarTemps(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        resp = MsgBox("Seguro quiere eliminar los datos de temperatura de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdxTempaire WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Cierra resultado de consulta
            'adoRs.Close
        End If
    
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub getTemp(idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    
    '*********************************
    '   Consulta Temperaturas de la estacion
    '*********************************
    temperaturas(0) = "" 'Promedio
    temperaturas(1) = "" 'Ambiente
    temperaturas(2) = "" 'Máxima
    temperaturas(3) = "" 'Mínima
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT DailyValue, MsgCode, MaxInstValue, MinInstValue FROM ddxTempAire where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            temperaturas(0) = Format(adoRs!DailyValue, "0.0") 'Promedio
            temperaturas(1) = Format(adoRs!MsgCode, "0.0") 'Ambiente
            temperaturas(2) = Format(adoRs!MaxInstValue, "0.0") 'Máxima
            temperaturas(3) = Format(adoRs!MinInstValue, "0.0") 'Mínima
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

''****************VALORES DE EVAPORACIÓN*********************
Sub addEvap(valorEvap As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdEvaporacion (Station, Datee, Valuee, MsgCode, acumValue, numDia, cantEstac)" & _
                                    " VALUES ('" & idEstacion & "','" & fecha & "','" & valorEvap & "','','0','" & Day(fecha) & "','0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repEvap(valorEvap As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DdEvaporacion (Station, Datee, Valuee, MsgCode, acumValue, numDia, cantEstac)" & _
                                    " VALUES ('" & idEstacion & "','" & fecha & "','" & valorEvap & "','','0','" & Day(fecha) & "','0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarEvap(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar la Evaporacion de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdEvaporacion WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
    'Fin de la conexión
    dbSIH.Close
End Sub

Function getEvaporacion(idEstacion As String, fecha As String) As String
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    
    '*********************************
    '   Consulta Evaporacion de la estacion
    '*********************************
    getEvaporacion = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, valuee FROM ddevaporacion where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getEvaporacion = Format(adoRs!valuee, "0.00")
        Else
            getEvaporacion = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getEvapGrp(idGrupo As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*******************************************
    '   Consulta Evaporación de un grupo de estacines
    '*******************************************
    iValG = 0
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT t1.station, t1.Valuee as evap FROM ddevaporacion t1, stationGroups t2 where t2.stationgroup = '" & idGrupo & "' and t1.station = t2.station and t1.datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        Do While Not adoRs.EOF
            'Almacena valor de nivel
            mValGrp(iValG, 0) = adoRs!station
            mValGrp(iValG, 1) = Format(adoRs!evap, "0.00")
            iValG = iValG + 1
            adoRs.MoveNext
        Loop
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

''****************VALORES DE PRESIÓN *********************
Sub addPresion(val As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        
        'Consulta insertar nuevo dato
        query = "INSERT INTO ddxPresBarometr (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & val & "', '', '" & fecha & "', '', '" & val & "', '" & fecha & "', '', '" & val & "', '0', '0', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repPresion(val As String, idEstacion As String, fecha As String)
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO ddxPresBarometr (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & val & "', '', '" & fecha & "', '', '" & val & "', '" & fecha & "', '', '" & val & "', '0', '0', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarPresion(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        resp = MsgBox("Seguro quiere eliminar la presion de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM ddxPresBarometr WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Cierra resultado de consulta
            'adoRs.Close
        End If
    
    'Fin de la conexión
    dbSIH.Close
End Sub

Function getPresion(idEstacion As String, fecha As String) As String
    Dim err As Boolean
    If dns = "" Then iniDataBase
    '*********************************
    '   Consulta PRESIÓN de la estacion
    '*********************************
    getPresion = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue FROM ddxPresBarometr where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getPresion = Format(adoRs!DailyValue, "0.0")
        Else
            getPresion = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

''****************VALORES DE HUMEDAD RELATIVA *********************
Sub addHumedad(val As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        
        'Consulta insertar nuevo dato
        query = "INSERT INTO ddxHumRelativa (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & val & "', '', '" & fecha & "', '', '" & val & "', '" & fecha & "', '', '" & val & "', '0', '0', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repHumedad(valHum As String, idEstacion As String, fecha As String)
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO ddxHumRelativa (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac) " & _
                                 "VALUES ('" & idEstacion & "', '" & fecha & "', '" & valHum & "', '', '" & fecha & "', '', '" & valHum & "', '" & fecha & "', '', '" & valHum & "', '0', '0', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        'Cierra resultado de consulta
        'adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarHumedad(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        resp = MsgBox("Seguro quiere eliminar la humedad de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM ddxHumRelativa WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
            'Cierra resultado de consulta
            'adoRs.Close
        End If
    
    'Fin de la conexión
    dbSIH.Close
End Sub
Function getHumedad(idEstacion As String, fecha As String) As String
    Dim err As Boolean
    If dns = "" Then iniDataBase
    '*********************************
    '   Consulta HUMEDAD RELATIVA de la estacion
    '*********************************
    getHumedad = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue FROM ddxHumRelativa where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getHumedad = Format(adoRs!DailyValue, "0")
        Else
            getHumedad = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function


''****************VALORES DE VERTEDOR*********************
Sub addVertedor(valorVert As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdxVertedor (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorVert & "', '', '" & fecha & "', '', '" & valorVert & "', '" & fecha & "', '', '" & valorVert & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repVertedor(valorVert As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DdxVertedor (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorVert & "', '', '" & fecha & "', '', '" & valorVert & "', '" & fecha & "', '', '" & valorVert & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarVertedor(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar valor Vertedor de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdxVertedor WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
    'Fin de la conexión
    dbSIH.Close
End Sub
Function getVertedor(idEstacion As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*********************************
    '   Consulta Vertedor de la estación
    '*********************************
    getVertedor = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue as vert FROM DdxVertedor where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getVertedor = Format(adoRs!vert, "0.000")
        Else
            getVertedor = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

''****************VALORES DE DERRAME*********************
Sub addDerrame(valDerrame As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO ddxDerrame (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valDerrame & "', '', '" & fecha & "', '', '" & valDerrame & "', '" & fecha & "', '', '" & valDerrame & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repDerrame(valDerrame As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO ddxDerrame (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valDerrame & "', '', '" & fecha & "', '', '" & valDerrame & "', '" & fecha & "', '', '" & valDerrame & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarDerrame(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar valor Vertedor de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM ddxDerrame WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
    'Fin de la conexión
    dbSIH.Close
End Sub

''****************VALORES DE O. T. Capturado en O.T.2 por estructura de la BD*********************
Sub addOT2(valorOt As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdxObraToma2 (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorOt & "', '', '" & fecha & "', '', '" & valorOt & "', '" & fecha & "', '', '" & valorOt & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repOT2(valorOt As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DdxObraToma2 (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorOt & "', '', '" & fecha & "', '', '" & valorOt & "', '" & fecha & "', '', '" & valorOt & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarOT2(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar valor O.T. de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdxObraToma2 WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
    'Fin de la conexión
    dbSIH.Close
End Sub

''****************VALORES Gasto en Río *********************
Sub addGastoRio(valorGastoR As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO DdxGastoenRio(Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES ('" & idEstacion & "', '" & fecha & "', '" & valorGastoR & "', '', '" & fecha & "', '', '" & valorGastoR & "', '" & fecha & "', '', '" & valorGastoR & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repGastoRio(valorGastoR As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        query = "REPLACE INTO DdxGastoenRio (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numDia, cantEstac)" & _
        " VALUES('" & idEstacion & "', '" & fecha & "', '" & valorGastoR & "', '', '" & fecha & "', '', '" & valorGastoR & "', '" & fecha & "', '', '" & valorGastoR & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminarGastoRio(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
        resp = MsgBox("Seguro quiere eliminar valor Gasto en " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
        If resp = vbOK Then
            query = "DELETE FROM DdxGastoenRio WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
            'MsgBox query
            adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        End If
    'Fin de la conexión
    dbSIH.Close
End Sub
Function getGasRio(idEstacion As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*********************************
    '   Consulta Gasto del rio
    '*********************************
    getGasRio = ""
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT station, DailyValue as gasto FROM DdxGastoenRio where station = '" & idEstacion & "' and datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            getGasRio = Format(adoRs!gasto, "0.000")
        Else
            getGasRio = ""
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getGasRioGrp(idGrupo As String, fecha As String) As String
    If dns = "" Then
        iniDataBase
    End If
    '*******************************************
    '   Consulta nivel de un grupo de estacines
    '*******************************************
    iValG = 0
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT t1.station, t1.DailyValue as gasto FROM DdxGastoenRio t1, stationGroups t2 where t2.stationgroup = '" & idGrupo & "' and t1.station = t2.station and t1.datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        Do While Not adoRs.EOF
            'Almacena valor de nivel
            mValGrp(iValG, 0) = adoRs!station
            mValGrp(iValG, 1) = Format(adoRs!gasto, "0.000")
            iValG = iValG + 1
            adoRs.MoveNext
        Loop
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

''****************VALORES DE ÁREA ELEVACIÓN*********************
Sub addArea(valArea As String, idEstacion As String, fecha As String)
    Dim err As Boolean
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "INSERT INTO ddxAreaPresa (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numdia, cantEstac) " & _
                              "VALUES ('" & idEstacion & "', '" & fecha & "', '" & valArea & "', '', '" & fecha & "', '', '" & valArea & "', '" & fecha & "', '', '" & valArea & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        'Cierra resultado de consulta
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub repArea(valArea As String, idEstacion As String, fecha As String)
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta remplazar dato
        'Consulta insertar nuevo dato
        query = "REPLACE INTO ddxAreaPresa (Station, Datee, DailyValue, MsgCode, MaxInstTime, MaxInstMsgCode, MaxInstValue, MinInstTime, MinInstMsgCode, MinInstValue, acumValue, numdia, cantEstac)" & _
                              " VALUES ('" & idEstacion & "', '" & fecha & "', '" & valArea & "', '', '" & fecha & "', '', '" & valArea & "', '" & fecha & "', '', '" & valArea & "', '0', '" & Day(fecha) & "', '0')"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
    'Fin de la conexión
    dbSIH.Close
End Sub

Sub eliminaArea(idEstacion As String, fecha As String)
    Dim resp As Integer
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    resp = MsgBox("Seguro quiere eliminar el Area de " & idEstacion & " con fecha: " & fecha, vbOKCancel + vbCritical)
    If resp = vbOK Then
        query = "DELETE FROM ddxAreaPresa WHERE station = '" & idEstacion & "' AND datee = '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    End If
    'Fin de la conexión
    dbSIH.Close
End Sub

Function getUltNiv(clvEstacion As String, fecha As String) As String
    'Variable para rango de fechas
    Dim d As String
    'Conexión
    If dns = "" Then iniDataBase
    
    'Rango de fecha para obtener el último nivel (2 días atras)
    d = Format(DateAdd("m", -1, fecha), "yyyy/mm/dd")
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta el último nivel de escala registrado al día
    query = "SELECT dailyValue AS val FROM ddxNivel WHERE station = '" & clvEstacion & _
            "' AND datee >= '" & d & "' AND datee < '" & Format(fecha, "yyyy/mm/dd hh:mm") & "' ORDER BY Datee DESC LIMIT 1"
            'MsgBox query
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
        If Not adoRs.EOF Then
            'Devuelve el resultado de la consulta
            If Not IsNull(adoRs!val) Then
                getUltNiv = adoRs!val
            End If
        End If
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close
End Function


Sub getDosUNivls(idEstacion As String, fecha As String)
    'Variable para rango de fechas
    Dim d As String
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    ReDim dosNiveles(1)
    'Rango de fecha para obtener el último nivel (2 días atras)
    d = Format(DateAdd("d", -10, fecha), "yyyy/mm/dd")
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    'Consulta el último nivel de escala registrado al día
    query = "SELECT dailyValue AS val FROM ddxNivel WHERE station = '" & idEstacion & _
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

Function hayDosNiveles(idEstacion As String, fecha As String) As Boolean
    'Variable para rango de fechas
    Dim d As String
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    'Rango de fecha para obtener el último nivel (2 días atras)
    d = Format(DateAdd("d", -10, fecha), "yyyy/mm/dd")
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    'Consulta el último nivel de escala registrado al día
    query = "SELECT count(*) AS n FROM ddxNivel WHERE station = '" & idEstacion & _
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

'*********************************************************
'   Consulta desviacion estandar del nivel de la estación
'*********************************************************
Function getDesviacionStd(clvEstacion As String, fecha As String) As String
    Dim fchB As String
    'Inicia variables en caso de no existir
    If dns = "" Then iniDataBase
    'Fecha de un mes anterior
    fchB = Format(DateAdd("m", -1, fecha), "yyyy/mm/dd")
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT STD(DailyValue) as desEs FROM ddxNivel WHERE station = '" & clvEstacion & "' AND datee >= '" & fchB & "' AND datee<= '" & fecha & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            If Not IsNull(adoRs!desEs) Then
                getDesviacionStd = adoRs!desEs
            Else
                getDesviacionStd = ""
            End If
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

'*********************************************************
'   Consulta desviacion estandar del nivel de la estación
'*********************************************************
Function getNamo(clvEstacion As String) As String
    'Inicia variables en caso de no existir
    If dns = "" Then iniDataBase
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
        'Consulta insertar nuevo dato
        query = "SELECT esccrit FROM stations WHERE station = '" & clvEstacion & "'"
        'MsgBox query
        adoRs.Open query, dbSIH, adOpenDynamic, adLockOptimistic
        If Not adoRs.EOF Then
            If Not IsNull(adoRs!esccrit) Then
                If adoRs!esccrit > 0 Then
                    getNamo = adoRs!esccrit
                Else
                    getNamo = ""
                End If
            Else
                getNamo = ""
            End If
        End If
        'Cierra resultado de consulta
        adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getDerAlmace(clv As String, niv As Double)
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta almacenamietos
    '        SELECT * FROM transformtables where xvalue= '57.47' and transformtableid = 'cdooxeleva'
    query = "select * from transformtables where xvalue= '" & niv & "' and transformtableid = '" & clv & "'"
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        getDerAlmace = Format(adoRs!Yvalue, "0.000")
    End If
    'Cierra resultado de consulta
    adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getDerNivel(clv As String, alm As Double)
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta almacenamietos
    '        SELECT * FROM transformtables where xvalue= '57.47' and transformtableid = 'cdooxeleva'
    query = "select * from transformtables where yValue= '" & alm & "' and transformtableid = '" & clv & "'"
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        getDerAlmace = Format(adoRs!xValue, "0.00")
    End If
    'Cierra resultado de consulta
    adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getDerArea(clv As String, niv As Double)
    'Conexión
    If dns = "" Then
        iniDataBase
    End If
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta almacenamietos
    '        SELECT * FROM transformtables where xvalue= '57.47' and transformtableid = 'cdooxeleva'
    query = "select * from transformtables where xvalue= '" & niv & "' and transformtableid = '" & clv & "'"
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        getDerArea = adoRs!Yvalue
    End If
    'Cierra resultado de consulta
    adoRs.Close
    'Fin de la conexión
    dbSIH.Close
End Function

Function getNombreEstacion(clv As String) As String
    'Prueba de conexión
    If dns = "" Then iniDataBase
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta nombre de la estación
    query = "select stationName from stations where station = '" & clv & "'"
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If Not IsNull(adoRs!stationName) Then
            getNombreEstacion = adoRs!stationName
        End If
    End If
    
    'Cierra resultado de consulta
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close
End Function

Function getNombreEstado(clv As String) As String
    'Prueba de conexión
    If dns = "" Then iniDataBase
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta nombre de la estación
    query = "select estado from stations where station = '" & clv & "'"
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If Not IsNull(adoRs!estado) Then
            Select Case adoRs!estado
                Case "VC"
                    getNombreEstado = "VER"
                Case "PB"
                    getNombreEstado = "PUE"
                Case "OX"
                    getNombreEstado = "OAX"
            End Select
        End If
    End If
    
    'Cierra resultado de consulta
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close
End Function

Function getNombreCuenca(clv As String) As String
    'Prueba de conexión
    If dns = "" Then iniDataBase
    
    'Conexión a la base de datos
    dbSIH.ConnectionString = dns
    dbSIH.Open
    
    'Consulta nombre de la estación
    query = "select t2.nomCuenca from stations t1, cuencas t2 where t1.station = '" & clv & "' and t1.cuenca = t2.cuenca"
    'Ejecuta la consulta
    adoRs.Open query, dbSIH, adOpenStatic, adLockReadOnly
    If Not adoRs.EOF Then
        If Not IsNull(adoRs!nomCuenca) Then
            getNombreCuenca = adoRs!nomCuenca
        End If
    End If
    
    'Cierra resultado de consulta
    adoRs.Close
    
    'Fin de la conexión
    dbSIH.Close
End Function

