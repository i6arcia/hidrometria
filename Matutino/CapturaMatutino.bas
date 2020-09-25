Attribute VB_Name = "CapturaMatutino"
'Option explicit

'Variables para hojas de excel
Dim presas As Excel.Worksheet
Dim hidro As Excel.Worksheet
Dim norte As Excel.Worksheet
Dim sur As Excel.Worksheet

'Variables para conexion a la base de datos
Private dbSIH As New ADODB.Connection
Private adoRs As New ADODB.Recordset
Private query As String
Public dns As String

Sub titulos()
    Set presas = Worksheets("PRESAS")
    Set hidro = Worksheets("HIDROMETRICA")
    Set norte = Worksheets("No.1")
    Set sur = Worksheets("No.2")
    
    presas.Range("A5").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & _
                            Format(Now, "dd") & " de " & Format(Now, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
    hidro.Range("A5").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & _
                            Format(Now, "dd") & " de " & Format(Now, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
    norte.Range("A5").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & _
                            Format(Now, "dd") & " de " & Format(Now, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
    sur.Range("A5").Value = "Xalapa, Ver. -- " & Format(Now, "dddd") & " " & _
                            Format(Now, "dd") & " de " & Format(Now, "mmmm") & _
                            " de " & Format(Now, "yyyy") & " --"
End Sub

'###Obtener niveles###
Sub obtenerPresas()
    'Limpiar contenido
    presas.Range("E12:I52").ClearContents
    presas.Range("J12:K23").ClearContents
    presas.Range("J41:K48").ClearContents
End Sub


Sub capturaPresas()

'Variables




End Sub


