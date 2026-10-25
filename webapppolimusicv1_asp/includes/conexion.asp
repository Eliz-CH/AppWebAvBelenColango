<%
'==============================================================================
' Archivo:      conexion.asp
' Descripcion:  Cadena de conexion compartida a SQL Server (BDD_PoliMusic)
' Proyecto:     PoliMusic v1 - ASP Clasico (VBScript)
' Autor:        WorkBuddy AI
'==============================================================================
' Este archivo se incluye con <!--#include file="includes/conexion.asp"-->
' Expone la variable global: conn (ADODB.Connection ya abierta)
'==============================================================================

Option Explicit

' --- Configuracion de la cadena de conexion (SQLOLEDB) -----------------
Dim strConnectionString
strConnectionString = _
    "Provider=SQLOLEDB.1;" & _
    "Data Source=(local);" & _
    "Initial Catalog=BDD_PoliMusic;" & _
    "User ID=usr_polimusic;" & _
    "Password=Politecnica1;" & _
    "Network Library=DBMSSOCN;"

' --- Crear y abrir la conexion ADODB ------------------------------------
Dim conn
Set conn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
conn.Open strConnectionString
If Err.Number <> 0 Then
    On Error GoTo 0
    Response.Status = "500 Internal Server Error"
    Response.Write "<!DOCTYPE html><html><head><meta charset=""utf-8"">" & _
                   "<title>Error de conexion</title></head><body>" & _
                   "<h2>No se pudo conectar a la base de datos.</h2>" & _
                   "<p>Verifique que SQL Server este en ejecucion y que las " & _
                   "credenciales de <strong>usr_polimusic</strong> sean correctas.</p>" & _
                   "</body></html>"
    Response.End
End If
On Error GoTo 0
%>