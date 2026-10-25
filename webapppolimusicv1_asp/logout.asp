<%@ Language="VBScript" CodePage="65001" %>
<%
'==============================================================================
' Archivo:      logout.asp
' Descripcion:  Cierra la sesion del usuario, destruye todas las variables
'               de sesion y redirige a login.asp
' Proyecto:     PoliMusic v1 - ASP Clasico (VBScript)
'==============================================================================

' --- Forzar expiracion de cache para evitar retroceso con sesion muerta ----
Response.CacheControl  = "no-cache"
Response.AddHeader     "Pragma", "no-cache"
Response.Expires       = -1

' --- Limpiar todas las variables de sesion ---------------------------------
Session.Contents.RemoveAll()

' --- Abandonar la sesion completa (destruye el objeto Session en servidor) -
Session.Abandon()

' --- Redirigir al login ----------------------------------------------------
Response.Redirect "login.asp"
Response.End
%>