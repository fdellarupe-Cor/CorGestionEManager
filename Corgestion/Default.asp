<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<html>

<head>
<style>
</style>
<title>COR GESTION</title>
</head>

<body>

<%
Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

XXX = SYSLOGON(cnn,"LOGON")

cnn.close
%>

<FORM name='init' method='post'  ACTION='index.asp' >
<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >
</FORM>

</body>
<script language="javascript" >
   document.all.init.submit();
</script>
</html>