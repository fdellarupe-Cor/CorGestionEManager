<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<!-- #INCLUDE FILE="../coress20/IncFile/CorStyles.inc" -->
<%





ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  Response.Write sHTML
else

call ofv.browser()



set oPrtG = new ClsReportes

Set oPrtG.Ofv = Ofv



sbloque1 =    oPrtG.Generar()



'response.write request.form("formato")
IF request.form("formato")="HTML" or request.form("formato")="" THEN

'oPrtG.shead=replace(oPrtG.shead,"&&","")



'response.write  oPrtG.srephead & oPrtG.shead


'sbloque1=replace(sbloque1,"'iconos/'" ,"'iconos/ICorAFin.jpg'")
'response.write  sbloque1 ' printer(sbloque1,corstyler2())
 response.write  printer(sbloque1,corstyler2(),oPrtG.srephead & oPrtG.shead)

end if
end if

Public Function printer( pagina,estilos,ENCABEZADO)



	SET objSoapClient = Server.CreateObject("MSSOAP.SoapClient30")
	objSoapClient.ClientProperty("ServerHTTPRequest") = True

	' needs to be updated with the url of your Web Service WSDL and is
	' followed by the Web Service name
	Call objSoapClient.mssoapinit("http://server-20/corv34des/ws/service1.asmx?WSDL")

	' use the SOAP object to call the Web Method Required
	printer = objSoapClient.Parser(pagina,estilos,ENCABEZADO)
End Function

%>