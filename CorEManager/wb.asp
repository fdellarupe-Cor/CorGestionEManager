<%
sRBAC="http://192.168.1.150/ws/service1.asmx"
CC=getinfoUserADBCRED("")

FUNCTION getinfoUserADBCRED(legajo)

 Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

 Dirc=Dirc  & "Coressdll\INFO.xml"

 set xmlHttp2=SERVER.createobject("MSXML2.XMLHTTP")

    xml4=leerArc(dirc)
  '  xml4=REPLACE(XML4,"ZZZZ","<P><TABLE><TR><TD>SSSSSSSSSSSSSSSSSS</TD></TR></TABLE></P>")

 ' RESPONSE.WRITE XML4

    xmlHttp2.open "POST", sRBAC, FALSE

  ' strSoapAction = "http://tempuri.org/Parser"


    xmlHttp2.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
  ' xmlHttp2.setRequestHeader "SOAPAction", SoapActionUrl


    xmlHttp2.send (xml4 )

    set xml3 =xmlHTTP2.responsexml

  '  Set xmldoc=xml3.selectSingleNode("//ParserResult")

    response.write xmldoc.text
  set xmlHTTP2=nothing
  'set xml4=nothing
  set xmldoc=nothing
  set xml3=nothing

END FUNCTION



function leerArc(archivo)

'response.write archivo
Set FS = Server.CreateObject("Scripting.FileSystemObject")
Set Fichero = FS.OpenTextFile(archivo, 1)

leerArc=Fichero.readall

Fichero.Close
Set Fichero = Nothing

Set FS = Nothing

end function

%>