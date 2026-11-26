<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  server.scripttimeout = 600
  response.buffer = true



''''''''''''''



    sID = Request.form("textot")
	iF SID = "" THEN sID = Request.form("BUSCAR")
	
'RESPONSE.WRITE "DATO "&SID&" "&rEQUEST.FORM

   Session("Form")="CLIMS157172171124281051096140"

   shtml ="<form name='portexto' style='display:none;'  action=listadoconsultas.asp?Menu=PORTEXTO&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=TEXTO method=post>"
   shtml = shtml & "<input type=text name=Tcons value='PORTEXTO'>"
   shtml = shtml & "<input type=text name=tgrp value='T'>"
   shtml = shtml & "<input type=text name=tcod value='99999'>"
   shtml = shtml & "<input type=text name=tdoc value='99999'>"
   shtml = shtml & "<input type=text name=lista value='99999'>"
   shtml = shtml & "<input type=text name=tsel value='C'>"
   shtml = shtml & "<input type=text name=ttip value='TODOS'>"
   shtml = shtml & "<input type=text name=tnam value='TODOS'>"
   shtml = shtml & "<input type=text name=buscar value='"&sid&"'>"


   shtml =shtml & "</form>"
   shtml=shtml & "<script language='javascript' > "
   shtml=shtml & " document.portexto.submit();"
   shtml=shtml & "</script>"

'Tcons : PORTEXTO --<br>
'tgrp : T --<br>
'tcod : 99999 --<br>
' tdoc : 99999 --<br>
' lista : 99999 --<br>
' tsel : C --<br>
' ttip : TODOS --<br>
' tnam : TODOS --<br>
' buscar : caja --<br>


'    response.redirect "consBuscadorWeb.asp?texto=" & SID



'''''''''''''

end if

response.write sHTML





%>
