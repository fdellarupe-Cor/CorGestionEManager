<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

  response.expires=0

  Tabla = request.querystring("Tabla")
  vuelta = request.querystring("vuelta")
  Regisvuelta = request.querystring("Regisvuelta")

  set cn = ofv.conectar(ofv.strconn2)

nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

        'xmsgusr = "Texto del Mensaje"
        dFechai = clng(date()) - nFecini
        dFechae = 0



	strsql = "Insert into novegenpar ("
	strsql = strsql & " fecvig,fecexp,dias,texto,"
	strsql = strsql & " importancia,tipodist,contacto,enviado,"
	strsql = strsql & " MSGVERSION,MSGTITULO,MSGCOPETE,MSGIMAGEN,"
	strsql = strsql & " MSGURL,MSGURLDSC,"
	strsql = strsql & " MSGVAENC,MSGVAFORO,MSGCODENC,MSGCODFORO,"
	strsql = strsql & " MSGTIPOMESS,MSGAPLIC,MSGOBJCOD,MSGOBJTIP,CODEMP"
	strsql = strsql & " ) "
	strsql = strsql & " values ("
	strsql = strsql & dFechai & "," & dFechae & ",1,'',"
	
		If Session("EMPMUN")  = "S" Then	'Si es Multiempresa
			strsql = strsql & "1,2,'" & Session("Usuario") & "','N',"	'Toma en el campo tipodist el valor 2 (LISTA)
			Else	'Si no es Multiempresa
			strsql = strsql & "1,1,'" & Session("Usuario") & "','N',"	'Se modifica el valor que se inserta en el campo tipodist para que por defecto sea 1 (TODOS)
		End If	
		
	strsql = strsql & "2,'','','0',"
	strsql = strsql & "'1','',"
	strsql = strsql & "'N','N','0','0',"
	strsql = strsql & "'1','GLB','0','','" & session("un") & "'"
	strsql = strsql & ")"
response.write strsql
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        sqlz = "select objid from novegenpar order by objid desc"
        set rs = ofv.crearconsultaEx(sqlz,cn,1,parametros)

		codigo = ""
        if not rs.eof then codigo = rs(0)
        call ofv.cerrarconsulta(rs)

        sqlx = "insert into novegenpardist(codigo,usugrp,sector,fecexp) values ("
		sqlx = sqlx & "'" & codigo & "','TODOS','" & session("un") & "','0')"
		call ofv.crearconsultaEx(sqlx,cn,1,parametros)

       response.redirect "adnovedades.asp"


  call ofv.cerrarconn(cn)

%>