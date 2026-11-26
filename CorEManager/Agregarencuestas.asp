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

        dFechai = clng(date()) - nFecini
        dFechae = 0

        strsql = "Insert into emencuestas ("

        strsql = strsql & " ENCcodigo,ENCtipo,ENCfecinicio,ENCfecexp,"
        strsql = strsql & " ENCfecexpview,ENCmensaje,ENCimportancia,ENCtipodist,"
        strsql = strsql & " ENCrelacion,ENCcodobjeto,ENCtipobjeto,ENCworkflowid,"
        strsql = strsql & " ENCcontacto,ENChora,ENCOBSERVACION,ENCMODERADOR,"
        strsql = strsql & " ENCtitulo,ENCcopete,ENCimagen,ENCurl,"
        strsql = strsql & " ENCurldsc"

        strsql = strsql & " ) "
        strsql = strsql & " values (" 

        strsql = strsql & "'" & gtkey("901") & "','1','" & dFechai & "','" & dFechae & "',"
        strsql = strsql & "'" & dFechae & "','Mensaje de la encuesta','1','1',"
        strsql = strsql & "'0','0','','0',"
        strsql = strsql & "'" & Session("Usuario") & "','','N','',"
        strsql = strsql & "'Titulo de la encuesta','','0','',"
        strsql = strsql & "''"

        strsql = strsql & ")" 


        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

       response.redirect "adencuestas.asp"


  call ofv.cerrarconn(cn)

%>