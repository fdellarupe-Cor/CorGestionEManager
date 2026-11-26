<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
    response.expires=0






menuSel = request.querystring("menu")
menu    = menuSel
xmsg    = request.querystring("xmsg")

nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

   dFechai  = clng(date()) - nFecini


    set cn = ofv.conectar(ofv.strconn2)

    strsql = "Update novedades set fecexp = '" & dFechai & "' where codigo = '" & xmsg & "' "
 '''   response.write strsql
    call     ofv.crearconsultaEx(StrSql,cn,1,"")

    call ofv.cerrarconn(cn)


    response.redirect "NOVEDADM.asp?menu=" & menu

end if


%>
