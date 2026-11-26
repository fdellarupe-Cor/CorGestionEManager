<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<!-- #INCLUDE FILE="IncFile/reportes.inc" -->

<!-- #INCLUDE FILE="../coress20/IncFile/CorStyles.inc" -->
<%

Session("usuario") = "GOOGLE"
Session("NOAI") = "S"
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  Response.Write sHTML
else


aplcheck()


set oPrtG = new ClsReportes

Set oPrtG.Ofv = Ofv



sbloque1 =     oPrtG.Generar()



response.write "<html><body>"

response.write corstyler()


response.write sbloque1




response.write " </body></html>"



END IF

%>

