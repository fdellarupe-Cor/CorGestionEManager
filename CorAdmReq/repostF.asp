
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%
Conset(session("conset"))
ofv.wstyles = CorStyler()
ofv.CorLogo = CorLogo()
'response.write CorLogo()
ofv.wbody = cstr("bgcolor=" & Session("APLBGD") & " style={border-style:ridge;border-width:4;} ")
ofv.CorFormHeader = CorFormHeader2(FormN,"S")
ofv.TitCls = "tt"


ok=ofv.CheckUsuario()

if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader(FormN)
  response.write shtml

   response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"
    AFiles=split(session("Sfilename"),"#zzzz#")
    ATam=split(session("Stamano"),"#zzzz#")
    for z=0 to ubound(afiles)-1
     Celda = ofv.gencelda("","tt","","","","","","Archivo copiado:","si")
      Celda = Celda & ofv.gencelda("","cc","","","","","","<b>" & afiles(z) & "</b>","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.gencelda("","tt","","","","","","Tamaño:","si")
      Celda = Celda & ofv.gencelda("","cc","","","","","","<b>" & atam(z) & " Kb</b>","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   next
    sHTML =  ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  response.write shtml

  response.write "</BODY></HTML>"
  end if

  %>