<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegMover = cdbl(request.querystring("CantRegMover"))
  if CantRegMover > 0 then CantRegAMover = CantRegMover
  CantRegAMostrar = 10
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
  OpcAgregar = request.querystring("OpcAgregar")

  set cn = ofv.conectar(ofv.strconn0)
  set cn2 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  SECT = request.querystring("SECT")
  RESP = request.querystring("RESP")


    celdas = ofv.GenCelda("","tt COLSPAN=2","CENTER","","","","","AFECTACION DE USUARIOS","si")
    filas =  ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","tt","","","","","","Lista de usuarios","si")
    celdas = celdas & ofv.GenCelda("","tt","right","","","","","Pagina " & NPagina,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    StrSql = "Select id,Descripcion From Usuarios where sector = " & SECT & " Order By descripcion"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then rs.move(CantRegAMover)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegAMostrar)
    nID = rs(0)

       sHref = "PROYRESP.asp?OPCION=AU&RESP=" & RESP & "&Nusuario=" & rs(0)
       sLink = ofv.GenLink("","nt",sHref,"","","","<b><font size=2>" & rs(1) & "<font></b>","si")
       celdas = ofv.GenCelda("","vt","","80%","","","",sLink,"si")

''''       sHref = "PROYRESP.asp?OPCION=AU&Nusuario=" & rs(0)
       sLink = ofv.GenLink("","nt",sHref,"","","","<b>" & rs(0) & "</b>","si")
       celdas = celdas & ofv.GenCelda("","vt","","20%","","","",sLink,"si")

    filas =  filas & ofv.GenRow("","","","","10","","",celdas,"si")

  posi = posi + 1
  rs.movenext
  loop

       celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRESP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")
    filas =  filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sHTML = sHTML & "<BR>"
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:ridge;border-width:2;' ","","100%","","0","","-1","-1","","",filas,"si")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>