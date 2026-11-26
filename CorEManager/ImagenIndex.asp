<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("ID")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From Complementos Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sNombre = rs(3)
    sCodint = rs(1)
    sTipcom = rs(2)
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "Select * From Tipocomp Where tipo = '" & sTipcom & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sDescom = rs(1)
    if len(sDescom) > 1 then
       sDescom = UCase(left(sDescom,1)) & LCase(right(sDescom,len(sDescom)-1))
    end if
  end if
  call ofv.cerrarconsulta(rs)

  if sDescom = "" then sDescom = "Complemento"

  strsql = "Select compasoc From Compcompuesto Where compmaster = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  k = 0

  'if not rs.eof then

  do until rs.eof
    if rs(0) <>"" then
    strsql = "Select * From Complementos Where codcomp = " & rs(0)
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
      redim preserve aNombre(k)
      redim preserve aID(k)
      redim preserve atipo(k)
      redim preserve adtipo(k)
      redim preserve aitipo(k)
      aNombre(k) = rs2(3)
      aID(k) = rs2(0)
      atipo(k) = rs2(2)
      adtipo(k) = rs2(2)
      aitipo(k) = ""
      k = k + 1
    end if
    call ofv.cerrarconsulta(rs2)
    end if
    rs.movenext
  loop
  call ofv.cerrarconsulta(rs)

  'if atipo <> "" then 'Or Ubound(atipo) <> "" then
if k <>0 then
  for x=0 to  Ubound(atipo)

  strsql = "Select descri,archivo From Tipocomp Where tipo = '" & atipo(x) & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    adtipo(x) = rs(0)
    aitipo(x) = rs(1)
   end if

  next
  end if

  'end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------


    sLink = Ucase(sDescom)
    Celda = ofv.GenCelda("","tt colspan=2","center","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sLink = "<b>" & sCodint & " - " & sNombre & "</b>"
    Celda = ofv.GenCelda("","cc colspan=2","","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","tt colspan=2","","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



 if k <> 0 then

  for k = lbound(aNombre) to ubound(aNombre)
    sImagen = ofv.GenImage("","","","","Iconos/" & aitipo(k),"0","")
  '''''''  sHref = "ImagenCont.asp?ID=" & aID(k)
    sHref = "EditarComp.asp?codcomp=" & aID(k) & "&tipocomp=" & atipo(k) & "&funcion=C&Complemento=" & ofv.convertircar(aNombre(k)," ","_")
    sLink = ofv.GenLink("","tt",sHref,"_new","","",sImagen,"si")
    Celda = ofv.GenCelda("","tt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"_new","","",aNombre(k),"si")
    Celda = Celda & ofv.GenCelda("","cc","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  next
  'else

  end if
  sHTML = sHTML & ofv.GenTabla("","aaa","","100%","","","","","","","",Fila,"si")
  sHTML = sHTML & "</center>"

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>