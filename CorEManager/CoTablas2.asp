<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sdocum = request.querystring("sdocum")

  sft = "<font face=tahoma,arial size=1 color=#336699>&nbsp;<b>"
  sfta = "<font face=tahoma,arial size=1 color=#993333>&nbsp;<b>"
  sfidc = "<font face=corflaticon size=2 color=#336699>0</font>" 
  sfida = "<font face=corflaticon size=2 color=#993333>1</font>" 
  sfitc = "<font face=corflaticon size=2 color=#336699>0</font>" 
  sfita = "<font face=corflaticon size=2 color=#993333>1</font>" 
  sftc = "</b></font>" 
  
  if sdocum = "" then
     sfuncion = ""
  else
     sfuncion = " onload=" & chr(34)
     sfuncion = sfuncion & " document.cons.submit(); "
     sfuncion = sfuncion & chr(34)
  end if

  sHTML = ofv.menuheader("#ffffff",sfuncion)

  set cn = ofv.conectar(ofv.strconn2)

    Strsql = "SELECT * "
    Strsql = Strsql & " from ObjTabDef Where codobjeto = " & nCodigo
    Strsql = Strsql & " order by tabsecuen "
  Formulario = ""
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof
    if sdocum = cstr(rs(0)) then
       sarch = rs(0)
       stexto = sfida & sfta & replace(rs(4)," ","&nbsp;") & sftc
    else 
       stexto = sfidc & sft & replace(rs(4)," ","&nbsp;") & sftc
    end if 
    stexto1 = "<A href=CoTablas2.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(0)
    stexto1 = stexto1 & "   >" & stexto & "</a>"


    if rs(3) <> "" then
    Celda = ofv.GenCelda("","cc","","1%","","","",stexto1,"si")
    Celda = Celda & ofv.GenCelda("","","","2","","","","&nbsp;","si")

    Formulario = Formulario & Celda
    end if
    rs.movenext
    loop
  else
    sestr = "<font color=#336699><h5>Sin&nbsp;datos&nbsp;disponibles"
    sestr = sestr & "</h5></font>" 
    Formulario = ofv.GenCelda("","dt","","80%","","","",sestr,"si")
  end if

  Fila = ofv.GenRow("","","","","","","",Formulario,"si")
  sHTML = sHTML & ofv.GenTabla("","","","5%","","0","0","0","0","","",Fila,"si")
  if not sdocum = "" then  
     sAccion = "CoTablas3.asp?Tabla=" & sarch
     sHTML = sHTML & ofv.GenForm("cons","",sAccion,"arch","post","",sInput,"si")
  end if
  sHTML = sHTML & "</body></html>"
  response.write sHTML

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


end if



%>