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

     SKIND = SESSION("SKIND")

  sft = "&nbsp;<b>"
  sfta = "&nbsp;<b>"
  sfidc = "<font face=corflaticon size=2 >0</font>" 
  sfida = "<font face=corflaticon size=2 >1</font>" 
  sfitc = "<font face=corflaticon size=2 >0</font>" 
  sfita = "<font face=corflaticon size=2 >1</font>" 
  sftc = "</b>" 
  
  if sdocum = "" then
     sfuncion = ""
  else
     sfuncion = " onload=" & chr(34)
     sfuncion = sfuncion & " document.cons.submit(); "
     sfuncion = sfuncion & chr(34)
  end if

  sHTML = ofv.menuheader("",sfuncion)

  set cn = ofv.conectar(ofv.strconn2)

    Strsql = "SELECT descri, archivo, objid "
    Strsql = Strsql & " from ObjHtm Where codobjeto = " & nCodigo
    Strsql = Strsql & " order by secuen "
  Formulario = ""
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    do until rs.eof

    if sdocum = cstr(rs(2)) then
       sarch = rs(1)
       clase = "<font color=#993333><u>" 
       stexto = sfida & sfta & replace(rs(0)," ","&nbsp;") & sftc
    else 
       clase = "<font color=#000080><u>" 
       stexto = sfidc & sft & replace(rs(0)," ","&nbsp;") & sftc
    end if 
    stexto1 = "<A class=nt  href=CoArchivos.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(2)
    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"

    if rs(1) <> "" then
       Celda = ofv.GenCelda("","vt  ","","1%","","","",stexto1,"si")



       Formulario = Celda
    end if
    rs.movenext
    if not rs.eof then
    if sdocum = cstr(rs(2)) then
       sarch = rs(1)
       clase = "<font color=#993333><u>" 
       stexto = sfida & sfta & replace(rs(0)," ","&nbsp;") & sftc
    else 
       clase = "<font color=#000080><u>" 
       stexto = sfidc & sft & replace(rs(0)," ","&nbsp;") & sftc
    end if 
    stexto1 = "<A class=nt  href=CoArchivos.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
    stexto1 = stexto1 & "&sdocum=" & rs(2)
    stexto1 = stexto1 & "   >" & clase & stexto & "</u></font></a>"

    if rs(1) <> "" then
       Celda = ofv.GenCelda("","vt  ","","1%","","","",stexto1,"si")



       Formulario = Formulario & Celda
    end if
    rs.movenext
    end if
    Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")
    loop
  else
    sestr = "<h5>Sin&nbsp;datos&nbsp;disponibles"
    sestr = sestr & "</h5>" 
    Formulario = ofv.GenCelda("","et","","80%","","","",sestr,"si")
    Fila = ofv.GenRow("","","","","","","",Formulario,"si")
  end if


  sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","","","90%","","0","0","0","0","","",Fila,"si")
  if not sdocum = "" then  
     sAccion = "CoArchivos3.asp?Archivo=" & sarch
     sHTML = sHTML & ofv.GenForm("cons","",sAccion,"_blank","post","",sInput,"si")
  end if
  sHTML = sHTML & "</body></html>"
  response.write sHTML

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


end if



%>