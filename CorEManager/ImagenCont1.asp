<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

  ID = request.querystring("ID")

  sHTML = ofv.MenuHeader(session("CMFMENU"),"")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From Complementos Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sNombre = rs(3)
    sCodint = rs(1)
    sTipcom = rs(2)
  end if 
  call ofv.cerrarconsulta(rs)

  sImaSrc = " "
  strsql = "Select * From Tipocomp Where tipo = '" & sTipcom & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sImaSrc = rs("archivo")
    sDescom = rs(1)
    if len(sDescom) > 1 then 
       sDescom = UCase(left(sDescom,1)) & LCase(right(sDescom,len(sDescom)-1))
    end if
  end if 
  call ofv.cerrarconsulta(rs)

  if sDescom = "" then sDescom = "Componente:"


'  sImagen = ofv.GenImage("","aaa title='Imprimir Componente' ","","","Imagenes/print.jpg","0","")
'  Fila = ofv.GenRow("","","","","","","","<a href=generarComp.asp?id="& ID &" target=_new >" & sImagen & "</a>","si")
'  sHTML = sHTML & ofv.GenTabla("","","","10%  ","","0","0","0","0","","",Fila,"si")

  Formulario = ""

    sImagen = ofv.GenImage("","","","","Imagenes/print.jpg","0","")
    sInput =  "<a href=generarComp.asp?id="& ID &" target=_new >" & sImagen & "</a>"
    Celda = ofv.GenCelda("","tt","","1%","","","",sInput,"si")

    sImagen = ofv.GenImage("","","","","Iconos/" & sImaSrc,"0","")
    sInput =  "<font color=#336699 size=2>"
    sInput = sInput & "&nbsp;" & sDescom & ": <b>" & sCodint & "&nbsp;-&nbsp;" & sNombre & "</b></font>"
    Celda = Celda &  ofv.GenCelda("","tt","","1%","","","",sImagen,"si")
    Celda = Celda &  ofv.GenCelda("","cc","","90%","","","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")

    Formulario = Formulario & Celda

    sInput = ofv.GenerarInput("","IMAGEN","submit","","10","bt","")
    Celda = ofv.GenCelda("","cc","","1%","","","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
    sAccion = "imagencont2.asp?ID=" & ID

    Formulario = Formulario & ofv.GenForm("Imagen1","",sAccion,"dato","post","",Celda,"si")

    sInput = ofv.GenerarInput("","ESTRUCTURA","submit","","10","bt","")
    Celda = ofv.GenCelda("","cc","","1%","","","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
    sAccion = "imagencont3.asp?ID=" & ID

    Formulario = Formulario & ofv.GenForm("Imagen2","",sAccion,"dato","post","",Celda,"si")

    sInput = ofv.GenerarInput("","ARCHIVOS","submit","","10","bt","")
    Celda = ofv.GenCelda("","cc","","1%","","","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
    sAccion = "imagencont4.asp?ID=" & ID

    Formulario = Formulario & ofv.GenForm("Imagen4","",sAccion,"dato","post","",Celda,"si")


  Fila = ofv.GenRow("","","","","","#f7eedd","",Formulario,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100%  ","","0","0","0","0","","",Fila,"si")

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>