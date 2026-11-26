<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  CantRegAMostrar = 1
CantRegAMover = 1
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
  OpcAgregar = request.querystring("OpcAgregar")

  set cn = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos request.querystring("seg"),""
  Session("Form") = request.querystring("seg")

  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

if request.form("EM") <> "" then
session("emisor") = request.form("em")
end if

if request.form("DP") <> "" then
session("destprin") = request.form("dp")
end if

if request.form("DA") <> "" then
session("destad") = request.form("da")
end if


X1 = getinfouser(session("emisor"))
X1MAIL = session("SGU_MAIL")
X1NOMBRE = session("SGU_NOMBRE")
X2 = getinfouser(session("destprin"))
X2MAIL = session("SGU_MAIL")
X2NOMBRE = session("SGU_NOMBRE")
X3 = getinfouser(session("destad"))
X3MAIL = session("SGU_MAIL")
X3NOMBRE = session("SGU_NOMBRE")

if request.querystring("borrar") <> "no" then
session("emisor") = ""
session("destprin") = ""
session("destad") = ""
end if


  sxopc1 = "<option value=' '>Ninguno</option>"
  sxopc2 = ""
  sxopc3 = ""

  strsql = "Select email,descripcion From Usuarios  where ID = 'COREMANAGER'"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
     smail = trim(rs(0))

     if len(smail) > 0 then
        sxopc2 = "<option value='" & smail & "' selected>" & rs(1) & "</option> "
     end if
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "Select email,descripcion From Usuarios  order by descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof
     smail = trim(rs(0))

     if len(smail) > 0 then
        sxopc3 = sxopc3 & "<option value='" & smail & "'>Usuario:" & rs(1) & "</option> "
     end if

  rs.movenext
  loop
  call ofv.cerrarconsulta(rs)

  sxopc2 = sxopc2 & sxopc3
  sxopc3 = sxopc1 & sxopc3

        sxopc4 = "<option value=0 >BAJA</option> "
        sxopc4 = sxopc4 & "<option value=1 selected >NORMAL</option> "
        sxopc4 = sxopc4 & "<option value=2 >ALTA</option> "

        sxopc5 = "<option value='NO' >NO</option> "
        sxopc5 = sxopc5 & "<option value='SI' >SI</option> "



    sHTML = sHTML & "<center>"

'sel2 = "<select class=at name=from >" & sxopc2 & "</select>"
'sel3 = "<select class=at name=toWho >" & sxopc3 & "</select>"
'sel4 = "<select class=at name=toWho2 >" & sxopc3 & "</select>"
sel5 = "<select class=at name=impo >" & sxopc4 & "</select>"
sel6 = "<select class=at name=aviso >" & sxopc5 & "</select>"
sel2 = "<INPUT TYPE='HIDDEN' VALUE='" & X1MAIL & "' NAME='from' >"
sel3 = "<INPUT TYPE='HIDDEN' VALUE='" & X2MAIL & "' NAME='towho' >"
sel4 = "<INPUT TYPE='HIDDEN' VALUE='" & X3MAIL & "' NAME='towho2' >"


    celdas = ofv.GenCelda("","ttt colspan=2","","","","","","Enviar Mail","si")
    filas  = ofv.GenRow("","","","","22","","",celdas,"si")


    celdas = ofv.GenCelda("","tt","","","","","","Emisor:","si")
    'sImagenP = "&#0232;"
	sinput = "<a href='AgregarLista.asp?op=EM' "
	sinput = sinput  & " title='Lista de autorizantes' ><i class='las la-arrow-right'></i></a>"

    celdas = celdas & ofv.GenCelda("","ut","","20","","","","&nbsp;" & X1NOMBRE & "&nbsp;" & "(" & X1MAIL & ")" & sel2 & sinput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas ,"si")


    celdas = ofv.GenCelda("","tt","","","","","","Destinatario Principal:","si")
    'sImagenP = "&#0232;"
	sinput = "<a href='AgregarLista.asp?op=DP' "
	sinput = sinput  & " title='Lista de autorizantes' ><i class='las la-arrow-right'></i></a>"

    celdas = celdas & ofv.gencelda("","ft","","","15%","","","&nbsp;" & X2NOMBRE & "&nbsp;" & "(" & X2MAIL & ")" & sel3 & sInput,"si")


    filas = filas & ofv.GenRow("","","","","10","","",celdas ,"si")

    celdas = ofv.GenCelda("","tt","","","","","","Destinatario Adicional:","si")
    'sImagenP = "&#0232;"
	sinput = "<a href='AgregarLista.asp?op=DA' "
	sinput = sinput  & " title='Lista de autorizantes' ><i class='las la-arrow-right'></i></a>"
    celdas = celdas & ofv.GenCelda("","ut","","20","","","","&nbsp;" & X3NOMBRE & "&nbsp;" & "(" & X3MAIL & ")" & sel4 & sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas ,"si")


    celdas = ofv.GenCelda("","tt","","","","","","Asunto:","si")

    celdas = celdas & ofv.GenCelda("","ut","","","","","","<input class=dt name=subject type=text size=80>","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ttr","","","","","","Opciones","si")
    celdas = celdas & ofv.GenCelda("","ttr","center","","","","","Importancia:&nbsp;" & sel5 & "&nbsp;&nbsp;Aviso de lectura:&nbsp;" & sel6,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt valign=top","","","","","","Mensaje:","si")

    celdas = celdas & ofv.GenCelda("","ut ","","","","","","<textarea style=border-radius:5px; name=message cols=80 rows=10></textarea>","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut colspan=2","","","","","","&nbsp;","si")
    filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt colspan=2 align=center","","","","","","<input class=bt type=submit value=Enviar >","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      sAccion = "EnvMail2.asp"
      formulario2 = ofv.GenForm("Usuarios3","",sAccion,"","post","",filas,"si")




    filas = ofv.GenRow("","","","","","","",formulario2,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","80%","","0","","0","0","","",filas,"si")

  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)



  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>