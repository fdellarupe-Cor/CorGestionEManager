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
  sCalificador = request.querystring("Calificador")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")
  sUN = request.querystring("UN")
opc = request.querystring("opc")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Sector")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

'response.write request.querystring




    XMFTGRP = split(session("EmpHbl"),chr(9))

	delim=""
	id=""
	sImagenP = "&#0232;"

	for i=0 to ubound(XMFTGRP)

	id = id & delim & XMFTGRP(i)
	delim=","

	next

	celdax = ofv.GenCelda("","tt","","15%","","","","ID","si")
	celdax = celdax & ofv.GenCelda("","tt","","40%","","","","SITES","si")
	celdax = celdax & ofv.GenCelda("","tt","","15%","","","","SECTORES","si")
	filax = ofv.GenRow("","","","","","","",Celdax,"si")

	sql = "select id,cliente from clienteplan where id in (" & id & ")"
	'response.write sql
	set rx = ofv.crearconsultaEx(sql,cn,1,parametros)

	if not rx.eof then

	do while not rx.eof

	celdax = ofv.GenCelda("","tt","","15%","","","",rx(0),"si")
	celdax = celdax & ofv.GenCelda("","tt","","40%","","","",rx(1),"si")

	sImagenP = "&#0232;"
	sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
	sinput = "<a class='ft1n' " & sFuncAnc & " href='AgregarSectorEN.asp?tabla=objsector&UN=" & rx(0) & "&codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&calificador=O&opc=si&volver=" & volver
	sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
	sinput = sinput  & "' title='Areas/sectores de " & rx(1) & "' >" & sImagenP & "</a>"
	celdax = celdax & ofv.GenCelda("","tt","","15%","","","",sInput,"si")

	filax = filax & ofv.GenRow("","","","","","","",Celdax,"si")

	rx.movenext
	loop
	end if

	tablax= ofv.GenTabla("","","","70%","","0","","0","0","","",filax,"si")


sHTML = sHTML & tablax & "<br>" & "<br>" & "<br>" & "<br>"



if opc="si" then

sForm = "Agregarb"

sql = "select cliente from clienteplan where id=" & sUN
Set rc = ofv.crearconsultaEx(sql, cn, 1, parametros)

    sInput = "AGREGAR DATOS DE " & rc(0)
     call ofv.cerrarconsulta(rc)
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","SECTOR","si")
     'sqleco = "Select * From eco where estado=3 Order By descri"

      if session("empmun") = "S" then


	        sqleco = "Select a.*,b.cliente From Eco a "
	        sqleco = sqleco & " inner join clienteplan b on (a.ecoUN=b.Id) "
	        sqleco = sqleco & " where a.ecoun=" & sUN &  " and a.estado=3 "
	        sqleco = sqleco & " Order By b.cliente,a.descri "

	  else
	       sqleco = "Select * From Eco where eco.estado=3  Order By descri"
    end if

'  response.write sqleco
      Set rsSectorn = ofv.crearconsultaEx(sqleco, cn, 1, parametros)
        'sInput = ofv.GenerarCombo("edProcedimientos",rs(6).name,rs(6).value,Strsql,cn,0,1,"sb")

            scomK = ""'"null" & Chr(9)
        scomD = ""
        SCOMSEP =""
      do while not rsSectorn.eof
         scomK=scomK & SCOMSEP & rsSectorn(0)
         scomd=scomd & SCOMSEP & rsSectorn(11) & "&nbsp;-&nbsp;" & rsSectorn(1)
         SCOMSEP = chr(9)
         rsSectorn.movenext
        loop
            sInput = ofv.gencomboFX("Agregarb",rs(4).name,"",scomK,scomD,"")

  sAccion  = "asp/AgregarEN.asp?Tabla=" & Tabla
  sAccion  = sAccion & "&volver=" & volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sAccion  = sAccion & "&Calificador='" & sCalificador & "'"
  Formulario = ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,sInput,"si")

     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",Formulario,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = ""


  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  'Botones Aceptar y Cancelar
  sAccion = "edSectores.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sAccion = sAccion & "&Calificador=" & sCalificador
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")
end if
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"



  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>