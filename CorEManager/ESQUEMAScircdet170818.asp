<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar


  nCodigo = "0"
  sTipo   = request.querystring("Tipobjeto")
  sVolver = request.querystring("volver")
   ' incidente nro 1374
  sCod    = "0" 'request.querystring("cod")
  sacc   = request.querystring("acc")
  sCIRCUITO    = request.querystring("CIRCUITO")


  vuelta  = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")

  if sCIRCUITO<> "" then
   session("sCIRCUITO")=sCIRCUITO
  else
   sCIRCUITO=session("sCIRCUITO")
  end if


  sHTML = ofv.FormHeader(session("FormN"))



  set cn = ofv.conectar(ofv.strconn2)
  set cx = ofv.conectar(ofv.strconn5)
   if sacc="updX" then
   usr=request.form("grpusr")
   usr= mid(usr,5,len(usr))
   sql= "update TIPOBJETO set wkfapro='"& usr &"' where tipobjeto='" &sTipo & "'"
    set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
   end if

 sPerm="Comentarios"

  sInput = "<big>CIRCUITO DE AUTORIZACION</big>"
  if sTipo = "SN" then
	Celda = ofv.gencelda("","ttt colspan=6","center","2%","","","",sInput,"si")
  else
  	Celda = ofv.gencelda("","ttt colspan=3","center","2%","","","",sInput,"si")
  end if
  Fila = fila & ofv.GenRow("","","","","20","","",Celda,"si")

	 if sTipo = "SN" then
		Celda = ofv.gencelda("","ttt colspan=6","center","2%","","","","LISTA&nbsp;DE&nbsp;AUTORIZANTES&nbsp;","si")
	  else
		Celda = ofv.gencelda("","ttt colspan=3","center","2%","","","","LISTA&nbsp;DE&nbsp;AUTORIZANTES&nbsp;","si")
  end if

  Fila = fila & ofv.GenRow("","","","","20","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","ttr colspan=2","","2%","","","","ORDEN  DESCRIPCION","si")

  if sTipo = "SN" then
	Celda = Celda & ofv.gencelda("","ttr colspan=2","","2%","","","","PASO","si")
	Celda = Celda & ofv.gencelda("","ttr colspan=2","","2%","","","","ESTADO","si")
  end if
  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

  xSql = "select tipocpt,wkfauto from tipobjeto where tipobjeto='" & sTipo & "'"
  set rj = ofv.crearconsultaEx(xSql,cn,1,parametros)

  if not rj.eof then

  	sTipoCpt = rj("tipocpt")
  	sWkfCod = rj("wkfauto")

  end if

if sTipo="PDCORXYW" then




else


   IF session("modalidad") = "PRGCOR" THEN

       Strsql = "SELECT OBJID, * FROM ORRSKCIRCUITOSDKDET  "
       Strsql = Strsql & " Where CODRSK=" & sCIRCUITO
       Strsql = Strsql & " And tiposeg = '" & sCod & "'"

      WKSPTABLA = "ORRSKCIRCUITOSDKDET"


   ELSE
      IF sTipo="PDCORXY" THEN
      Strsql = "SELECT * FROM pedidoSDKdet  "
        Strsql = Strsql & " Where codobjeto=0"
       Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
       Strsql = Strsql & " And tiposeg ='0' order by objid "
	    'Strsql = Strsql & " And tiposeg ='9' order by objid "

      WKSPTABLA = "pedidoSDKdet"


      ELSEIF sTipo="SN" THEN
      Strsql = "SELECT * FROM pedidoSDKdet  "
        Strsql = Strsql & " Where codobjeto=0"
       Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
       Strsql = Strsql & " order by objid "
	    'Strsql = Strsql & " And tiposeg ='9' order by objid "

      WKSPTABLA = "pedidoSDKdet"
      'response.write strsql
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if rs.eof then

			sql = "select b.codigo from WKFPSCRESP a "
			sql = sql & " inner join WKFRESP b on (a.codresp=b.id) "
			sql = sql & " where a.wkfcod='" & sWkfCod & "' "
			'response.write sql
	  		set rx = ofv.crearconsultaEx(sql,cx,1,parametros)

			do while not rx.eof

				ssql = "insert into pedidoSDKdet (codobjeto,tipobjeto,tiposeg,grpusr) values('0','" & sTipo & "','" & rx("codigo") & "','NO')"
				call ofv.crearconsultaEx(ssql,cn,1,parametros)

			rx.movenext
			loop

	  end if

	   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      ELSE



       Strsql = "SELECT * FROM OBJSDKdet  "
       Strsql = Strsql & " Where codobjeto=0"
       Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
        ' incidente nro 1374
       Strsql = Strsql & " And tiposeg ='0'"

      WKSPTABLA = "OBJSDKdet"

  END IF
' response.write strsql & "<br>"
end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

    sForm = "Sdk" & rs(0)

    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver="
      sHref = sHref  & vuelta & "?Cod=" & sCod & "&CIRCUITO=" & SCIRCUITO
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Permiso",sHref)
    else
      Celda = ofv.GenCelda("","a","","2%","","","","&nbsp;","si")
    end if

	if rs(4) <> "NO" then
    	sUbic  = instr(rs(4),"USR-")
    	sBusca = mid(rs(4),5,len(rs(4))-4)
    	sGrpUsr = "S/D"
    else
    	sBusca = ""
	end if
    if not sUbic > 0 then
       sTabla = "Grupo"

       set cx = ofv.conectar(ofv.strconn0)
       Strsql = "SELECT * FROM " & sTabla & "s"
       Strsql = Strsql & " Where ID = '" & sBusca & "' "
     '  response.write strsql & "<br>"
       set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       if not rx.eof then sGrpUsr = sTabla & ": " & rx(0) & " - " & rx(1)
       call ofv.cerrarconsulta(rx)
       call ofv.cerrarconn(cx)


    else
       sTabla = "Usuario"

        xxxx     = getInfoUser(sBusca)
        if sBusca & "" = "" then
        	sGrpUsr = "S/D"
        else
        	sGrpUsr = session("SGU_nombre")
        	sGrpUsr = sTabla & ":&nbsp;" & sBusca & " - " & sGrpUsr
        end if

    end if


	if sTipo = "SN" then

		sImagenP ="&#0232;"
		sAccion = "AgregarCircuitoESQUEMA.asp?Codobjeto=" & nCodigo & "&CIRCUITO=" & SCIRCUITO
  		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&tiposeg=" & rs("tiposeg") & "&volver=" & vuelta
		sInputI = "<a class='ft1n' href='" & sAccion & "'"
		sInputI = sInputI  & " title='Agregar aprobador' >" & sImagenP & "</a>"
		if sGrpUsr & "" = "" then sGrpUsr = "S/D"
    	sinput = ofv.gencelda("","ttt colspan=2","","60%","","","",sGrpUsr & "&nbsp;&nbsp;&nbsp;" & sInputI,"si")

    else
    	sinput = ofv.gencelda("","ttt colspan=2","","60%","","","",sGrpUsr,"si")
	end if

    Celda = Celda & sinput

    if sTipo = "SN" then
		Celda = Celda & ofv.gencelda("","ttt colspan=2","","2%","","","",rs("tiposeg"),"si")
		sEst = ""
		if rs("tiposeg") = "RESPONSABLE" then
			sEst = "PARA ASIGNAR"
		elseif rs("tiposeg") = "GERENTE" then
			sEst = "ASIGNADO Y EN EJECUCION"
		end if
		Celda = Celda & ofv.gencelda("","ttt colspan=2","","2%","","","",sEst,"si")
  	end if

    Fila  = fila & ofv.GenRow("","","","","20","","",Celda,"si")

   rs.movenext
   loop
end if

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","","","90% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"

 if bModificar and sTipo <> "SN" then
  sAccion = "AgregarCircuitoESQUEMA.asp?Codobjeto=" & nCodigo & "&CIRCUITO=" & SCIRCUITO
  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
  Agregar = ofv.BotonAgregar(sAccion)
END IF
  sAccion  = "TIPOSDEESQUEMAPROP.asp?" & REQUEST.QUERYSTRING & "&wPerfil=" &sTipo
  Retornar = ofv.BotonVolver(sAccion)





  sHTML = sHTML & "</center>"



'--------------------------------------------------------------------------------
sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>