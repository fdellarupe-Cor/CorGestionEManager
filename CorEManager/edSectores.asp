<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
CONT =0

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")



  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
     Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  sHTML = ofv.FormHeader(session("FormN"))

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

'----------------------------------------------

 shtml = shtml & CorPnlFnc("N","")


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

%>
<!-- #INCLUDE FILE="IncFile/CorVinculadorPL.inc" -->
<%

  IF WKSIGUE = "S" THEN

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  if sTipo = "CON" or sTipo = "ESQ" Then
    sCalificador = "R"
    sSector = "Sectores responsables"
  else
    sCalificador = "O"
    sSector = "Sectores responsables"
  end if

  'Encabezado sector

  Formularioz = ofv.Encabezados(sSector)

 '' sHTML = sHTML & Formulario
  xformulario=""


 ' EMPR = EMPRESAS()
'response.write EMPR
'  EMP = SPLIT(EMPRESAS, CHR(9))

'for i=0 to ubound(emp)
'response.write emp(i) & "<br>"
'next


IF session("EMPMUN") = "S" then

  strsql = "SELECT a.*, c.descri,b.cliente FROM  OBJSECTOR a "
  strsql = strsql & " inner join ECO c on (a.codsector=c.eco) "
  strsql = strsql & " inner join clienteplan b on (b.ID=c.ecoUN) "
  strsql = strsql & " where a.codsector = c.eco "
  strsql = strsql & " and c.estado=3 and a.codobjeto=" & nCodigo
  strsql = strsql & " And (calificador = 'O' or calificador = 'C' or calificador = 'O') "
  strsql = strsql & " order by b.cliente"
else

   Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
    Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and codobjeto=" & nCodigo
  Strsql = Strsql & " And (calificador = '" & sCalificador & "' or calificador = 'C' or calificador = 'O') "
   Strsql = Strsql & " order by eco.descri"
end if
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
  CONT= CONT +1
    sID = "edSectores1" & rs(0)

    'Eliminar
    if bModificar  and (rs(12) & "") <> "PL" then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objsector&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
''    Celda = ofv.BotonEliminar("Sector",sHref)
      Celda = ofv.BotonEliminar("Eliminacion",sHref)
      Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
    else
      Celda = ofv.GenCelda("","","","","","",""," ","si")
    end if

    'Sectores origen
    if bModificar  and (rs(12) & "") <> "PL" then

     if session("APLMDCLI")="BC" then
   '  response.write StrSql
        Strsql = "Select * From Eco where eco.estado=3 and ecoun=1 Order By descri"
	     set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	      scomK = ""'"null" & Chr(9)
	      scomD = "" '"" & Chr(9)
	      SCOMSEP =""

	      do while not rsSector.eof
	      CONT= CONT+1
	       scomK=scomK & SCOMSEP & rsSector(0)
	       scomd=scomd & SCOMSEP & rsSector(1)
	       SCOMSEP = chr(9)
	       rsSector.movenext
	    loop

	      sInput =  ofv.gencomboFX(sID,rs(4).name,rs(4).value, scomK, scomD, "")


    else
     if session("EMPMUN") = "S" then
	     sinput=rs(14) & "&nbsp;-&nbsp;" & rs(13)
	     else
	        sInput =  rs("descri")
    end if
    end if
  '    sInput = replace(sInput,"CLASS=TBCM ","CLASS=TBCM WIDTH=80% ")
    else
      if cstr(rs(4)) <> "" then
        Strsql = "Select descri From Eco Where eco = " & rs(4)

        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
        sOrigen = ofv.GenerarInput("",sOrigen,"readonly","","90","dt","")
      end if
      sInput = sOrigen
    end if
   ' Celda = Celda & ofv.GenCelda("","vti","","90%","","","",sInput ,"si")
 '   Fila = ofv.GenRow("","","","","","","",Celda,"si")


    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objsector&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      'tabla= ofv.GenTabla("","","","70%","","0","","0","0","","",Formularioz &fila,"si")
      xformulario=  ofv.GenForm(sID,"",sAccion,"","post","",sInput,"si")
      'Formulario = Formulario & ofv.GenForm(sID,"",sAccion,"","post","",Fila,"si")
    else
    ''  tabla= ofv.GenTabla("","","","70%","","0","","0","0","","",Formularioz &fila,"si")
      xformulario=  ofv.GenForm("","","","","post","",sInput,"si")
      'Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
    end if

    Celda = Celda & ofv.GenCelda("","vti","","90%","","","",xformulario ,"si")

     wComentario = ""
	 strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codinx2 = " & rs("objid") & " and "
     strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"
      strsql = strsql & " and codsol = " & session("codSol")
     strsql = strsql & " order by codcircuito, codusr, objid"
	'   response.write strsql
	      set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	      if not rss.eof then
	         wComentario = rss(13)
	      end if
      call ofv.cerrarconsulta(rss)
     if (Nestado=5  or Nestado=6)  and bModificar then

	      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
	      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)  & "&ID2=" & rs(0)

	      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
	      	       sHref   = sAccion

	''''''agregado para color del cuadrito de Observaciones version 4.6

	               if wComentario = "" then
	      	          sInput  =""' "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
	           '''''''sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

	       	       else
	       	          sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
	       	       end if
	  ''''''



	                'se cambio link por celda'
	              Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")

	      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")



	               'Codigo
	              ' Celda  = ofv.GenCelda("","tt","","20%","","","","Observaciones","si")
	               Celda = Celda & Celdaobs
	              ' Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

      end if

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
  loop

''  sHTML = sHTML & "<br><br><br><br><center>"
  'sHTML = sHTML & ofv.GenTabla("","","","70%","","0","","0","0","","",Formulario,"si")
	if CONT = 0 then
	    sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes sectores agregados.</div>"
	else
	    tabla = ofv.GenTabla("","aaa","","70%","","0","","0","0","","",formularioz & fila,"si")
	    sHTML = sHTML & tabla
end if
'  tabla= ofv.GenTabla("","aaa","","70%","","0","","0","0","","",formularioz & fila,"si")
'  sHTML = sHTML & tabla


end if

  'Agregar
  if bModificar then
    sAccion = "AgregarSectorEN.asp?Tabla=Objsector&volver=" & vuelta & "&Codobjeto="
    sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Calificador='" & sCalificador & "'"
    Agregar = ofv.BotonAgregar(sAccion)

  end if


  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if




'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","") & GTCOBS()

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML &  "</body></html>"
end if



response.write sHTML

%>
