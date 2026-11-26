<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  XWIN			 = request.querystring("XWIN")

  if XWIN <> "" THEN
   SESSION("XWIN")=XWIN
  ELSE
   XWIN=SESSION("XWIN")
  END IF

  SACTIVIDAD=    request.querystring("ACTIVIDAD")
  SRIESGO=    request.querystring("RIESGO")
  scodmaster = request.querystring("codmaster") & ""
  stipmaster = request.querystring("tipmaster") & ""


  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sOPCION       = request.querystring("OPCION")
  SCODINDI       = request.querystring("INDI")


  IF XWIN <> "" THEN
  IF  SESSION("ADDRELIND")="" THEN
     SESSION("ADDRELIND")= "&XWIN=" & XWIN &"&RIESGO=" & SRIESGO & "&ACTIVIDAD=" & SACTIVIDAD & "&CODMASTER=" &SCODMASTER & "&TIPMASTER=" & stipmaster
  END IF
  ELSE
   SESSION("ADDRELIND")=""
  END IF

  'vuelta=ofv.vueltaasp2("../")  &"_XWIN=" & XWIN

  vuelta2=ofv.vueltaasp("",CantRegMostrarx,Cantregmoverx)  &"&XWIN=" & XWIN




  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



'''''''''''''''''''''''''''''''''''''''
'' INCORPORO INDICADOR

  if sopcion = "SI" then

        sfechai = clng(date()) - clng(cdate("01-01-1900"))

        IF XWIN="ADDRELIND" THEN
         TIPO=2
        ELSE
         TIPO=1
        END IF
        strsql = "Insert into ORINDICADORES "
        strsql = strsql & " (CLASIFICACION,acceso,lstusr,FORMULA,CANTVAR,TIPO,DESCRIPCION,FECHACREACION,estado,usuario"
        strsql = strsql & " ) Values (1,0,1,16,1,"&TIPO&","
        strsql = strsql &  "'NUEVO INDICADOR'," & sfechai
        strsql = strsql & ",1,'" & session("usuario") & "')"
         xxx=SqlExLog("INSERT",strsql,CN,"ORINDICADORES","","","PAR")
        'call ofv.crearconsultaEx(StrSql,cn,1,volver)

       'RESPONSE.WRITE STRSQL



  end if


  if sOPCION = "ELIM" then



        strsql = "DELETE FROM ORINDICADORES WHERE OBJID = 0" & SCODINDI
         '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  END IF



SELECT CASE XWIN

 CASE "ADDRELIND"
    Strsql = "select * FROM orindicadores WHERE  (ESTADO IN (1,2) AND TIPO IN(1,2) AND USUARIO='"&SESSION("USUARIO")&"') OR (ESTADO=2 AND TIPO IN(1,2)) " 'AND tipo in (1,2) and USUARIO='" & SESSION("USUARIO") & "')"


  str1 = "Select descripcion From orindicadores WHERE  (ESTADO IN (1,2) AND TIPO IN(1,2) AND USUARIO='"&SESSION("USUARIO")&"') OR (ESTADO=2 AND TIPO IN(1,2)) " 'AND tipo in (1,2) and USUARIO='" & SESSION("USUARIO") & "')"    ' orden pagina
  str2 = "Select descripcion From orindicadores WHERE  (ESTADO IN (1,2) AND TIPO IN(1,2) AND USUARIO='"&SESSION("USUARIO")&"') OR (ESTADO=2 AND TIPO IN(1,2)) " 'AND tipo in (1,2) and USUARIO='" & SESSION("USUARIO") & "') order by descipcion "    ' orden alfabetico



 CASE ELSE
   Strsql = "select * FROM orindicadores WHERE  tipo<>999"

  str1 = "Select descripcion From orindicadores WHERE  tipo<>999"    ' orden pagina
  str2 = "Select descripcion From orindicadores WHERE  tipo<>999 order by descripcion"    ' orden alfabetico

 END SELECT

  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

   Celda = ""

   wkproccod = "&nbsp;"

   Celda = ofv.GenCelda("","uti  ","","","","","","TABLA DE INDICADORES","si")
  ' Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;'   ","","98% ","30","0","","-1","1","","",Fila,"si")




  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""



' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

    Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY CARGADOS INDICADORES</big>","si")
    Formulario = ofv.genrow("","","","","60","","",Celda,"si")

  ELSE



    Celda = ""
    Celda = ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","5%","","","","NOMBRE&nbsp;DEL&nbsp;INDICADOR","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","5%","","","","DESCRIPCION","si")


    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","5%","","","","FECHA&nbsp;CREACION","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' ALIGN=CENTER","","5%","","","","RESPONSABLE","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' align=center ","","5%","","","","PROPIEDADES","si")
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' align=center ","","1%","","","","ESTADO","si")
    IF XWIN="ADDRELIND" THEN
     Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#bbbbbb;color:#ffffff;' align=center ","","1%","","","","RIESGOS ASOCIADOS","si")
    END IF

    'Formulario = ofv.Encabezados("Descripcion","fecha","formula","cant var")

    Formulario = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  posI = 0

  WKSETTABLA="orindicadores"
  if not rs.eof then rs.move(CantRegMoverx)
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    xmodifica=cbool(Fmodifica(rs("estado"),"",""))

   if not xmodifica  then
	smod=" readonly "
    else
	 smod=""
    end if
    Celda = ""

     Strsql = "select * FROM orrelaindiriesgo WHERE idindicador=0" & RS(0)
	      'response.write strsql
	      set rst = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		   if not rst.eof and rs("estado")=2 then
		     TieneRel="SI"
		   else
		     TieneRel="NO"
		   end if
	  call ofv.cerrarconsulta(rst)

    'Eliminar
           IF XWIN="ADDRELIND" THEN
           VALUSU=rs("USUARIO")=SESSION("USUARIO")
           ELSE
            VALUSU= TRUE
           END IF
	    if bSuprimir  AND VALUSU  AND TIENEREL="NO" and rs("estado")<> 2 then
	      sHref = "asp/EliminarRegistro.asp?XWIN="&XWIN&"&ID=" & ofv.convertircar(rs(0)," ","_")
	      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
	      Celda = ofv.BotonEliminar(rs(1),sHref)
	    else
	      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
	    end if


    'ID

    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","1%","","","",sInput,"si")

    'indicador
    if rs("tipo")<> 999 then
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text" & smod,sForm,"45","dt","")
    else
     sinput=rs(1)
    end if
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'descripcion


	       nCol = 33


           wdescri =""
	       if  not  isnull(rs(3)) then
	           wdescri = rs(3)
	       end if

	       if  len(wdescri) = 0 then wdescri = " "

	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
	       t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=2 " & smod
	       if  bModificar then
	           t3b = t3b & " onblur=" & sForm & ".submit() "
	           t3b = t3b & " onkeypress=" & chr(34)
	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
	       else
	           t3b = t3b & " readonly >"
	       end if
	       xdatos =  wdescri
	       t3b    = t3b & xdatos & "</textarea></font>"
	       SINPUT = t3b



    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


    sInput = GTFechadorSH(RS("FECHACREACION"))
	Celda = Celda & ofv.gencelda("","UT ALIGN=CENTER","","","","","",sInput,"si")
	strsql_responsable = ""
	strsql_responsable = "SELECT U.DESCRIPCION FROM orusrlst O JOIN USUARIOS U ON U.ID = substring(O.USRRESP,5, len(O.USRRESP)) where indicador = " & RS("OBJID")
	'response.Write(strsql_responsable)
	SET RS_RESPONSABLE = ofv.crearconsultaEx(strsql_responsable,cn,1,parametros)
	sInput = ""
	do while (not RS_RESPONSABLE.eof)
		if sInput = "" then
			sInput = RS_RESPONSABLE("DESCRIPCION")
		else
			sInput = sInput & " /<br> " & RS_RESPONSABLE("DESCRIPCION")
		end if
		RS_RESPONSABLE.movenext
	loop
	call ofv.cerrarconsulta(RS_RESPONSABLE)
    'CALL GETINFOUSER(RS("USUARIO") )

    'sInput=session("SGU_nombre")
    Celda = Celda & ofv.gencelda("","UT ALIGN=CENTER","","","","","",sInput,"si")

    'CARGA

  	sImagenP = "&#0232;"
  	sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
  	sinput = "<a class='ft1n' " & sFuncAnc & " href='ORCARGAINFOINDI.asp?XFUNCION="&RS("FORMULA")&"&cantvar="&rs("cantvar")&"&ID=" & rs(0) & "&CantRegMover=" & CantRegMoverx
  	sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&volver=" & vuelta2
  	sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"
    'Celda = Celda & ofv.gencelda("","ft","LEFT","","","","",sInput,"si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:none;border-width:1;border-color:#BBBBBB;' align=center","","","","","",sInput,"si")

    SELECT CASE RS("ESTADO")
    CASE 3
      SINPUT="NO VIGENTE"
     CASE 2
      SINPUT="INICIADO"
     CASE  1
      SINPUT="NUEVO"
    END SELECT
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:none;border-width:1;border-color:#BBBBBB;' align=center","","","","","",sInput,"si")
    sImagenP = "&#0232;"
    sFuncAnc = ""
    sImageOF   = "<font face='corflaticon' size=2 color='#00FF00'  >" & sImagenP & "</font>"
    url="'ORELRIESGOS.asp?SCODINDI=" &rs(0) & "' ,'mywindow','status=0,toolbar=0,width=800,height=600'"
    IF XWIN="ADDRELIND" THEN
    Celda = Celda & ofv.gencelda("","","CENTER onclick="&chr(34)&"window.open("&url&");"&chr(34)  & " style='cursor:pointer;'","","","","",sImageOF,"si")
    END IF




    Filax  = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORINDICADORES&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar  then
    sAccion = "ORINDICADORES.asp?opcion=SI" & SESSION("ADDRELIND")
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  'consulta
  sAccion = "ORINDICADORESCons.asp?" & request.querystring
  Celda   = ofv.BotonConsultar(sAccion)


   'response.write chr(13) & XWIN & "22222222222"
   IF XWIN="ADDRELIND" THEN
    sAccion = "ORASIGNAINDICADOR.ASP?VOLVER=ORCIRCUITODET.ASP&OPCION=LISTAR" & SESSION("ADDRELIND")
    VOLVER = ofv.BotonVOLVER(sAccion)
   ELSE
   VOLVER   = ""
   END IF
   sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Buscar,Agregar,Celda,volver)
  ' sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",AGREGAR,Buscar,VOLVER)

  call ofv.cerrarconsulta(rs)




  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>