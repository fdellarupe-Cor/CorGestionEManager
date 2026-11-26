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
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  WSPSESION = "ORWHEREAGENDA"
  WSPSESION_PEN = "ORWHEREAGENDAPEN"

  session(WSPSESION)=""
  session(WSPSESION_PEN)=""
  sHTML = ofv.FormHeader(session("FormN"))

	session("SFILTROS") = ""
  WKVISION = ""


  VISION="AGNALL"
  'response.write vision
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

    sForm = "formlst"
    Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BUSQUEDA AGENDA" ,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")






  'Encabezados

    Filax  = ""

    Formulario = ""

    Celda = ""

    'LOTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","NUMERO PLAN","si")

    SINPUT = ofv.GenerarInput("ORPLAN","TODOS","agregar",sForm,"10","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")





    Celda = ""
	'FECHA INICIO
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA&nbsp;INI&nbsp;TAREAS&nbsp;PENDIENTES","si")
	SINPUT = ofv.GenerarInput("FINICIO","TODOS","agregar",sForm,"80","dt","")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	'Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

	Celda = ""
		'FECHA FIN
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA&nbsp;FIN&nbsp;TAREAS&nbsp;PENDIENTES","si")
	SINPUT = ofv.GenerarInput("FFIN","TODOS","agregar",sForm,"70","dt","")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

'	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


   Celda = ""
	'FECHA FIN DE PLAN DESDE
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA FIN DE PLAN DESDE","si")
	SINPUT = ofv.GenerarInput("FINICIOP","TODOS","agregar",sForm,"80","dt","")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

'	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

	Celda = ""
		'FECHA FIN DE PLAN HASTA
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA FIN DE PLAN HASTA","si")
	SINPUT = ofv.GenerarInput("FFINP","TODOS","agregar",sForm,"70","dt","")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	'Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

'********************************************************************************************************************************************************
	Celda = ""

		Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PERIODO","si")

		Strsql2 = "Select * From ORPERIODOS  where estado in ('OK', 'CL')"
		set rsPeriodos = ofv.crearconsultaEx(StrSql2,cn,1,parametros)

		SCOMSEP = chr(9)
		scomK   = "TODOS"
		scomD   = "TODOS"

		do while not rsPeriodos.eof
			scomK = scomK & SCOMSEP & rsPeriodos(0)
			scomD = scomD & SCOMSEP & rsPeriodos(1)
			rsPeriodos.movenext
		loop


		XTPRIODO = request.Form("ORPERIODO")


		if XTPRIODO = "" then XTPRIODO = "TODOS"

		sInput = ofv.gencomboFX(sForm,"ORPERIODO",XTPRIODO,scomK,scomD,"si")

		Celda = Celda & ofv.gencelda("","VTI colspan=1 STYLE='border-style:SOLID;border-width:1;border-color:#888888;background-color:#cacaca;'","","","","","",SINPUT,"si")
		Celda = Celda & ofv.gencelda("","VTI colspan=1 STYLE='border-style:SOLID;border-width:1;border-color:#888888;background-color:#cacaca;color:#ffffff;'","","","","","","Solo para elegir Proceso y SubProceso","si")

		'scodmaster = WKORPERIODOMAP
		'stipmaster = "PROCESO"

	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

		call ofv.cerrarconsulta(rsPeriodos)

'********************************************************************************************************************************************************



	Celda = ""
		     'PROCESOS
			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PROCESOS","si")

		IF (XTPRIODO = "" OR XTPRIODO = "TODOS" OR XTPRIODO = "0") THEN
			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","","Para seleccionar Procesos, elija un Período.","si")
		ELSE
			STRSQLPERIODO = "Select RGAMBAUTOEV From ORPERIODOS  where  OBJID = 0" & XTPRIODO
			set rsx11 = ofv.crearconsultaEx(STRSQLPERIODO,cn,1,parametros)
				PERIODO = rsx11(0)
			call ofv.cerrarconsulta(rsx11)


			Strsql = "Select OBJETOS.codobjeto ,OBJETOS.descri From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      		Strsql = Strsql & " AND codobjmaster = '" & PERIODO & "' and tipobjmaster = 'PROCESO'  "
      		Strsql = Strsql & " ORDER BY OBJETOS.DESCRI "

			'RESPONSE.Write(CHR(13) & CHR(13))
			'RESPONSE.Write(STRSQLPERIODO)
			'RESPONSE.Write(CHR(13) & CHR(13))
			'RESPONSE.Write(Strsql)
			'RESPONSE.Write(CHR(13) & CHR(13))

		      set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			 SCOMSEP = chr(9)
			 scomK   = "TODOS"
			 scomD   = "TODOS"

			 do while not rsx1.eof
				  scomK = scomK & SCOMSEP & rsx1(0)
				  scomD = scomD & SCOMSEP & rsx1(1)
				 rsx1.movenext
			 loop

			'response.write(chr(13) & chr(13))
			'response.write(Strsql)
			'response.write(chr(13) & chr(13))

			XTPROC = request.Form("ORPROC")
			if XTPROC = "" then XTPROC = "TODOS"

			sInput = ofv.gencomboFX(sForm,"ORPROC",XTPROC,scomK,scomD,"si")

			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
		END IF

	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

call ofv.cerrarconsulta(rsx1)

        Celda = ""
'**************************************************************************************************************************************************
'Comentado hasta nuevo aviso.

		'SUB PROCESO
			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","SUB PROCESOS","si")

		XTPROC = request.Form("ORPROC")
		XTPRIODO = request.Form("ORPERIODO")
		'response.Write(xtproc)
		'XTPROC = 523

	if(XTPROC = "" or XTPROC = "TODOS" or XTPROC = "0" or XTPRIODO = "" or XTPRIODO = "TODOS" or XTPRIODO = "0") then
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","","Para seleccionar SubProcesos, elija un Proceso.","si")
	else
		Strsql ="SELECT   ax.cobj,B.DESCRI , O1.descri, O2.descri AS Expr1, O3.descri AS Expr2, O4.descri AS Expr3, O5.descri AS Expr4, "
	   Strsql = strsql & "  O6.descri AS Expr5, O7.descri AS Expr6,"
	   Strsql = strsql & "  O8.descri AS Expr7, ax.mc1, ax.mc2, ax.mc3, ax.mc4, ax.mc5, ax.mc6, ax.mc7, ax.cobj, ax.TSEC "
	   Strsql = Strsql & " FROM         dbo.CsRelPath ax LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                       dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN"
	   Strsql = Strsql & "                dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto "
	   Strsql = Strsql & " inner join objetos b on (ax.cobj = b.codobjeto) "
	   Strsql = Strsql & "  inner join objprty n on (ax.cobj=n.codobjeto and estado=3)"
	   Strsql = Strsql & " WHERE (B.CODVINOBJ = 'VN' or B.CODVINOBJ = 'CR' or B.CODVINOBJ = 'AI') AND ("
	   Strsql = Strsql & " Ax.MC1 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC2 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC3 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC4 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC5 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC6 = 0" & XTPROC & " OR "
	   Strsql = Strsql & " Ax.MC7 = 0" & XTPROC & "  "
	   Strsql = Strsql & "  ) order by B.descri"

	'response.write "<br>" & strsql & "<br>" & chr(13)
       set rsSubProc = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       scomK="TODOS" & CHR(9)
       scomD="TODOS" & CHR(9)
       do while not rsSubProc.eof
        scomK = scomK & rsSubProc(0) & chr(9)
        scomD = scomd & rsSubProc(1) & chr(9)
        rsSubProc.movenext
       loop

	'response.write(chr(13) & chr(13))
	'response.write(Strsql)
	'response.write(chr(13) & chr(13))

       call ofv.cerrarconsulta(rsSubProc)

	   IF XSUBPROCESO = "" THEN XSUBPROCESO = "TODOS"

       sInput = ofv.gencomboFX(sForm,"ORSUBPROCESO",XSUBPROCESO,scomK,scomD,"no")

	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	end if



      'Celda  = ofv.GenCelda("","tt","","","","","","SUBPROCESO","si")
      'Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      'Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")



	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




        Celda = ""

'**************************************************************************************************************************************************

  		     'estado
  			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","ESTADOS","si")


			Strsql = "SELECT * FROM ORAGENDAESTADOS WHERE ACTIVO = 'S'"
			set rsEstados = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  			 SCOMSEP = chr(9)
  			 scomK   = "TODOS"
  			 scomD   = "TODOS"

			do while not rsEstados.eof
				  scomK = scomK & SCOMSEP & rsEstados(1)
				  scomD = scomD & SCOMSEP & "<font color=#ff0000>"& rsEstados(2) &"</font>"'&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &rsx1(1)
				  rsEstados.movenext
			loop

  			sInput = ofv.gencomboFX(sForm,"ORESTADO","TODOS",scomK,scomD,"no")

  			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

  	        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
			call ofv.cerrarconn(rsEstados)

     Celda = ""

	 'Fecha Cumplido
  			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","VENCIDOS / A VENCER","si")


  			 SCOMSEP = chr(9)
  			 scomK   = "TODOS"
  			 scomD   = "TODOS"

			 scomK = scomK & SCOMSEP & 0
		     scomD = scomD & SCOMSEP & "VENCIDO"

		     scomK = scomK & SCOMSEP & 1
		     scomD = scomD & SCOMSEP & "SIN VENCER"


  			sInput = ofv.gencomboFX(sForm,"ORCUMPLIDO","TODOS",scomK,scomD,"no")

  			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

  	        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

            Celda = ""

	'Fecha Cumplido


  			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","USUARIO","si")


			 Strsql = "select distinct responsable from oragenda where responsable is not null and responsable <> ''"
			 set rsUsuarios = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  			 SCOMSEP = chr(9)
  			 scomK   = "" '"TODOS"
  			 scomD   = "" '"TODOS"
  			XUID=request.form("ORUSUARIOS")

  			if XUID ="" then
              IF NOT rsUsuarios.eof THEN
                XUID=rsUsuarios(0)
               END IF
            end if

			do while not rsUsuarios.eof
				  scomK = scomK & SCOMSEP & rsUsuarios(0)
				  xx = getinfouser(rsUsuarios(0))
				  scomD = scomD & SCOMSEP & session("sgu_nombre")
				  rsUsuarios.movenext
			loop


  			sInput = ofv.gencomboFX(sForm,"ORUSUARIOS",XUID,scomK,scomD,"no")

  			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

  	        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
			call ofv.cerrarconn(rsUsuarios)



            Celda = ""




	  	    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","REASIGNAR TODOS","si")
             SCOMSEP = chr(9)
  			 scomK   = "SI" & SCOMSEP & "NO"
  			 scomD   = scomK
  			  XUID=request.form("tipo")

  			if XUID ="" then
               XUID="NO"
            end if






  			sInput = ofv.gencomboFX(sForm,"TIPO",XUID,scomK,scomD,"no")

  			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

  	        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


				call ofv.cerrarconn(rsUsuarios)


     	Celda = ""



		sAccion1 = "ORREASIGNAAGENDAEX.asp?OPCION=VS&VISION=" & VISION
		SINPUT = "<INPUT TYPE='BUTTON' VALUE='SELECCIONAR' CLASS='BT' onClick='document." & sForm & ".action=" & chr(34) & sAccion1 & chr(34) &";document." & sForm & ".submit();'>&nbsp;&nbsp;"

		Celda = Celda & ofv.gencelda("","ut colspan=3 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

		Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

		sAccion = "ORREASIGNAAGENDA.asp?OPCION=AGN&VISION=" & VISION
		Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")

       sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa STYLE='border-style:SOLID;border-width:1;border-color:#888888;' ","","70% ","","0","","0","0","","",Formulario,"si") & "</center>"

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""



  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>