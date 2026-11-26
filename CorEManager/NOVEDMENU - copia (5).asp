
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok = ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	shtml = ok
Else
	response.expires = 0


	If (SESSION("UN") & "") = "" Then


		xxxx = NEWEMPRE("NO")


		SEMPRE = ATREMP(0, 0)
		SESSION("UN") = SEMPRE & ""

		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1


	End If



	Set cn = ofv.conectar(ofv.strconn2)
	Set cn0 = ofv.conectar(ofv.strconn0)
	Set cn5 = ofv.conectar(ofv.strconn5)
	Novedform = "CLIMN16317206150930885546227"
	ofv.ObtenerAtributosNOdo Novedform, "", "EDITNOVED", "3"
	bModificarp = ofv.AModificar

	If bModificarP Then
		bModificarP="S"
	Else
		bModificarP="N"
	End If
	'response.write bModificar

	shtml = ofv.FormHeader(SESSION("FormN"))
	''shtml = replace(SHTML,"<body ","<body background-image='../CorEManager/Imagenes/Background_img/1396502.jpg' ")
	Shtml = shtml & "<style>"
	Shtml = shtml & "body {"
	'Shtml = shtml & "  background-image: url('../CorEManager/Imagenes/Background_img/montañas_y_agua.jpg');"
	Shtml = shtml & "  background-color: #f1f2f6;"
	'Shtml = shtml & "  background-repeat: no-repeat;"
	'Shtml = shtml & "  background-size: cover;"
	Shtml = shtml & "}"

	Shtml = shtml & "</style>"

	Shtml = shtml & " <form ACTION=UPDDATOS.ASP style='display:none' id=enviard name=enviard target=enviaf method=POST ><INPUT TYPE=TEXT VALUE='' ID=inforeg NAME=inforeg ></FORM><iframe  style='display:none' id=enviaf name=enviaf></iframe>"
	'Shtml = shtml & "<input type=text id=demo class=demo value=#ff6161> "&    armacontendor3("") '& "<script>document.getElementById('sTBH').style.display='none'</script>"
	Shtml = shtml &  armacontendor3("") '& "<script>document.getElementById('sTBH').style.display='none'</script>"
	xs = ""

	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & " function getPosition(obj) {"
	xs = xs & "    var curleft = curtop = 0;"
	xs = xs & "    if (obj.offsetParent) {"
	xs = xs & "        do {"
	xs = xs & "            curleft += obj.offsetLeft;"
	xs = xs & "            curtop += obj.offsetTop;"
	xs = xs & "        } while (obj = obj.offsetParent);"
	xs = xs & "    }"
	xs = xs & "   return curtop;"
	xs = xs & "}"

	xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If SESSION("APLMDCLI") = "BC" Then
		xs = xs & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If
	xs = xs & " myWin= open(Datos"

	xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom,"    '& Chr(34)
	' xs = xs & "'menubar=yes,toolbar=yes,height=h,width=w,resizable=yes'" '& Chr(34)
	xs = xs & "  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'"    '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

	s = "<script language=JavaScript>" & Chr(13)
	s = s & "function AbrirVentana(Datosx) {" & Chr(13)
	s = s & " var h=570;w=790 ;"
	'If SESSION("APLMDCLI") = "BC" Then
	s = s & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If

	s = s & " myWin= open(Datosx"
	s = s & "," & Chr(34) & "edi" & Chr(34) & ","    '& Chr(34)
	'  s = s & "menubar=no,toolbar=no,height=570,width=790,status=yes,resizable=yes" & Chr(34)
	s = s & "  'status=yes,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'"    '& CHR(34)
	s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)



	s = s & " } " & Chr(13)

	s = s & "function evtclick(e) {"

	s = s & "	        e.preventDefault();"
	s = s & "	        e.stopPropagation();"

	s = s & "	        return false;"



	s = s & "}"
	s = s & " </script>" & Chr(13)


	sHTML = sHTML & s & ofv.botoneraESSX("0","","","","","","","","","","","","")
	%>
	<script>

		function openFullscreen() {
		//var elem = document.getElementById("ix");
		//If (elem.requestFullscreen) {
		//	elem.requestFullscreen();
		//	} Else If (elem.webkitRequestFullscreen) { /* Safari */
		//	elem.webkitRequestFullscreen();
		//	} Else If (elem.msRequestFullscreen) { /* IE11 */
		//	elem.msRequestFullscreen();
		//	}
			}
			</script>
			<%
			sHTML = sHTML  & autorizaciones() & SUSCRIPCION() & "</body></html>"
			Call ofv.cerrarconn(cn)
			Call ofv.cerrarconn(cn0)
			Call ofv.cerrarconn(cn5)
		End If


		Function armacontendor4(p1)
			str="<div class='grid-stack' data-gs-width='12'>"
			str=str &       "<div class='grid-stack-item' data-gs-x='0' data-gs-y='0' data-gs-width='4' data-gs-height='2'>"
			str=str &        "<div class='grid-stack-item-content'>etetetertertert</div>"
			str=str &    "</div>"

			str=str &"</div>"
			armacontendor4=str

		End Function

		Function armacontendor3(p1)

			StrSql = "SELECT objid,idsolapa,titulo,activo,cantreg,codEmp,ORDEN,WIDTH,HEIGHT,posx,posy,icono FROM novedtitex where  ACTIVO='S' AND codemp=1 order by orden "
			'	response.write strsql
			Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			btne=""

			If bModificarP="S" Then btne="<div class='toggle-drag-button'><input type='checkbox' id='editSwitch' title='Modo edición'></div>"
			str= btne & "<div class='grid' >"
			Do While Not rsx.eof

				'xconteledor="<table WIDTH=100% HEIGHT=100% ><tr><td HEIGHT=10% align=center >" & RSX("TITULO")& "</td></tr><tr><td HEIGHT=80% ></td></tr> <tr><td HEIGHT=10% >footer</td></tr></table>"
				If rsx("idsolapa")=2 Then
					xconteledor= "<div class='contenedor2 rh'  >"
				Else
					xconteledor= "<div class='contenedor'  >"
				End If

				If rsx("idsolapa")=2 Then
					'xconteledor=xconteledor &     "<div class='header2'>" & RSX("TITULO") & " "  & rsx("idsolapa") & "</div>"
					'xconteledor=xconteledor &     "<div xwidth=100% xclass='contenedor2'   >"

				Else
					icono = "<i class='" & RSx("ICONO")& " '></i>" 'la-2x
					'sacar title con id del registro
					burbujaN=""
					'If rsx("idsolapa")=11 Or rsx("idsolapa") =5 or rsx("idsolapa") =8 Then
						'burbujaN="<span data-idnov='novedm" & rsx("idsolapa")  & "' id='burbula" & rsx("idsolapa")& "' class='burbujaNov'>0</span>"

					'End If
					'xconteledor=xconteledor &     "<div class='header rh' title='" & RSX("TITULO") & " " & rsx("idsolapa") & "'>" & RSX("TITULO") & " " & xicono & " " & <div class='pill'> &"cantreg"&</div>  "</div>"                  '& rsx("idsolapa") &"</div>"
					xconteledor=xconteledor &     "<div class='header rh' title='" & RSX("TITULO") & " " & rsx("idsolapa") & "'>" & RSX("TITULO") & " " & xicono & " " & BurbujaN & "</div>"                  '& rsx("idsolapa") &"</div>"
					If rsx("idsolapa")=11 Or rsx("idsolapa") =5 or rsx("idsolapa") =8 Then
					xconteledor=xconteledor &     "<div class='cant-fav'>" & rsx("cantreg") & " Documentos</div>" ''' para saber la cantidad de formularios que tienen cada tarjeta 26/12
					end if
					xconteledor=xconteledor &     "<div class='contenido'   >"
				End If
				xconteledor=xconteledor &       ArmaContenido(rsx("idsolapa"),rsx("cantreg"))
				' xconteledor=xconteledor &       "Aquí iría el buscador, los botones o cualquier contenido que desees insertar."
				If rsx("idsolapa") <> 2 Then xconteledor=xconteledor &     "</div>"
				If rsx("idsolapa")<> 2 Then
					xconteledor=xconteledor &     "<div class='footer'>"
					'	xconteledor=xconteledor &       "Última actualización: Abr 2025"
					xconteledor=xconteledor &     "</div>"
				End If
				xconteledor=xconteledor &"  </div>"

				str=str &   "<div  style='height:" & RSX("HEIGHT") & "px; width:"& RSX("WIDTH") &"px'" & "data-height=" & RSX("HEIGHT") & " data-width=" & RSX("WIDTH") & " data-left=" & rsx("posx") & " data-top=" & RSX("posy") &" data-id=" & rsx("objid") & " class='grid-item Rsize'><div xstyle='font-size:10' >" & xconteledor &"</div></div>"
				RSX.MOVENEXT
			Loop

			str=str & "</div>"
			armacontendor3=str
			Call ofv.cerrarconsulta(RSX)
		End Function


		Function ArmaContenido(solapa,cantregN)
			scontenido=""
			nFecini = 0

			If nFecini < CLng(CDate("01-01-1900")) Then
				nFecini = CLng(CDate("01-01-1900"))
			Else
				nFecini = 0
			End If

			dFechai = CLng(Date) - nFecini

			Select Case solapa
				Case 2
				StrSql = "SELECT distinct top " & cantregN & " a.*,isnull(b.imagen,'NO') AS imagen FROM csnovedades a"
				'StrSql = StrSql & " left join NOVEDADESIMG b on (b.novedad=a.codigo) "
				StrSql = StrSql & "  left join NOVEGENPARIMG b on ('NGP'+ cast (b.novegenpar as nvarchar)=a.codigo) "
				'StrSql = StrSql & " Where fecexp > " & dFechai
				'StrSql = StrSql & " and fecinicio < " & dFechai + 1
				'StrSql = StrSql & " where ("
				StrSql = StrSql & " where ("
				StrSql = StrSql & " a.tipo = 1 or a.tipo = 2 or a.tipo = 51"
				StrSql = StrSql & ")"
				StrSql = StrSql & " And ("
				StrSql = StrSql & " (a.usugrp = '" & SESSION("usuario") & "') or "
				StrSql = StrSql & " (a.usugrp = 'TODOS' and a.sector = 999 )    or "
				StrSql = StrSql & " (a.usugrp = 'TODOS' and a.sector = 0" & SESSION("UN") & ") OR "
				StrSql = StrSql & " (a.usugrp = 'CORP' and a.sector = 0" & SESSION("UN") & ")"
				StrSql = StrSql & ") "
				StrSql = StrSql & " order by a.fecinicio desc, a.tipo desc, a.importancia, a.objid desc"
				'response.write strsql
				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido=" <div class='swiper mySwiper'>"
					scontenido= scontenido & " <div class='swiper-wrapper'>"
					'	<div class="swiper-slide">
					'scontenido="<ul class=slider>"
					scontenido= scontenido & "<div class='swiper-slide video-container'>"
					scontenido= scontenido & "       <iframe id=ixs name=ixs height=100% width=100% src='https://www.youtube.com/embed/hga-hbnw66I?enablejsapi=1' controls ></iframe>"
					scontenido= scontenido & "</div>"

					Do While Not rsx.EOF
						If rsx("imagen") <> "NO" Then
							scontenido= scontenido & " <div class='swiper-slide' STYLE=""background-size: 100% 100%;background-image: url('"& rsx("imagen") & "');"">"
						Else
							scontenido= scontenido & " <div class='swiper-slide'>"
						End If


						scontenido= scontenido & "     <div class='title' data-swiper-parallax='-300'>" & rsx("msgtitulo")& "</div>"
						scontenido= scontenido & "    <div class='subtitle' data-swiper-parallax='-200'>" & rsx("msgcopete") &"</div>"
						' if rsx("imagen") <> "NO" THEN
						' scontenido= scontenido & "     <div class='text' STYLE=""background-image: url('"& rsx("imagen") & "');""  data-swiper-parallax='-100'>"
						' ELSE
						scontenido= scontenido & "     <div class='text' data-swiper-parallax='-100'>"
						' END IF
						scontenido= scontenido & "            <p class=pnoved>"
						scontenido= scontenido & rsx("mensaje")
						scontenido= scontenido & "            </p>"
						scontenido= scontenido & "      </div>"
						scontenido= scontenido & "</div>"

						If Trim(rsx("mensaje")&"") <> "" Then



							If RSX("IMAGEN") <> "NO" Then
								'	scontenido= scontenido & "<li   ><img  width='100%' height='100%'  src='" & replace(RSX("IMAGEN") & "","///","/") & "'><span class='txtsld'>"  & rsx("mensaje") & "</span></li>"

							Else
								If rsx("msgimagen")=2 And Trim(rsx("msgurldsc")) & "" <> "" Then 'video
									'   response.write  ConvertYouTubeURL(rsx("msgurldsc"))
									'scontenido= scontenido & "<li ><div xstyle='display: grid;grid-template-columns: auto auto;grid-column-gap:13px;justify-content:space-between;'><iframe width='100%' height='100%' src='" & ConvertYouTubeURL(rsx("msgurldsc")) & "' frameborder='0' allow='accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture' allowfullscreen></iframe><span class='txtsld'></span></div></li>"
									'scontenido= scontenido & "<li><span class'=txtsld'>"  & rsx("msgurldsc") & "</span></li>"
								Else
									'scontenido= scontenido & "<li  ><span class='txtsld'>"  & rsx("mensaje") & "</span></li>"
								End If
							End If
						End If
						rsx.movenext


					Loop
					'scontenido= scontenido & "</ul>"
					scontenido= scontenido & "</div><div class='swiper-button-next2'></div>"
					scontenido= scontenido & "    <div class='swiper-button-prev2'></div>"
					scontenido= scontenido & "    <div class='swiper-pagination'></div>"
					scontenido= scontenido & "    <div class='autoplay-progress'>"
					scontenido= scontenido & "      <svg viewBox='0 0 48 48'>"
					scontenido= scontenido & "        <circle cx='24' cy='24' r='20'></circle>"
					scontenido= scontenido & "      </svg>"
					scontenido= scontenido & "      <span></span>"
					scontenido= scontenido & "    </div>"
					scontenido= scontenido & " </div>"


				End If
				Call ofv.cerrarconsulta(rsx)

				Case 26
				StrSql = "SELECT distinct top 10 a.*,isnull(b.imagen,'NO') AS imagen FROM csnovedades a"
				StrSql = StrSql & " left join NOVEDADESIMG b on (b.novedad=a.codigo) "
				'StrSql = StrSql & " Where fecexp > " & dFechai
				'StrSql = StrSql & " and fecinicio < " & dFechai + 1
				'StrSql = StrSql & " where ("
				StrSql = StrSql & " where ("
				StrSql = StrSql & " a.tipo = 1 or a.tipo = 2 or a.tipo = 51"
				StrSql = StrSql & ")"
				StrSql = StrSql & " And ("
				StrSql = StrSql & " (a.usugrp = '" & SESSION("usuario") & "') or "
				StrSql = StrSql & " (a.usugrp = 'TODOS' and a.sector = 999 )    or "
				StrSql = StrSql & " (a.usugrp = 'TODOS' and a.sector = 0" & SESSION("UN") & ") OR "
				StrSql = StrSql & " (a.usugrp = 'CORP' and a.sector = 0" & SESSION("UN") & ")"
				StrSql = StrSql & ") "
				StrSql = StrSql & " order by a.fecinicio desc, a.tipo desc, a.importancia, a.objid desc"
				'response.write strsql
				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				If  rsx.EOF Then
					scontenido=""
				Else
					'	scontenido="<div class=slider>"
					'	Do While Not rsx.EOF
					'		If trim(rsx("mensaje")&"") <> "" Then
					'			If RSX("IMAGEN") <> "NO" Then
					'				scontenido= scontenido & "<div  style='height:125%;overflow:auto;' ><img  src='" & RSX("IMAGEN")& "'><div style='color: #fff;top: 38px;left: 221px;position: absolute;font-size: 50px;'>"  & rsx("mensaje") & "</div></div>"

					'			Else
					'				scontenido= scontenido & "<div  style='height:125%;overflow:auto;' ><div xstyle='position:absolute;top:10%; '>"  & rsx("mensaje") & "</div></div>"
					'			End If
					'		End If
					'		rsx.movenext


					scontenido="<ul class=slider>"
					Do While Not rsx.EOF
						If Trim(rsx("mensaje")&"") <> "" Then
							If RSX("IMAGEN") <> "NO" Then
								scontenido= scontenido & "<li   ><img  width='100%' height='100%'  src='" & Replace(RSX("IMAGEN") & "","///","/") & "'><span class='txtsld'>"  & rsx("mensaje") & "</span></li>"

							Else
								If rsx("msgimagen")=2 And Trim(rsx("msgurldsc")) & "" <> "" Then 'video
									'   response.write  ConvertYouTubeURL(rsx("msgurldsc"))
									scontenido= scontenido & "<li ><div xstyle='display: grid;grid-template-columns: auto auto;grid-column-gap:13px;justify-content:space-between;'><iframe width='100%' height='100%' src='" & ConvertYouTubeURL(rsx("msgurldsc")) & "' frameborder='0' allow='accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture' allowfullscreen></iframe><span class='txtsld'></span></div></li>"
									'scontenido= scontenido & "<li><span class'=txtsld'>"  & rsx("msgurldsc") & "</span></li>"
								Else
									'scontenido= scontenido & "<li  ><span class='txtsld'>"  & rsx("mensaje") & "</span></li>"
								End If
							End If
						End If
						rsx.movenext


					Loop
					scontenido= scontenido & "</ul>"
				End If
				Call ofv.cerrarconsulta(rsx)




				Case 3, 9

				XEMPRE = Split(session("EmpHbl"),Chr(9))

				DELIM=""
				idempre = ""

				For i=0 To UBound(XEMPRE)

					idempre = idempre & DELIM & XEMPRE(i)
					DELIM = ","

				Next

				'''    session("ModMain") = "CL"
				StrSql = "SELECT  distinct top " & cantregN & " a.*, isnull(b.cliente,1),c.aplic FROM csnovedades a "
				'	select case xtipoT
				'	  case "K"
				StrSql = StrSql & "inner join tipobjeto a1 on (a.tipobjeto=a1.tipobjeto ) "
				'  case "D"
				'  StrSql = StrSql & "inner join tipobjeto a1 on (a.tipobjeto=a1.tipobjeto and a1.TIPOCPT = 'D') "
				StrSql = StrSql & " inner JOIN objetos c ON (c.codobjeto = a.codobjeto AND c.TIPOBJETO=a.TIPOBJETO)"

				'end select
				'--------------------------------------------------------------------------
				'NUEVO 02/08/2021 GP --> ALE
				'10317-0004/20
				Select Case session("APLMDCLI")

					Case "GP"

					' StrSql = StrSql & " inner JOIN objetos c ON (c.codobjeto = a.codobjeto AND c.TIPOBJETO=a.TIPOBJETO)"
					StrSql = StrSql & " inner JOIN OBJMEMPALC  c1 ON (c1.codobjeto = a.codobjeto AND c1 .TIPOBJETO=a.TIPOBJETO)"

				End Select
				'--------------------------------------------------------------------------
				StrSql = StrSql & " left join clienteplan b on (a.sector=b.id) "
				StrSql = StrSql & " Where a.fecexp > " & dFechai
				StrSql = StrSql & " and a.fecinicio <  " & dFechai + 1
				StrSql = StrSql & " And (a.tipo = 3 or a.tipo = 9)"
				'If session("APLMDCLI")<> "BPN" Then			'Se quita excepci n para BPN
				StrSql = StrSql & " And ("
				StrSql = StrSql & "    (a.usugrp = '" & session("usuario") & "') "
				StrSql = StrSql & " or ( a.usugrp = 'TODOS' and a.sector = 999)"
				If session("UN") = 1 Then
					StrSql = StrSql & " or ( a.usugrp = 'TODOS' and a.sector in (" & idempre & "))"
				Else
					StrSql = StrSql & " or ( a.usugrp = 'TODOS' and a.sector = 0" & session("UN") & ")"
				End If
				Select Case session("APLMDCLI")

					Case "GP"
					StrSql = StrSql & " or (c.objun=1 and c1.codempre=" & session("UN") & ")"
				End Select
				StrSql = StrSql & "     ) "


				StrSql = StrSql & " order by a.fecinicio desc, a.importancia, objid desc"
				'RESPONSE.WRITE STRSQL

				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'><ul >"

					Do While Not rsx.EOF
						xMAINAPL=rsx("aplic")
						sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&MAINAPL=" & xMAINAPL
						sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
						sHref = ""


						sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPxex(rsx(9), rsx(8), ofv,""), "si")
						scontenido= scontenido & "<li class=documento-publi><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
						rsx.movenext

					Loop
					scontenido= scontenido & "</ul></div>"
				End If


				Case 4

				'Set cn1 = ofv.conectar(ofv.strconn5)
				'Set cn4 = ofv.conectar(ofv.strconn6)


				If wkaplic = "OR" Or wkaplic = "PL" Then

					xtipos = " OBJTIPO = 'RSK' "


				Else

					xtipos = " OBJTIPO in ("
					separ = ""

					StrSql = "SELECT tipobjeto FROM tipobjeto "
					Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
					If Not rs.EOF Then
						Do While Not rs.EOF
							xtipos = xtipos & separ & "'" & rs(0) & "'"
							separ = ","
							rs.movenext
						Loop
					Else
						xtipos = xtipos & "' '"
					End If
					xtipos = xtipos & ")"

				End If





				SEGWHERE = " WHERE " & xtipos & " AND ( PSCFIN = 'NO'  or  ( fechadif < 0 and finp=1) ) AND ( (RESPASG = '" & session("Usuario") & "') "


				xxxx = getInfoUser(SESSION("Usuario"))
				xusrlog = SESSION("SGU_nombre")
				'RESPONSE.WRITE xusrlog

				If xusrlog = "NO" And SESSION("USRPRV") = "LACAJA" Then
					SESSION("USRPRV") = "I"
					SESSION("Usuario") = "USRCONSULTA"
					Call getInfoUser(SESSION("Usuario"))
					xusrlog = SESSION("SGU_nombre")
					SESSION("USRPRV") = "LACAJA"
				End If

				' xxxx = getInfoUser(session("Usuario"))
				' xusrlog = session("SGU_nombre")


				'  RESPONSE.WRITE xusrlog &" 222222222222222222222222222" & CHR(13)
				If xusrlog <> "NO" Then


					SEGWHERE = SEGWHERE & "OR ( ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & SESSION("Usuario") & "') ) "
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

					'''''''''''''''''''''''''''''''''''''''''''
					If SESSION("DELLOG") = "S" Then

						XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
						xcount = UBound(XMFTGRP) + 1
						XMFSEP = ""
						SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

						For XMFI = 0 To UBound(XMFTGRP)
							SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
							XMFSEP = ","
						Next

						SEGWHERE = SEGWHERE & ") ) )"

					Else
						SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
					End If


					''''''''''''''''''''''''''''''''''''''''''

					'''   IF  RS(21) > 0 THEN


					If CStr(SESSION("SGU_idcostcenter") & "") <> "0" Then
						SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
						SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & SESSION("SGU_idcostcenter") & "' OR SUBCODIGO = '" & SESSION("SGU_idcostcenter") & "' ) "
						SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
						SEGWHERE = SEGWHERE & " ) ) "
					End If


				End If
				If xusrlog <> "NO" Then
					SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "
				Else
					SEGWHERE = SEGWHERE & " ) ORDER BY PSCCOMIENZO "
				End If

				StrSql = "SELECT  top " & cantregN & " * FROM CKTVIEWex  " & SEGWHERE
				'response.write strsql
				Set rsx = ofv.crearconsultaEx(StrSql, cn5, 1, parametros)

				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow: scroll;height:100%;'> <ul >"

					Do While Not rsx.EOF
						StrSql = "SELECT a.*, b.cliente FROM objetos a "
						StrSql = StrSql & " inner join clienteplan b on (a.objun=b.id) "
						StrSql = StrSql & " Where a.codobjeto = 0" & rsx(4)
						Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						' response.write strsql & "<br>"
						If Not rsxp.EOF Then

							xMAINAPL = rsxp("APLIC")
						End If
						Call ofv.cerrarconsulta(rsxp)

						'xMAINAPL=rsx("aplic")
						'  sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
						' sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
						SESSION("ASP") = "novedmenu.asp?menu=4"
						SESSION("sEstado") = "PARA APROBAR"
						SESSION("sEstadon") = "2"
						SESSION("Form2") = "apr"
						If TIPOCPT="PD" Then
							aspeditar = "Editar.asp?tipocpt=PD&IDPEDIDO="&rsx(4)&"&ACC=PF&DOCMODAL=WKPED&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL
						ElseIf TIPOCPT="SN" Then
							aspeditar = "Editar.asp?tipocpt=SN&IDPEDIDO="&rsx(4)&"&ACC=SNN&DOCMODAL=WKPED&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL
						Else

							aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL
						End If
						sActcon=aspeditar
						sHref = ""


						sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
						XTSLNKM = ofv.GenLink("", "nt  title='editar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(3), rsx(4), ofv), "si")
						scontenido= scontenido & "<li class=documento-link><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
						rsx.movenext

					Loop
					scontenido= scontenido & "</ul></div>"
				End If
				Call ofv.cerrarconsulta(rsx)

				Case 5

				XORDER = " group by b1.codobjeto, d.cliente ORDER BY fecha desc "

				StrSql = ""
				StrSql = StrSql & "SELECT top " & cantregN & " b1.CODOBJETO, max(b1.TIPOBJETO), max(finicio) as fecha  "
				StrSql = StrSql & ",max(c.APLIC)as aplic FROM OBJALOGDET B1"
				StrSql = StrSql & " inner join  objprty b on (b1.CODOBJETO=b.codobjeto and b.estado <> 4) "
				StrSql = StrSql & " inner join objetos c on (b1.CODOBJETO=c.codobjeto and b1.tipobjeto=c.tipobjeto) "

				If solapa=11 Then
					StrSql = StrSql & " inner join  TIPOBJETO Tb on (b.TIPOBJETO=Tb.TIPOBJETO and tb.tipocpt ='K') "
				Else
					StrSql = StrSql & " inner join  TIPOBJETO Tb on (b.TIPOBJETO=Tb.TIPOBJETO and tb.tipocpt <>'K') "
				End If

				StrSql = StrSql & " inner join clienteplan d on (c.objun=d.id)  "

				If  session("UN")=1 Then
					StrSql = StrSql & " WHERE  TACC = 'CONSEN'  " & XWHERE
				Else
					StrSql = StrSql & " WHERE c.objun=" & session("UN") & " and TACC = 'CONSEN'  " & XWHERE
				End If
				StrSql = StrSql & XORDER

				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'><ul >"
					totalR=0
					Do While Not rsx.EOF
						xMAINAPL=rsx("aplic")
						sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
						sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
						sHref = ""


						sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")
						scontenido= scontenido & "<li class=documento-ultcons><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
						rsx.movenext
						totalR=totalR + 1

					Loop
					'scontenido = scontenido<div class='pill'>"&totalR&" Documentos favoritos</div>
					scontenido= scontenido & "</ul></div>"
				End If
				Call ofv.cerrarconsulta(rsx)
				Case 6
				scontenido=Favoritos(cantregN)
				Case 6777

				SESSION("Segce") = "S"
				SESSION("Segco") = "S"


				xtipos = " OBJTIPO in ('RSK'"
				separ = ","
				StrSql = "SELECT tipobjeto FROM tipobjeto "
				Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				If Not rs.EOF Then
					Do While Not rs.EOF
						xtipos = xtipos & separ & "'" & rs(0) & "'"
						separ = ","
						rs.movenext
					Loop
				Else
					xtipos = xtipos & "' '"
				End If
				xtipos = xtipos & ")"

				SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & SESSION("Usuario") & "') OR ("

				xxxx = getInfoUser(SESSION("Usuario"))
				xusrlog = SESSION("SGU_nombre")




				If xusrlog <> "NO" Then


					SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & SESSION("Usuario") & "') ) "
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

					'''''''''''''''''''''''''''''''''''''''''''
					If SESSION("DELLOG") = "S" Then

						XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
						xcount = UBound(XMFTGRP) + 1
						XMFSEP = ""
						SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

						For XMFI = 0 To UBound(XMFTGRP)
							SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
							XMFSEP = ","
						Next

						SEGWHERE = SEGWHERE & ") ) )"

					Else
						SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
					End If


					''''''''''''''''''''''''''''''''''''''''''

					'''   IF  RS(21) > 0 THEN


					If CStr(SESSION("SGU_idcostcenter") & "") <> "0" Then
						SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
						SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & SESSION("SGU_idcostcenter") & "' OR SUBCODIGO = '" & SESSION("SGU_idcostcenter") & "' ) "
						SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
						SEGWHERE = SEGWHERE & " ) ) "
					End If


				End If

				SEGWHERE = SEGWHERE & " ) ) ORDER BY objcod desc "

				StrSql = "SELECT  *,LEFT(pscdsc, CHARINDEX(':', pscdsc) - 1) as paso ,SUBSTRING(pscdsc, CHARINDEX(':', pscdsc) + 1, LEN(pscdsc)) AS descriPaso,fechadif as demora FROM CKTVIEWex " & SEGWHERE
				'response.write StrSql

				Set rsx = ofv.crearconsultaEx(StrSql, cn5, 1, parametros)


				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow: scroll;height:100%;'> "
					scontenido=scontenido &" <div class='swiper mySwiper2'>"
					scontenido= scontenido & " <div class='swiper-wrapper'>"
					cant=1
					sul=""
					sli=""
					Do While Not rsx.EOF
						wNomL = rsx("wkfdescri") & "-" & rsx("paso") & "&nbsp;-&nbsp;PASO&nbsp;" & rsx("descripaso") & "&nbsp;"  & "&nbsp; demora &nbsp;" & rsx("demora")

						nomDoc =  EsqnomPxex(rsx(3), rsx(4), ofv,"")
						' nomDoc=

						SESSION("ASP") = "novedmenu.asp?menu=4"
						SESSION("sEstado") = "PARA APROBAR"
						SESSION("sEstadon") = "2"
						SESSION("Form2") = "apr"

						aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8)


						''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

						aspeditar = CStr(aspeditar)
						sHref = "&nbsp;"
						'        response.write sActcon
						sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, nomDoc, "si")
						sli= sli & "<li class=documento-link><i class='lar la-file-alt la-1x></i>"  &wNomL &  XTSLNKM & "</li>"

						If cant Mod 3 = 0 Then
							sul=sul & "<div class='swiper-slide'>"& "<ul style='width:100%' >" & sli & "</ul>" & "</div>"
							sli=""
							cant=1
						Else
							cant=cant +1
						End If
						rsx.movenext

					Loop
					If sli <> "" Then  sul=sul & "<div class='swiper-slide'>"& "<ul style='width:100%' >" & sli & "</ul>" & "</div>"

					scontenido= scontenido & sul & "</div>"
					'scontenido= scontenido &    "<div class='swiper-button-next'></div>"
					'scontenido= scontenido &    "<div class='swiper-button-prev'></div>"
					scontenido= scontenido & "    <div class='swiper-pagination'></div>"
					scontenido= scontenido & " </div>"
					scontenido= scontenido &   "</div>"
				End If
				Call ofv.cerrarconsulta(rsx)
				Case 7


				If session("UN") = 1 Then
					StrSql = ""
					StrSql = StrSql & "SELECT top " & cantregN & " a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente,a.aplic"
					StrSql = StrSql & " from objetos a "
					StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
					StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
					StrSql = StrSql & " where  a.objun=1 and b.fecvto > 0 order by b.fecvto "
				Else
					StrSql = ""
					StrSql = StrSql & "SELECT top " & cantregN & " a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
					StrSql = StrSql & " from objetos a "
					StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
					StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
					StrSql = StrSql & " where b.fecvto > 0 and a.objun=" & session("UN") & " order by b.fecvto "
				End If
				' response.write strsql
				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul class=ulvencer>"
					Do While Not rsx.EOF
						dfecvto = rsx(2)
						dfecvto = CLng(dfecvto)
						dfechadia = CLng(Date) - CLng(CDate("01-01-1900"))
						If dfecvto > dfechadia Then
							'     response.write cdate(dfecvto) & " "  &  cdate(dfechadia) & rsx(1) & "<br>"
							StrSql = ""
							StrSql = StrSql & "SELECT tipobjeto, avisoant from tipobjeto where aplic in('EM','GLB') "
							StrSql = StrSql & " and tipobjeto = '" & rsx(1) & "' "
							'  RESPONSE.WRITE STRSQL & "<BR>"
							Set rsy = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
							If Not rsy.EOF Then
								'' RESPONSE.WRITE " linea 2190 "  & "<BR>"
								davisoant = 0
								davisoant = rsy(1)
								dfecaviso = rsx(2) - davisoant

								''RESPONSE.WRITE davisoant & " davisoant "  & "<BR>"
								''RESPONSE.WRITE dfecvto & " dfecvto "  & "<BR>"

								dfecaviso = CLng(dfecaviso)

								'	RESPONSE.WRITE dfecaviso & " dfecaviso "  & "<BR>"
								'RESPONSE.WRITE dFechadia & " dFechadia " & "<BR>"

								If (dfecaviso <= dfechadia And dfecvto >= dfechadia) Then
									'''''''''''''''''''''''''''''''If dfecaviso > dFechadia Then
									diferavis = 0
									diferavis = dfecaviso - dfechadia
									xdife = 0
									xdife = diferavis - davisoant

									xMAINAPL=rsx("aplic")
									sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
									sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
									sHref = ""


									sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
									XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")


									If CLng(dfechadia) = 0 Then
										dFecha = "-"
									Else
										dFecha = CDate(CLng(dfechadia) + CDate("01-01-1900"))
									End If
									scontenido= scontenido & "<li class=documento-vencer "& sFuncion & "><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "<div class='pill'>Fec. Venc: "& fechadorCor(dfecha,"")  &"</div></li>"


								End If
							End If
						End If
						rsx.movenext

					Loop
					scontenido= scontenido & "</ul></div>"
				End If
				Call ofv.cerrarconsulta(rsx)

				Case 711
				If session("UN") = 1 Then
					StrSql = ""
					StrSql = StrSql & "SELECT a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
					StrSql = StrSql & " from objetos a "
					StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
					StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
					StrSql = StrSql & " where b.fecvto > 0 order by b.fecvto "
				Else
					StrSql = ""
					StrSql = StrSql & "SELECT a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
					StrSql = StrSql & " from objetos a "
					StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
					StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
					StrSql = StrSql & " where b.fecvto > 0 and a.objun=" & session("UN") & " order by b.fecvto "
				End If
				'response.write strsql & "22"
				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
					dfecvto = rsx(2)
					dfecvto = CLng(dfecvto)
					dfechadia = CLng(Date) - CLng(CDate("01-01-1900"))
					Do While Not rsx.EOF
						xMAINAPL = rsx("APLIC")

						If dfecvto > dfechadia Then

							StrSql = ""
							StrSql = StrSql & "SELECT tipobjeto, avisoant from tipobjeto where aplic in('EM','GLB') "
							StrSql = StrSql & " and tipobjeto = '" & rsx(1) & "' "
							''RESPONSE.WRITE STRSQL & "<BR>"
							Set rsy = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
							If Not rsy.EOF Then
								'' RESPONSE.WRITE " linea 2190 "  & "<BR>"
								davisoant = 0
								davisoant = rsy(1)
								dfecaviso = rsx(2) - davisoant
								dfecaviso = CLng(dfecaviso)


								If (dfecaviso <= dfechadia And dfecvto >= dfechadia) Then

									sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
									sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
									sHref = "&nbsp;"

									sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


									xtfechavig = CLng("0" & rsx(2))
									xtfechavig = GTFechadorsh(xtfechavig)
									wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;&nbsp;&nbsp;<font color='DARKRED' >(" & xtfechavig & ")</FONT>"
									XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")
									scontenido= scontenido & "<li class=documento-link><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
								End If
							End If
						End If
						rsx.movenext

					Loop
					scontenido= scontenido & "</ul></div>"
				End If
				Call ofv.cerrarconsulta(rsx)


				Case 8

				StrSql = "SELECT distinct top " & cantregN & " a.codobjeto,a.tipobjeto FROM OBJPRTY a "
				StrSql = StrSql & " INNER JOIN  OBJOBSCMT B ON (b.codsol=9999 and B.tipo='CMD' AND B.ACTPEND is null  AND B.CODOBJETO=a.CODOBJETO AND B.TIPOBJETO=a.TIPOBJETO ) "
				StrSql = StrSql & " INNER JOIN  TIPOBJETO C ON (C.COMENTARIOS='S'  AND A.TIPOBJETO=C.TIPOBJETO ) "
				StrSql = StrSql & " WHERE  (a.ESTADO = 3 )"
				StrSql = StrSql & " And (a.autor='"& SESSION("usuario") &"')"
				'response.write strsql

				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
					Do While Not rsx.EOF
						sActcon = "ConsultarcMD.asp?PEND=S&Codigo=" & rsx(0) & "&Tipo=" & rsx(1)
						sHref = "&nbsp;"


						sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")
						scontenido= scontenido & "<li class=documento-com><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
						rsx.movenext

					Loop
					scontenido= scontenido & "</ul></div>"
				End If
				Call ofv.cerrarconsulta(rsx)




				Case 11

				XWHERE = " and USUARIO = '" & session("usuario") & "' "
				XWHERE = XWHERE & " AND b1.CODOBJETO > 0 "

				If wkaplic = "OR" Or wkaplic = "PL" Then

					XWHERE = XWHERE & " AND b1.TIPOBJETO = 'RSK' "

				Else

					XWHERE = XWHERE & " AND b1.TIPOBJETO <> 'RSK' "

				End If

				XORDER = " group by b1.codobjeto ORDER BY fecha desc "


				StrSql = ""
				StrSql = StrSql & "SELECT top " & cantregN & " b1.CODOBJETO, max(b1.TIPOBJETO), max(finicio) as fecha  "
				StrSql = StrSql & ",max(c.APLIC)as aplic FROM OBJALOGDET B1"
				StrSql = StrSql & " inner join  objprty b on (b1.CODOBJETO=b.codobjeto and b.estado =3) "
				StrSql = StrSql & " inner join objetos c on (b1.CODOBJETO=c.codobjeto and b1.tipobjeto=c.tipobjeto) "
				StrSql = StrSql & " inner join  TIPOBJETO Tb on (b.TIPOBJETO=Tb.TIPOBJETO and tb.tipocpt ='K') "

				'StrSql = StrSql & " WHERE TACC = 'CONSEN'  " & XWHERE
				If  session("UN")=1 Then
					StrSql = StrSql & " WHERE  TACC = 'CONSEN'  " & XWHERE
				Else
					' StrSql = StrSql &  "inner join objetos b2 on (b2.CODOBJETO=b.codobjeto) "
					StrSql = StrSql & " WHERE C.objun=" & session("UN") & " and TACC = 'CONSEN'  " & XWHERE
				End If
				StrSql = StrSql & XORDER

				Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


				If  rsx.EOF Then
					scontenido=""
				Else
					scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
					totalR=0
					Do While Not rsx.EOF
						xMAINAPL=rsx("aplic")
						sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
						sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
						sHref = ""


						sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
						XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")
						scontenido= scontenido & "<li class=documento-ultcons><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
						rsx.movenext
						totalR=totalR + 1
					Loop
					scontenido= scontenido & "</ul></div>"
				End If
				Call ofv.cerrarconsulta(rsx)
				Case 12
				scontenido="<div style='overflow-x: auto;overflow-y: auto;height:100%;width:100%'>  "
				scontenido=scontenido &  INDICE2() &"</div>"

				Call ofv.cerrarconsulta(rsx)

			End Select
			ArmaContenido=scontenido & "<div data-cant='"& totalR &"' style='display:none' id='novedm"  & solapa &"'></div>"
		End Function

		Function armacontendor2(p1)
			str= "<div class='grid'>"
			str=str &   "<div id=""resizable""	 class='grid-item grid-item--width2'></div>"
			str=str &  "<div class='grid-item grid-item--height2 grid-item--width3'></div>"
			str=str &  "<div class='grid-item'></div>"
			str=str &  "<div class='grid-item'></div>"
			str=str &  "<div class='grid-item grid-item--width2 grid-item--height2'></div>"
			str=str &  "<div class='grid-item grid-item--width2'></div>"
			str=str &  "<div class='grid-item grid-item--width2'></div>"
			str=str &  "<div class='grid-item grid-item--height2'></div>"
			str=str &  "<div class='grid-item grid-item--width3'></div>"
			str=str &   "<div class='grid-item grid-item--width2'></div>"
			str=str &  "<div class='grid-item grid-item--height2'></div>"
			str=str &  "<div class='grid-item'></div>"
			str=str &  "<div class='grid-item'></div>"
			str=str &  "<div class='grid-item grid-item--width2 grid-item--height2'></div>"
			str=str &  "<div class='grid-item'></div>"
			str=str &  "<div class='grid-item grid-item--width2'></div>"
			str=str &  "<div class='grid-item grid-item--height2'></div>"
			str=str &  "<div class='grid-item'></div>"

			str=str & "</div>"
			armacontendor2=str
		End Function

		Function armacontenedor(p1)

			sliderS="<div class=slider >"
			sliderS= sliderS & " <div >I am a slide.</div>"
			sliderS= sliderS & "  <div>I am another slide.</div>"
			sliderS= sliderS & "  <div>I am sdfsdfanother slide.</div>"
			sliderS= sliderS & "  <div>I am anobvcvbther slide.</div>"
			sliderS= sliderS & "  <div>I am anothfghher slide.</div>"
			sliderS= sliderS & "  <div>I am anotfghfher slide.</div>"
			sliderS= sliderS & "</div>"



			str="<div class='grid-container'>" & Chr(13)
			str= str & "<div ><ul id=ddd ><li><div>sdd<br>saddad</div></li><li><div>sdds<br>addad</div></li><li><div>sddsaddad</div></li><li><div>sddsaddad </div></li></ul></div>" & Chr(13)
			str= str & "<div>"  & dsliderS & "</div>" & Chr(13)
			str= str & "<div></div>" & Chr(13)
			str= str & "<div></div>" & Chr(13)
			str= str & "<div></div>" & Chr(13)
			str= str & "<div></div>" & Chr(13)
			str= str & "<div></div>" & Chr(13)

			str= str &  "</div>"



			armacontenedor=str
		End Function



		Function INDICE()
			session("sEstado")    = "APROBADA"
			session("sEstadon")   = "3"

			Fila    = ""



			xTrColor = "#FFFFFF"

			stitem = "Tema"
			sIcono = "ICorTema.jpg"
			sNuevoTipo = "SUBTEMAS"
			nNivel = 0

			sID = 0
			nv = 0
			cborderB = Session("CBRDBOT")
			cborderl = Session("CBRDLFT")
			borderB = ""
			borderl = ""



			sVerEstado2 = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

			If SESSION("EMPMUN") = "S" Then

				If SESSION("GLBMUN") = 1 Then
					sverestado2 = sverestado2 & " and Objetos.OBJUN = " & Session("UN") & " "
				Else
					sverestado2 = sverestado2 & " and "
					sverestado2 = sverestado2 & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
					sverestado2 = sverestado2 & "   ( OBJETOS.OBJUN = 1 AND "
					sverestado2 = sverestado2 & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
					sverestado2 = sverestado2 & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
					sverestado2 = sverestado2 & " ) "
				End If

			End If


			sverestado2 = sverestado2 & " and Objprty.Estado = 3 "

			xStrSqla = ""
			xStrSqla  = xStrSqla & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

			'''''''''''''''''''''''''''''''''''''''''''
			If  Session("DELLOG") = "S" Then

				XMFTGRP = Split(session("SegGrpLst"),Chr(9))
				xcount  = UBound(XMFTGRP) + 1
				XMFSEP  = ""
				xStrSqla = xStrSqla & " or grpusr in ("

				For XMFI = 0 To UBound(XMFTGRP)
					xStrSqla = xStrSqla  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
					XMFSEP  = ","
				Next

				xStrSqla = xStrSqla & ") "

			Else
				xStrSqla = xStrSqla & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
			End If

			''''''''''''''''''''''''''''''''''''''''''
			xStrSqla  = xStrSqla & ")"







			If SESSION("APLMDCLI") = "BGAL" Then


				StrSqlX  =  " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

				'''''''''''''''''''''''''''''''''''''''''''
				If  Session("DELLOG") = "S" Then

					XMFTGRP = Split(session("SegGrpLst"),Chr(9))
					xcount  = UBound(XMFTGRP) + 1
					XMFSEP  = ""
					StrSqlx = StrSqlx & " or grpusr in ("

					For XMFI = 0 To UBound(XMFTGRP)
						StrSqlX = StrSqlX  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
						XMFSEP  = ","
					Next

					StrSqlX = StrSqlX & ") "

				Else
					StrSqlX = StrSqlX & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				StrSqlX  = StrSqlX & ")"

				Strsql = "Select distinct  a.codtema,a.descri,a.pstipo,a.psurl,a.aplic, a.inda   From Tema a "
				Strsql = strsql & " inner join prodserv b on (a.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
				Strsql = strsql & " where a.PSTIPO <> 'K' AND (a.aplic = '" & wkaplic & "' or a.aplic = 'GLB') "
				Strsql = strsql & " union "
				Strsql = strsql & " Select distinct a.codtema,a.descri,a.pstipo,a.psurl,a.aplic, a.inda From Tema a "
				Strsql = strsql & " where (a.PSTIPO <> 'K' and a.PSTIPO = 'R') "
				Strsql = strsql & " order by a.inda, a.descri "
				'response.write strsql


				Strsql = " Select DISTINCT D.codtema,D.descri,D.pstipo,D.psurl,D.aplic,D.inda"
				Strsql = strsql & "  From "
				Strsql = strsql & " Prodserv A"
				Strsql = strsql & " INNER JOIN SUBTEMA B ON (A.marcaps = B.tipobjeto  )"
				Strsql = strsql & " INNER JOIN TEMA D ON (D.CODTEMA= A.CODTEMA )"
				Strsql = strsql & " LEFT JOIN subtemaskr C ON (A.codprodserv = C.codprodserv  " & StrSqlX &")"

				Strsql = strsql & " where   a.marcaps <> 'PEDIDO' and a.marcaps <> 'SNN' and a.PSTIPO <> 'K' AND (d.aplic = '" & wkaplic & "' or d.aplic = 'GLB')  and (A.SEGACESO <> '1' OR (C.OBJID IS NOT NULL))"



				Strsql = strsql & " union "
				Strsql = strsql & " Select distinct D.codtema,D.descri,D.pstipo,D.psurl,D.aplic, D.inda From Tema d "
				Strsql = strsql & " where (D.PSTIPO <> 'K' and D.PSTIPO = 'R') "
				Strsql = strsql & " ORDER BY  D.inda, D.descri"
				' response.write strsql


			Else

				Strsql = "Select codtema,descri,pstipo,psurl,aplic, inda  From Tema "
				Strsql = strsql & " where (aplic = '" & wkaplic & "' or aplic = 'GLB') and PSTIPO IN ('C','R') order by inda, descri"

			End If


			' response.write strsql
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  rs.eof Or session("un") = 0 Then

				Celda   = ""
				Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
				Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

				Fila    = fila & ofv.GenRow("","","","","25","","",Celda,"si")

			Else

				Celda   = ""



				Fila=""
				codtema=0
				isubtema=0
				ISUBCARPETA=""

				If (SESSION("UN") & "") = "" Then


					xxxx = NEWEMPRE("NO")


					SEMPRE = ATREMP(0, 0)
					SESSION("UN") = SEMPRE & ""

					If (SESSION("UN") & "") = "" Then SESSION("UN") = 1

				Else
					'  xxxx = NEWEMPRE(SESSION("UN"))


				End If
				indice= " <div class='swiper mySwiperInd'>"
				indice= indice & " <div class='swiper-wrapper'>"
				cant=1
				sul=""
				sli=""

				Do Until rs.eof

					xxnova = "S"
					Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs("codtema") & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
					'      response.write strsql
					Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

					If Not  rs1.eof Then
						If rs1(4) = 0 Then xxnova = "N"
					End If

					If xxnova = "S" Then
						sxlink = "ObjTreeSubTemaCons.asp?"
						sHref  = ""
						'   if codtema <> rs("codtema") then
						'   codtema=rs("codtema")

						If rs("aplic") = "GLB" Then
							'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;&</font>&nbsp;&nbsp;"
						Else
							'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>&nbsp;&nbsp;"
							'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'></font>&nbsp;&nbsp;"
						End If


						'    xTrColor="#eeeeff"
						sitem2   = ofv.ConvertirCar(rs("DESCRI")," ","_")
						sxlink  = sxlink & "ID=" & rs("codtema") & "&sitem=" & sitem2 & "&stitem=" & Replace(stitem," ","_")
						sxlink  = sxlink & "&stnivel=0&tipo=SUBTEMAS&TIPOVISTA=CONS"
						sHref   = sxlink

						starget=""
						If rs("pstipo")="R" Then
							'sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
							'sImage  = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#990000'>4</font>&nbsp;&nbsp;"
							sHref   = rs("psURL") & " "
							If InStr(UCase(sHref),"HTTP://") = 0 Then sHref   = "HTTP://" & sHref
							starget=" targer=_new"
						End If


						sLink2   = ofv.GenLink("","'nt xnti'" ,sHref,starget,"","",Replace(rs("DESCRI")," "," ") & " ","si")


						'celda= ofv.GenCelda("","'ut uti' xbgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",sImage & slink2,"si")

						'Fila    = Fila & ofv.GenRow("","'aaa rowI' ","","","20","","",Celda,"si")
						sli= sli & "<li class=documento-link><i class='lar la-file-alt la-1x'></i>"  & sLink2 & "</li>"
						If cant Mod 11 = 0 Then
							sul=sul & "<div class='swiper-slide'>"& "<ul style='width:100%' >" & sli & "</ul>" & "</div>"
							sli=""
							cant=1
						Else
							cant=cant +1
						End If



					End If
					RS.MOVENEXT
				Loop

				If sli <> "" Then  sul=sul & "<div class='swiper-slide'>"& "<ul style='width:100%' >" & sli & "</ul>" & "</div>"

				indice= indice & sul & "</div>"
				indice= indice & "    <div class='swiper-pagination'></div>"
				indice= indice & " </div>"
				'indice= indice &   "</div>"

			End If






			Call OFV.cerrarconsulta(rs)





			'INDICE2 = ofv.GenTabla("","aaa   ","","100% ","","0","","0","0","","",Fila,"si")
		End Function

		Function INDICE2()
			session("sEstado")    = "APROBADA"
			session("sEstadon")   = "3"

			Fila    = ""



			xTrColor = "#FFFFFF"

			stitem = "Tema"
			sIcono = "ICorTema.jpg"
			sNuevoTipo = "SUBTEMAS"
			nNivel = 0

			sID = 0
			nv = 0
			cborderB = Session("CBRDBOT")
			cborderl = Session("CBRDLFT")
			borderB = ""
			borderl = ""



			sVerEstado2 = "  and Objetos.objnivel >= " & Session("SegNivel") & " "

			If SESSION("EMPMUN") = "S" Then

				If SESSION("GLBMUN") = 1 Then
					sverestado2 = sverestado2 & " and Objetos.OBJUN = " & Session("UN") & " "
				Else
					sverestado2 = sverestado2 & " and "
					sverestado2 = sverestado2 & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
					sverestado2 = sverestado2 & "   ( OBJETOS.OBJUN = 1 AND "
					sverestado2 = sverestado2 & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
					sverestado2 = sverestado2 & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
					sverestado2 = sverestado2 & " ) "
				End If

			End If


			sverestado2 = sverestado2 & " and Objprty.Estado = 3 "

			xStrSqla = ""
			xStrSqla  = xStrSqla & " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

			'''''''''''''''''''''''''''''''''''''''''''
			If  Session("DELLOG") = "S" Then

				XMFTGRP = Split(session("SegGrpLst"),Chr(9))
				xcount  = UBound(XMFTGRP) + 1
				XMFSEP  = ""
				xStrSqla = xStrSqla & " or grpusr in ("

				For XMFI = 0 To UBound(XMFTGRP)
					xStrSqla = xStrSqla  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
					XMFSEP  = ","
				Next

				xStrSqla = xStrSqla & ") "

			Else
				xStrSqla = xStrSqla & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
			End If

			''''''''''''''''''''''''''''''''''''''''''
			xStrSqla  = xStrSqla & ")"







			If SESSION("APLMDCLI") = "BGAL" Then


				StrSqlX  =  " and (grpusr = 'USR-" & UCase(Session("usuario")) & "' "

				'''''''''''''''''''''''''''''''''''''''''''
				If  Session("DELLOG") = "S" Then

					XMFTGRP = Split(session("SegGrpLst"),Chr(9))
					xcount  = UBound(XMFTGRP) + 1
					XMFSEP  = ""
					StrSqlx = StrSqlx & " or grpusr in ("

					For XMFI = 0 To UBound(XMFTGRP)
						StrSqlX = StrSqlX  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
						XMFSEP  = ","
					Next

					StrSqlX = StrSqlX & ") "

				Else
					StrSqlX = StrSqlX & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				StrSqlX  = StrSqlX & ")"

				Strsql = "Select distinct  a.codtema,a.descri,a.pstipo,a.psurl,a.aplic, a.inda   From Tema a "
				Strsql = strsql & " inner join prodserv b on (a.codtema = b.codtema and b.marcaps <> 'PEDIDO' and b.marcaps <> 'SNN') "
				Strsql = strsql & " where a.PSTIPO <> 'K' AND (a.aplic = '" & wkaplic & "' or a.aplic = 'GLB') "
				Strsql = strsql & " union "
				Strsql = strsql & " Select distinct a.codtema,a.descri,a.pstipo,a.psurl,a.aplic, a.inda From Tema a "
				Strsql = strsql & " where (a.PSTIPO <> 'K' and a.PSTIPO = 'R') "
				Strsql = strsql & " order by a.inda, a.descri "
				'response.write strsql


				Strsql = " Select DISTINCT D.codtema,D.descri,D.pstipo,D.psurl,D.aplic,D.inda"
				Strsql = strsql & "  From "
				Strsql = strsql & " Prodserv A"
				Strsql = strsql & " INNER JOIN SUBTEMA B ON (A.marcaps = B.tipobjeto  )"
				Strsql = strsql & " INNER JOIN TEMA D ON (D.CODTEMA= A.CODTEMA )"
				Strsql = strsql & " LEFT JOIN subtemaskr C ON (A.codprodserv = C.codprodserv  " & StrSqlX &")"

				Strsql = strsql & " where   a.marcaps <> 'PEDIDO' and a.marcaps <> 'SNN' and a.PSTIPO <> 'K' AND (d.aplic = '" & wkaplic & "' or d.aplic = 'GLB')  and (A.SEGACESO <> '1' OR (C.OBJID IS NOT NULL))"



				Strsql = strsql & " union "
				Strsql = strsql & " Select distinct D.codtema,D.descri,D.pstipo,D.psurl,D.aplic, D.inda From Tema d "
				Strsql = strsql & " where (D.PSTIPO <> 'K' and D.PSTIPO = 'R') "
				Strsql = strsql & " ORDER BY  D.inda, D.descri"
				' response.write strsql


			Else

				Strsql = "Select codtema,descri,pstipo,psurl,aplic, inda  From Tema "
				Strsql = strsql & " where (aplic = '" & wkaplic & "' or aplic = 'GLB') and PSTIPO IN ('C','R') order by inda, descri"

			End If


			' response.write strsql
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  rs.eof Or session("un") = 0 Then

				Celda   = ""
				Celda   = Celda & ofv.GenCelda("","ttt","","2%","","","","&nbsp;","si")
				Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","","No hay elementos asociados","si")

				Fila    = fila & ofv.GenRow("","","","","25","","",Celda,"si")

			Else

				Celda   = ""



				Fila=""
				codtema=0
				isubtema=0
				ISUBCARPETA=""

				If (SESSION("UN") & "") = "" Then


					xxxx = NEWEMPRE("NO")


					SEMPRE = ATREMP(0, 0)
					SESSION("UN") = SEMPRE & ""

					If (SESSION("UN") & "") = "" Then SESSION("UN") = 1

				Else
					'  xxxx = NEWEMPRE(SESSION("UN"))


				End If

				Do Until rs.eof

					xxnova = "S"
					Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs("codtema") & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
					'      response.write strsql
					Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

					If Not  rs1.eof Then
						If rs1(4) = 0 Then xxnova = "N"
					End If

					If xxnova = "S" Then
						sxlink = "ObjTreeSubTemaCons.asp?"
						sHref  = ""
						'   if codtema <> rs("codtema") then
						'   codtema=rs("codtema")

						If rs("aplic") = "GLB" Then
							'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;&</font>&nbsp;&nbsp;"
						Else
							'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>&nbsp;&nbsp;"
							'sImage = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#f7aa22'></font>&nbsp;&nbsp;"
						End If


						'    xTrColor="#eeeeff"
						sitem2   = ofv.ConvertirCar(rs("DESCRI")," ","_")
						sxlink  = sxlink & "ID=" & rs("codtema") & "&sitem=" & sitem2 & "&stitem=" & Replace(stitem," ","_")
						sxlink  = sxlink & "&stnivel=0&tipo=SUBTEMAS&TIPOVISTA=CONS"
						sHref   = sxlink

						starget=""
						If rs("pstipo")="R" Then
							'sImage  = "<font face='corflaticon' size=3 color='#990000'>4</font>"
							'sImage  = "&nbsp;&nbsp;&nbsp;&nbsp;<font face='CorFlaticon' size=2 color='#990000'>4</font>&nbsp;&nbsp;"
							sHref   = rs("psURL") & " "
							If InStr(UCase(sHref),"HTTP://") = 0 Then sHref   = "HTTP://" & sHref
							starget=" targer=_new"
						End If


						sLink2   = ofv.GenLink("","'nti'" ,sHref,starget,"","",Replace(rs("DESCRI")," "," ") & " ","si")


						celda= ofv.GenCelda("","'ut uti' xbgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",sImage & slink2,"si")

						Fila    = Fila & ofv.GenRow("","'aaa rowI' ","","","20","","",Celda,"si")

					End If
					RS.MOVENEXT
				Loop

			End If






			Call OFV.cerrarconsulta(rs)




			INDICE2 = ofv.GenTabla("","aaa   ","","100% ","","0","","0","0","","",Fila,"si")
		End Function


		Function autorizaciones()
			SESSION("Segce") = "S"
			SESSION("Segco") = "S"


			xtipos = " OBJTIPO in ('RSK'"
			separ = ","
			StrSql = "SELECT tipobjeto FROM tipobjeto "
			Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
			If Not rs.EOF Then
				Do While Not rs.EOF
					xtipos = xtipos & separ & "'" & rs(0) & "'"
					separ = ","
					rs.movenext
				Loop
			Else
				xtipos = xtipos & "' '"
			End If
			xtipos = xtipos & ")"

			SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & SESSION("Usuario") & "') OR ("

			xxxx = getInfoUser(SESSION("Usuario"))
			xusrlog = SESSION("SGU_nombre")




			If xusrlog <> "NO" Then


				SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & SESSION("Usuario") & "') ) "
				SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

				'''''''''''''''''''''''''''''''''''''''''''
				If SESSION("DELLOG") = "S" Then

					XMFTGRP = Split(SESSION("SegGrpLst"), Chr(9))
					xcount = UBound(XMFTGRP) + 1
					XMFSEP = ""
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

					For XMFI = 0 To UBound(XMFTGRP)
						SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
						XMFSEP = ","
					Next

					SEGWHERE = SEGWHERE & ") ) )"

				Else
					SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
				End If


				''''''''''''''''''''''''''''''''''''''''''

				'''   IF  RS(21) > 0 THEN


				If CStr(SESSION("SGU_idcostcenter") & "") <> "0" Then
					SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
					SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & SESSION("SGU_idcostcenter") & "' OR SUBCODIGO = '" & SESSION("SGU_idcostcenter") & "' ) "
					SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
					SEGWHERE = SEGWHERE & " ) ) "
				End If


			End If

			SEGWHERE = SEGWHERE & " ) ) ORDER BY objcod desc "

			'StrSql = "SELECT  *,LEFT(pscdsc, CHARINDEX(':', pscdsc) - 1) as paso ,SUBSTRING(pscdsc, CHARINDEX(':', pscdsc) + 1, LEN(pscdsc)) AS descriPaso,fechadif as demora FROM CKTVIEW " & SEGWHERE
			cantregN=10
			StrSql = "SELECT top " & cantregN & " *, "
			StrSql = StrSql & " CASE WHEN CHARINDEX(':', pscdsc) > 0 THEN LEFT(pscdsc, CHARINDEX(':', pscdsc) - 1) ELSE pscdsc END as paso, "
			StrSql = StrSql & " CASE WHEN CHARINDEX(':', pscdsc) > 0 THEN SUBSTRING(pscdsc, CHARINDEX(':', pscdsc) + 1, LEN(pscdsc)) ELSE '' END AS descriPaso, fechadif as demora "
			StrSql = StrSql & " FROM CKTVIEW " & SEGWHERE
			'response.write StrSql

			Set rsx = ofv.crearconsultaEx(StrSql, cn5, 1, parametros)


			If  rsx.EOF Then
				scontenido=""
			Else
				scontenido="<div  style='overflow: scroll;height:100%;'> "
				scontenido=scontenido &" <div class='swiper mySwiper2'>"
				scontenido= scontenido & " <div class='swiper-wrapper'>"
				cant=0
				sul=""
				sli=""
				Do While Not rsx.EOF
					wNomL = rsx("wkfdescri") & "-" & rsx("paso") & "&nbsp;-&nbsp;PASO&nbsp;" & rsx("descripaso") & "&nbsp;"  & "demora " & rsx("demora")

					nomDoc =  EsqnomPx(rsx(3), rsx(4), ofv)
					' nomDoc=

					SESSION("ASP") = "novedmenu.asp?menu=4"
					SESSION("sEstado") = "PARA APROBAR"
					SESSION("sEstadon") = "2"
					SESSION("Form2") = "apr"

					aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8)


					''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

					aspeditar = CStr(aspeditar)
					sHref = "&nbsp;"
					'        response.write sActcon
					sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

					XTSLNKM = ofv.GenLink("", "'nt nt2'  title='Consultar Documento...' ", sHref, "", "", sFuncion, nomDoc, "si")
					'sli= sli & "<li class=documento-link><i class=Ciclass></i><TABLE BORDER=1><TR><TD>Paso " & RSX("PASO") & "</TD><TD>" &  XTSLNKM & rsx("descripaso") &"</TD><TD>Fecha comienzo " & RSX("comienzo") & "</TD> <TD>Dias demora " & RSX("demora") & "</TD></TR></TABLE></li>"
					'sli = sli & "<li class='documento-link'><i class='Ciclass'></i><table border='0' style='width:100%; table-layout:fixed;'><tr><td style='width:10%'>Paso " & RSX("PASO") & "</td><td style='width:40%'>" & XTSLNKM & "</td><td style='width:30%'>Fecha comienzo " & RSX("comienzo") & "</td><td style='width:30%'>Dias demora " & RSX("demora") & "</td></tr>"

					sli = sli & "<li class='documento-link'><i class='lar la-file-alt la-1x'></i><table border='0' style='width:100%; table-layout:fixed;'>"

					' --- Fila 1: Cabeceras (Agregadas en el paso anterior) ---
					sli = sli & "<tr><th class='th-noti' style='width:10%; xbackground-color:" & Session("CMFPNTIT") & "' >Paso</th><th class='th-noti' style='width:40%'>Descripcion</th><th class='th-noti' style='width:30%'>Fecha Comienzo</th><th class='th-noti' style='width:30%'>Dias Demora</th></tr>"

					' --- Fila 2: Datos Principales ---
					sli = sli & "<tr><td style='padding: 5px; text-align: center;'>" & RSX("PASO") & "</td><td class='pill' style='text-align: center;'>" & XTSLNKM & "</td><td style='text-align: center;'>" & RSX("comienzo") & "</td><td style='text-align: center;'>" & RSX("demora") & "</td></tr>"

					' --- Fila 3: Descripción Detallada (Nueva fila con Colspan 4) ---
					sli = sli & "<tr><td colspan='4' style='border-radius: 15px; font-weight: 650; text-align: center; padding: 5px; background-color: #f9f9f9; color: #555;'>" & RSX("descripaso") & "</td></tr>"

					sli = sli & "</table></li>"


					rsx.movenext
					cant=cant +1
				Loop


			End If
			Call ofv.cerrarconsulta(rsx)




			xcontenedor="<div id=pfav data-cant='" & cant& "' title='MIS AUTORIZACIONES'><div class='contenedor'  >"
			'  xcontenedor=xcontenedor &     "<div class='header'></div>"
			xcontenedor=xcontenedor &     "<div class='contenido'  >"
			xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "
			xcontenedor=xcontenedor &  "<ul>" & sli &"</ul></div>"
			'    xcontenedor=xcontenedor &     "</div>"
			xcontenedor=xcontenedor &"  </div></div>"
			autorizaciones = xcontenedor

		End Function


		Function Favoritos(cantregN)
			Dim XTablaN
			StrSql = " Select top " & cantregN & " *,B.ESTADO,B.tipobjeto as tipobjeto,B.fecvigen as fecvigen  From objmsg a " '
			StrSql = StrSql & " LEFT JOIN OBJPRTY B ON(A.MSGOBJCOD= B.CODOBJETO AND B.tipobjeto = A.msgobjtip ) " 'agrego AND B.tipobjeto = A.msgobjtip
			StrSql = StrSql & " inner JOIN objalogdet C ON (C.OBJID=a.objid) "
			StrSql = StrSql & " where a.MSGTIPOMESS=1 and b.estado=3 AND usr = '" & SESSION("USUARIO") & "'"
			'response.write strsql
			xTrColor = ""
			Set rsfav = ofv.crearconsultaEx(StrSql, cn, 1, parametros)




			If  rsfav.EOF  Then



			Else
				totalR=0
				Do While Not rsfav.EOF


					sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsfav("MSGOBJCOD") & "&Tipo=" & rsfav("MSGOBJTIP") & "&MAINAPL=" & xMAINAPL
					sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
					sHref = "&nbsp;"
					sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)

					If rsfav("ESTADO") & "" = "3" Then

						TITULOLNK = "Consultar Documento..."

					Else
						sFuncion = ""
						sHref = "#"
						TITULOLNK = "EL Documento NO encuentra vigente debe eliminarse de Favoritos..."
					End If

					sTipobjeto = EsqnomPx(rsfav("MSGOBJTIP"), rsfav("MSGOBJCOD"), ofv)



					wmsgx = sTipobjeto & ""
					wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



					sImagen = wNomL

					If xarc = "NO" Then
						xarc = "&nbsp;"
					End If

					sLink = sImagen

					XTSLNKM = ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", ssFuncion, wmsgx, "si")

					sImage = ofv.GenImage("", "", "", "", "Imagenes/Botdelx.jpg", "0", "")
					vlk = "../novedmenu.asp?tab=view2"
					sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=objmsg&volver=" & vlk
					sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

					'CELDA = BotonEliminar("", sHref, "fav" & CLAVE,"N")

					'**** TITULO DE LAS AUTORIZACIONES ****'
					'**************************************'

					'***** Incluye el link al WKF ******'

					If CLng(rsfav("fecvigen")) = 0 Then
						dFecha = "-"
					Else
						dFecha = CDate(CLng(rsfav("fecvigen")) + CDate("01-01-1900"))
					End If
					sli= sli & "<li class=documento-fav "& sFuncion & "><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "<div onclick='evtclick(event)' class='pill'>Fec. Vig.: "& fechadorCor(dfecha,"") &"</div></li> "

					rsfav.movenext
					totalR = totalR + 1
				Loop


			End If




			xcontenedor="<div id=pfav2 title='FAVORITOS'><div class='cant-fav'>"&totalR&" Documentos favoritos</div><div class='contenedor'>"
			' xcontenedor = xcontenedor & "<div data-cant='"& totalR &"' style='display:none' id='novedm"  & idsolapa &"'></div>"
			'ArmaContenido=scontenido & "<div data-cant='"& totalR &"' style='display:none' id='novedm"  & solapa &"'></div>"
			'  xcontenedor=xcontenedor &     "<div class='header'></div>"
			xcontenedor=xcontenedor &     "<div class='contenido'  >"
			xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "
			xcontenedor=xcontenedor &  "<ul>" & sli &"</ul></div>"
			'    xcontenedor=xcontenedor &     "</div>"
			Favoritos = xcontenedor

		End Function


		Function SUSCRIPCION()
			Dim XTablaN
			'Strsql=" Select *,B.ESTADO   From objmsg a "
			'STRSQL =STRSQL  & " LEFT JOIN OBJPRTY B ON(A.MSGOBJCOD= B.CODOBJETO) "
			'strsql=strsql & " where a.MSGTIPOMESS=2  AND usr = '" & SESSION("USUARIO") & "'"
			nFecini = 0
			totalR = 0
			If nFecini < CLng(CDate("01-01-1900")) Then
				nFecini = CLng(CDate("01-01-1900"))
			Else
				nFecini = 0
			End If

			dFechai = CLng(Date) - nFecini
			xsclase = " CLASS=ietx "
			xsclase2 = "class=zzzzz "
			StrSql = "SELECT *, cs.Expr6 as Ex6, cs.Expr7 as Ex7, cs.tobjdescri as temaobjdesc, cs.temadescri as temadescri, cs.stemadescri as subtemadescri " ' agrego campos cs de csrelpathdescriex CG 9-12
			StrSql = StrSql & "FROM csnovedades   a inner join objprty c on(a.codobjeto=c.codobjeto and c.estado=3) "
			StrSql = StrSql & "inner join objmsg b on( a.codobjeto=b.MSGOBJCOD and b.MSGTIPOMESS=2 and  b.usr = '" & SESSION("USUARIO") & "')"

			StrSql = StrSql & "inner join csrelpathdescriex cs on ( cs.objun=0" & SESSION("UN") &" and cs.cobj = b.MSGOBJCOD AND cs.tobj = b.msgobjtip)" 'agrego left join para traerme expr6, expr7, tobjdescri, temadescri, stemadescri
			'StrSql = StrSql & " Where a.fecinicio > " & dFechai - 1
			StrSql = StrSql & " and a.fecinicio <  " & dFechai + 6
			StrSql = StrSql & " And a.tipo = 3 "
			Strsql = StrSql & " ORDER BY 1 desc"
			'response.write strsql
			xTrColor = ""
			' vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID
			'response.write vuelta
			Set rsfav = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


			xsclase = " CLASS=ietx "
			xsclase2 = "class=istt "
			'''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

			XTablaN ="<div id='DTNOVF' style='width:100%;xbackground-color:#dddddd;' "
			XTablaN = XTablaN & "xborder-style:solid;border-width:3;border-color:#ffffff;' >"




			If tit= "S" Then
				'********* ENCABEZADO DE SECCIONES

				XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1     >"
				XTablaN = XTablaN & "<tr HEIGHT=25  >"
				''''XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;") &  "</TD>"
				XTablaN = XTablaN & "<td " & xsclase &"   >SUSCRIPCIONES</TD>"
				XTablaN = XTablaN & "</TR>"
				XTablaN = XTablaN & "</table>"

				XTablaN = XTablaN & "<table  width=98% cellspacing=0 cellpadding=0  >"

			Else
				XTablaN = XTablaN & "<table  width=98% cellspacing=0 cellpadding=0  >"
			End If

			'XTablaN = XTablaN & "<table   width=98% cellspacing=0 cellpadding=0  >"


			If rsfav.EOF Then


				xblink = ""

			Else

				xblink = " <script> var PgColor='#FF9B42'; function SetPgColor()"
				xblink = xblink & " {"
				xblink = xblink & " "

				xblink = xblink & " if (PgColor=='#FF9B42')"
				xblink = xblink & "{"
				xblink = xblink & "PgColor='#ffffff';document.getElementById('Xsus').style.backgroundColor='#ffffff';setTimeout('SetPgColor()', 500);"
				xblink = xblink & " }"
				xblink = xblink & " else"
				xblink = xblink & " {"
				xblink = xblink & " PgColor='#FF9B42';document.getElementById('Xsus').style.backgroundColor='#FF9B42';setTimeout('SetPgColor()', 500);"
				xblink = xblink & " } "

				xblink = xblink & " }"
				xblink = xblink & "setTimeout('SetPgColor()', 500);"
				xblink = xblink & "</script>"

				XTablaNdelim = ""  'XTablaNdelim & "</TR>"
				'     XTablaNdelim = XTablaNdelim & "<tr><td><hr>"
				'   XTablaNdelim = XTablaNdelim  & "</TD></tr>"


				Do While Not rsfav.EOF

					CLAVE = rsfav(0)
					If xTrColor = "#FFFFFF" Then
						xTrColor = "#eeeeff"
					Else
						xTrColor = "#FFFFFF"
					End If

					'	XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=25 >"
					sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsfav("MSGOBJCOD") & "&Tipo=" & rsfav("MSGOBJTIP") & "&MAINAPL=" & xMAINAPL
					sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
					sHref = "&nbsp;"

					sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)



					sTipobjeto = EsqnomPx(rsfav("MSGOBJTIP"), rsfav("MSGOBJCOD"), ofv)



					wmsgx = sTipobjeto & ""
					wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



					sImagen = wNomL

					If xarc = "NO" Then
						xarc = "&nbsp;"
					End If

					sLink = sImagen

					XTSLNKM = ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, wmsgx, "si")

					xtfechavig = fechaArg(CDate(rsfav(3) + CLng(CDate("01-01-1900"))))
					XTMSGSEND = "<B><b class='pill' font><font color='#AA0000' size=2 >&#0149;</font>" & xtfechavig & "</font></B>"
					'XTMSGSEND = XTMSGSEND '& "<B>< font color='#0000AA' size=2 >&#0124;</font></B>"
					wmsgx = Trim(rsfav(5) & "")

					If Len(wmsgx) < 5 Then
						wmsgx = EsqnomPx(rsfav(9), rsfav(8), ofv) '& " VIGENTE A PARTIR DEL " & fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900")))) & ". "
					Else
						'if xmenu=q3 then
						' wmsgx = UCase(wmsgx) ' & "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
						' else'response.write "aa" & len(wmsgx) & rsx(5)
						wmsgx = UCase(wmsgx)  '& "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
						' end if
					End If

					XTMSGSEND = XTMSGSEND & "<font >&nbsp;" & wmsgx & "</font>"
					If  session("APLMDCLI")="BGAL" Then
						XTSLNKM = XTSLNKM  '& XTMSGSEND
						XTSLNKM =  ofv.GenLink("", "nt title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, XTMSGSEND, "si")
					Else
						XTSLNKM = XTSLNKM & "<br><br>" '& XTMSGSEND
						XTSLNKM = "<br>" & ofv.GenLink("", "nt title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, XTMSGSEND, "si") '& "<br>"
					End If
					sImage = ofv.GenImage("", "", "", "", "Imagenes/Botdelx.jpg", "0", "")
					vlk = "../novedmenu.asp?tab=view3"
					sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=objmsg&volver=" & vlk
					sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

					CELDA = "" 'BotonEliminar("", sHref, "sus" &CLAVE)

					'sli = sli & "<li class=documento-link><i class=Ciclass></i>"  & XTSLNKM & "</li>"
					sli = sli & "<li class='documento-link'><i class='lar la-file-alt la-1x '></i><table border='0' style='width:100%; table-layout:fixed;'>"
					' Apertura de la tabla y fila de encabezados
					sli = sli & "  <thead>"
					'sli = sli & "    <tr style='background-color:" & Session("CMFPNTIT") & "; color: #fff;'>" ' Usa tu color de sesión
					sli = sli & "      <th class='th-noti' style='padding: 5px; width: 10%;'>Norma</th>"
					sli = sli & "      <th class='th-noti' style=' width: 20%;'>Descripcion Norma</th>"
					sli = sli & "      <th class='th-noti' style=' width: 20%;'>Descripcion Objeto</th>"
					sli = sli & "      <th class='th-noti' style=' width: 25%;'>Tema</th>"
					sli = sli & "      <th class='th-noti' style=' width: 25%;'>Subtema</th>"
					sli = sli & "    </tr>"
					sli = sli & "  </thead>"



					' Fila de datos
					sli = sli & "    <tr>"

					' Columna Ex6
					sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px solid #ddd;'>" & rsfav("Ex6") & "</td>"

					' Columna XTSKNKM
					sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px; vertical-align: middle; word-wrap: break-word; max-width: 200px;'>" & XTSLNKM & "</td>"

					' Columna tobjdescri
					sli = sli & "      <td style='padding: 5px; text-align: center	; border: 0px solid #ddd;'>" & rsfav("tobjdescri") & "</td>"

					' Columna temadescri
					sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px solid #ddd;'>" & rsfav("temadescri") & "</td>"

					' Columna subtemadescri
					sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px solid #ddd;'>" & rsfav("subtemadescri") & "</td>"

					sli = sli & "    </tr>"

					sli = sli & "</table></li>"

					rsfav.movenext
					totalR=totalR+1
				Loop


			End If




			'XTablaN = XTablaN & "</table></DIV>" & xblink




			'''SUSCRIPCIONES
			'SUSCRIPCION = XTablaN

			xcontenedor="<div id=psus data-cant2='" & totalR & "' title='SUSCRIPCIONES'><div class='contenedor'  >"
			'  xcontenedor=xcontenedor &     "<div class='header'></div>"
			xcontenedor=xcontenedor &     "<div class='contenido'  >"
			xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "
			xcontenedor=xcontenedor &  "<ul>" & sli & "</ul></div>"
			xcontenedor=xcontenedor &     "</div>"
			xcontenedor=xcontenedor &"  </div></div>"
			SUSCRIPCION = xcontenedor


		End Function

		Function ConvertYouTubeURL(url)
			Dim videoId, pos

			' Si es un enlace tipo watch?v=
			pos = InStr(url, "watch?v=")
			If pos > 0 Then
				videoId = Mid(url, pos + 8)
				If InStr(videoId, "&") > 0 Then
					videoId = Left(videoId, InStr(videoId, "&") - 1)
				End If
				ConvertYouTubeURL = "https://www.youtube.com/embed/" & videoId & "?enablejsapi=1"
				Exit Function
			End If

			' Si ya es embed (por si acaso)
			If InStr(url, "embed/") > 0 Then
				ConvertYouTubeURL = url
				Exit Function
			End If

			' Si es un short link tipo youtu.be
			pos = InStr(url, "youtu.be/")
			If pos > 0 Then
				videoId = Mid(url, pos + 9)
				If InStr(videoId, "?") > 0 Then
					videoId = Left(videoId, InStr(videoId, "?") - 1)
				End If
				ConvertYouTubeURL = "https://www.youtube.com/embed/" & videoId & "?enablejsapi=1"
				Exit Function
			End If

			' Si no matchea nada, devolver la original
			ConvertYouTubeURL = url
		End Function

		Function fechaArg(f)
			dia = Day(f)
			mes = Month(f)
			ano = Year(f)
			fechaArg = dia & "/" & mes & "/" & ano
		End Function
		response.write shtml
		%>



