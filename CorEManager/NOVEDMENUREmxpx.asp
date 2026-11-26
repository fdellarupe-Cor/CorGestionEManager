
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
	Dim sdocvencer



	'response.write bModificarvencer
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
	Shtml = shtml & "  "
	'Shtml = shtml & "  background-repeat: no-repeat;"
	'Shtml = shtml & "  background-size: cover;"
	Shtml = shtml & "}"

	Shtml = shtml & "</style>"

	Shtml = shtml & " <div id='videoDialog' style='display:none; padding:0; overflow:hidden;'></div> <form ACTION=UPDDATOS.ASP style='display:none' id=enviard name=enviard target=enviaf method=POST ><INPUT TYPE=TEXT VALUE='' ID=inforeg NAME=inforeg ></FORM><iframe  style='display:none' id=enviaf name=enviaf></iframe>"
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


	sHTML = sHTML & s &  ofv.botoneraESSX("0","","","","","","","","","","","","")



	sHTML = sHTML  & sdocvencer & autorizaciones() & SUSCRIPCION() & "</body></html>"
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

	StrSql = "SELECT objid,idsolapa,titulo,activo,cantreg,codEmp,ORDEN,WIDTH,HEIGHT,posx,posy,icono FROM novedtitex where (codemp=1 and idsolapa=7) or  (ACTIVO='S' AND codemp=1) order by orden "
	'	response.write strsql
	'	response.write StrSql
	Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
	btne=""

	If bModificarP="S" Then btne="<div class='toggle-drag-button'><input type='checkbox' id='editSwitch' title='Modo edición'></div>"
	str= btne & "<div class='grid-canvas' >"
	Do While Not rsx.eof

		'xconteledor="<table WIDTH=100% HEIGHT=100% ><tr><td HEIGHT=10% align=center >" & RSX("TITULO")& "</td></tr><tr><td HEIGHT=80% ></td></tr> <tr><td HEIGHT=10% >footer</td></tr></table>"
		If rsx("idsolapa")=2 Then
			xconteledor= "<div data-id='3333' class='contenedor2 rh'  sxtyle='top: " & rsx("posy") & "rem; left:"& rsx("posx")& "rem;'><div class='drag-handle'>ARRASTRAR AQUÍ</div>"
		ElseIf rsx("idsolapa")=7 Then
			Novedform = "CLIMN16317206150930885546227"
			ofv.ObtenerAtributosNOdo Novedform, "", "Tipo", "7"
			ConsultarDocvencer = ofv.AConsultar
			If  ConsultarDocvencer Then
				sdocvencer=DocumentosAvencer(rsx("cantreg"))

			Else
				sdocvencer=""
			End If

		Else
			xconteledor= "<div class='contenedor'  data-id='2222' vxstyle='top: " & rsx("posy") & "px; left:"& rsx("posx")& "px;'><div class='drag-handle'>ARRASTRAR AQUÍ</div>"
		End If
		Select Case rsx("idsolapa")

			Case 7
			Case Else
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
				xconteledor=xconteledor &     "<div class='header rh' title='" & RSX("TITULO") & " " & rsx("idsolapa") &"'>" & RSX("TITULO") & " " & xicono & " " & BurbujaN & "</div>"                  '& rsx("idsolapa") &"</div>"
				If rsx("idsolapa")=11 Or rsx("idsolapa") =5 Or rsx("idsolapa") =8 Or rsx("idsolapa") =6 Then
					'xconteledor=xconteledor &     "<div class='cant-fav'>" & rsx("cantreg") & " Documentos</div>" ''' para saber la cantidad de formularios que tienen cada tarjeta 26/12
				End If
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

			'str=str &   "<div  style='top: " &RSX("posy") & "px;left:" & RSX("posx") & "px;height:" & RSX("HEIGHT") & "px; width:"& RSX("WIDTH") &"px'" & "data-height=" & RSX("HEIGHT") & " data-width=" & RSX("WIDTH") & " data-left=" & rsx("posx") & " data-top=" & RSX("posy") &" data-id=" & rsx("objid") & rsx("cantreg") &" class='grid-item'>" & xconteledor &"</div>"
			str=str &   "<div  style='top: " &RSX("posy") & "rem;left:" & RSX("posx") & "rem;height:" & RSX("HEIGHT") & "rem; width:"& RSX("WIDTH") &"rem;' " & "data-height=" & RSX("HEIGHT") & " data-width=" & RSX("WIDTH") & " data-left=" & rsx("posx") & " data-top=" & RSX("posy") &" data-id=" & rsx("objid")  &" class='grid-item'>" & xconteledor &"</div>"

		End Select

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
		'response.write strsql & " 333"
		Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
		If  rsx.EOF Then
			scontenido=""
		Else
			scontenido=" <div class='swiper mySwiper'>"
			scontenido= scontenido & " <div class='swiper-wrapper'>"
			'	<div class="swiper-slide">
			'scontenido="<ul class=slider>"
			'scontenido= scontenido & "<div class='swiper-slide video-container'>"
			'scontenido= scontenido & "       <iframe id=ixs name=ixs height=100% width=100% src='https://www.youtube.com/embed/hga-hbnw66I?enablejsapi=1' controls ></iframe>"
			'scontenido= scontenido & "</div>"

			Do While Not rsx.EOF


				idyoutube="NO"

				Select Case rsx("msgimagen")
					Case 29999
					' scontenido= scontenido & "<div class='swiper-slide video-container'>"
					'scontenido= scontenido & "       <iframe id='ixs" & rsx(0) & "' name='ixs" & rsx(0) & "' height=100% width=100% src='"& rsx("msgurldsc") & "?enablejsapi=1' controls ></iframe>"
					' scontenido= scontenido & "</div>"
					' scontenido= scontenido & "<div class='swiper-slide video-launcher' data-video-id='lhyxKBQDjEc'>"
					' scontenido= scontenido & "    <img width=100% height=100% src='https://img.youtube.com/vi/xlmMySVdhkM/0.jpg' alt='Video 1'>"
					' scontenido= scontenido &"     <div class='play-button'>play</div>"
					'scontenido= scontenido & "</div>"
					Case Else
					If rsx("imagen") <> "NO" Or rsx("msgimagen")=2 Then
						If rsx("msgimagen")=2 Then
							idyoutube=obtenerYoutubeID(rsx("msgurldsc"))
							imgtube="https://img.youtube.com/vi/" & idyoutube & "/0.jpg"
							' response.write imgtube
							scontenido= scontenido & " <div rsid='"& rsx("objid") & "' data-video-id='" & idyoutube & "' class='swiper-slide' STYLE=""background-repeat: no-repeat;background-size: cover;background-position: center;xbackground-size: 100% 100%;background-image: url('"& imgtube & "');z-index:1;margin-right: 30px;display: flex;justify-content: center;flex-direction: column;align-items: flex-start;gap: 10px;padding-left: 2rem;padding-right: 2rem;font-family:" & Session("CMFPNTITFTFC") & ";"">"
							'scontenido= scontenido & "    <i class='las la-play' ></i><iframe  frameborder='0' allow='autoplay; encrypted-media; '   style='display:none' id='" & rsx("objid") & idyoutube & "' name="& rsx("objid") &idyoutube& "></iframe>"
						Else
							scontenido= scontenido & " <div class='swiper-slide' STYLE=""background-size: cover;background-position: center;background-repeat: no-repeat;xbackground-size: 100% 100%;background-image: url('"& rsx("imagen") & "');z-index:1;margin-right: 30px;display: flex;justify-content: center;flex-direction: column;align-items: flex-start;gap: 10px;padding-left: 2rem;padding-right: 2rem;font-family:" & Session("CMFPNTITFTFC") & ";"">"
						End If

						'scontenido= scontenido & "<button class='btn-ver-mas' onclick='openPopup()'>Ver más</button>"
					Else
						scontenido= scontenido & " <div class='swiper-slide' STYLE='z-index:1;background-color:" & Session("CMIETT") & ";font-family:" & Session("CMFPNTITFTFC") & ";margin-right: 30px;display: flex;justify-content: center;flex-direction: column;align-items: flex-start;gap: 10px;padding-left: 2rem;padding-right: 2rem;'>"
					End If


					scontenido= scontenido & "     <div class='title' data-swiper-parallax='-300'>"
					If rsx("msgimagen")=2 And Trim(rsx("msgurldsc")) & "" <> "" Then
						scontenido= scontenido & "<div class='pill-video' ><i class='lab la-youtube' style='font-weight: 400 !important;'></i> VIDEO</div>"
					End If
					scontenido= scontenido &  rsx("msgtitulo")
					If CLng(rsx("fecinicio")) = 0 Then
						dFecha = "-"
					Else
						dFecha = CDate(CLng(rsx("fecinicio")) + CDate("01-01-1900"))
					End If
					scontenido= scontenido & "     <div class='fecvigen' data-swiper-parallax='-300'><i class='las la-calendar'></i>&nbsp;Fecha de Publicación: " & fechadorCor(dfecha,"") & "</div>"

					scontenido= scontenido & "</div>"
					scontenido= scontenido & "    <div class='subtitle' data-swiper-parallax='-200'><font>" & rsx("msgcopete") &"</font>"
					If Trim(rsx("mensaje")&"") <> ""  Then
						scontenido= scontenido & "            <div xdata-video-id='" & rsx("msgurldsc") & "' data-idN='"& rsx("codigo") &"'  class='Nvermas' style=color:" & Session("CMFPNBOT") & "> Ver más "
						scontenido= scontenido & " <div data-video-id='" & idyoutube & "' class=NovVer name='" & rsx("codigo") & "' id=" & rsx("codigo") &" >" & ArmaVerNovedad(rsx,"S")&  "</div>"
						scontenido= scontenido & "            </div>"
					ElseIf idyoutube <> "" And Trim(rsx("mensaje")&"") = "" Then 'video sin text
						If Trim(rsx("msgtitulo")&"") = ""  And Trim(rsx("msgcopete")&"") = "" And rsx("msgimagen")=1 Then
							scontenido= scontenido & " <div xdata-video-id='" & rsx("msgurldsc") & "' data-idN='"& rsx("codigo") &"'  class='Nvermas' style=color:" & Session("CMFPNBOT") & ">"
                            sarmacontenidover=ArmaVerNovedad(rsx,"N")
						Else
							scontenido= scontenido & " <div xdata-video-id='" & rsx("msgurldsc") & "' data-idN='"& rsx("codigo") &"'  class='Nvermas' style=color:" & Session("CMFPNBOT") & "> Ver más "
						    sarmacontenidover=ArmaVerNovedad(rsx,"S")
						End If
						scontenido= scontenido & " <div data-video-id='" & idyoutube & "' class=NovVer name='" & rsx("codigo") & "' id=" & rsx("codigo") &" >" & sarmacontenidover &  "</div>"
						scontenido= scontenido & "</div>"
					End If

					scontenido= scontenido & "</div>"
					' if rsx("imagen") <> "NO" THEN
					' scontenido= scontenido & "     <div class='text' STYLE=""background-image: url('"& rsx("imagen") & "');""  data-swiper-parallax='-100'>"
					' ELSE
					scontenido= scontenido & "     <div class='xtext' data-swiper-parallax='-100'>"
					' END IF
					scontenido= scontenido & "            <div class=xpnoved>"
					If rsx("imagen") <> "NO" Then
						' scontenido= scontenido & replace(rsx("mensaje"),chr(34),"")

					End If
					scontenido= scontenido & "            </div>"

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

				End Select
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




		Case 3   ', 9

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
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'><ul >"

			Do While Not rsx.EOF
				xMAINAPL=rsx("aplic")
				sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&MAINAPL=" & xMAINAPL
				sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
				sHref = ""


				sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
				XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPxex(rsx(9), rsx(8), ofv,""), "si")
				'scontenido= scontenido & "<li class=documento-publi "& sFuncion & "><i class='lar la-file-alt la-1x'></i>"  & XTSLNKM & "</li>"
				cardborder=""
				If Trim(rsx("mensaje") &"") <> "" Then
					cardborder="border-radius: 8px 8px 0px 0px"
				End If
				scontenido= scontenido & "<li class=lista-docs><div class=documento-cards style='" & cardborder & "' >" & armatarjeta(rsx,3)  & "</div>"
				SQL="SELECT * FROM OBJINFO A "
				SQL=SQL & " LEFT JOIN OBJINFOIMG B ON (B.CODOBJETO=A.CODOBJETO)"
				SQL=SQL & "WHERE A.CODOBJETO=" & RSX("CODOBJETO")
				'response.write sql
				Set rsxINFO = ofv.crearconsultaEx(SQL, cn, 1, parametros)
				If Not RSXINFO.EOF Then


					IMGINFO=RSXINFO("OBJIMAGEN")& ""
					If  IMGINFO = "" Then
						scontenido= scontenido & "<div class=publi-card >"
					Else
						scontenido= scontenido & "<div class=publi-card ><div class=publi-imagenovideo><img src='" &IMGINFO &"' ></div>"
					End If

					scontenido= scontenido & "<div class=publi-contenido ><span class=publi-titulo >" & RSXINFO("OBJTITULO") &"</span>"
					scontenido= scontenido & "<p class=publi-cuerpo >"& RSXINFO("OBJMENSAJE") & "</p>"
					scontenido= scontenido & "</div></div>"
				End If
				Call OFV.CERRARCONSULTA(RSXINFO)
				If Trim(rsx("mensaje") &"") <> "" Then
					scontenido= scontenido & "<div class=documento-cards-footer><span  class=notas-card ><i style='font-size:1rem;margin-right:5px;vertical-align:middle;' class='las la-thumbtack'></i>" & rsx("mensaje")& "</span></div>"
				End If


				scontenido= scontenido & "</li>" 'armatarjeta(id,tipo,cn,nove)
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
		StrSql = StrSql & "SELECT top " & cantregN & " b1.CODOBJETO, max(b1.TIPOBJETO) as tipobjeto, max(finicio) as fecha  "
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
		'response. write StrSql
		Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


		If  rsx.EOF Then
			scontenido=""
		Else
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'><ul >"
			totalR=0
			Do While Not rsx.EOF


				'scontenido= scontenido & "<li class=documento-ultcons><div class='fondo-icono'><i class='lar la-file-alt la-1x'></i></div>"  & XTSLNKM & "</li>"
				scontenido= scontenido & "<li class=lista-docs ><div class=documento-cards >" & armatarjeta(rsx,noveid)  & "</div>"
				'if trim(rsx("mensaje") &"") <> "" then
				' scontenido= scontenido & "<div class=documento-cards-footer><span class=notas-card ><i class='las la-sticky-note'></i>TOME NOTA: DEBE FIRMAR ESTE ARCHIVO DE FORMA INMEDIATA EN LA ZONA DE FIRMAS</span></div>"
				'end if
				scontenido= scontenido & "</li>" 'armatarjeta(id,tipo,cn,nove)
				rsx.movenext
				totalR=totalR + 1

			Loop
			scontenido= scontenido & "</ul></div>"
		End If
		Call ofv.cerrarconsulta(rsx)
		Case 6

		StrSql = " Select top " & cantregN & " *,B.ESTADO,B.tipobjeto as tipobjeto,B.fecvigen as fecvigen,b.codobjeto,ct.aplic  From objmsg a " '
		StrSql = StrSql & " LEFT JOIN OBJPRTY B ON(A.MSGOBJCOD= B.CODOBJETO AND B.tipobjeto = A.msgobjtip ) " 'agrego AND B.tipobjeto = A.msgobjtip
		StrSql = StrSql & " inner JOIN objalogdet C ON (C.OBJID=a.objid) "
		StrSql = StrSql & " inner JOIN tipobjeto Ct ON (Ct.tipobjeto=b.tipobjeto) "

		StrSql = StrSql & " where a.MSGTIPOMESS=1 and b.estado=3 AND usr = '" & SESSION("USUARIO") & "'"
		'	response.write strsql
		xTrColor = ""
		Set rsfav = ofv.crearconsultaEx(StrSql, cn, 1, parametros)




		If  rsfav.EOF  Then



		Else
			totalR=0
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'><ul >"
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
				'scontenido= scontenido& "<li class=documento-fav "& sFuncion & "><div class='fondo-icono'><i class='lar la-file-alt la-1x'></i></div>"  & XTSLNKM & "<div onclick='evtclick(event)' class='pill'>Fec. Vig.: "& fechadorCor(dfecha,"") &"</div></li> "
				'scontenido= scontenido & "<li class=documento-fav "& sFuncion & "><div class='fondo-icono' ><i class='lar la-file-alt la-1x'></i></div>" & XTSLNKM & "<div onclick='evtclick(event)' class='tooltip-container' data-tooltip='Fec. Vig.: "& fechadorCor(dfecha,"") &"'><i class='las la-info-circle'></i></div></li>"
				scontenido= scontenido & "<li class=lista-docs ><div class=documento-cards >" & armatarjeta(rsfav,6)  & "</div></li>" 'armatarjeta(id,tipo,cn,nove)
				rsfav.movenext
				totalR = totalR + 1
			Loop
			'scontenido = scontenido & "<div class='pill'>"&totalR&" Documentos favoritos</div>"
			scontenido= scontenido & "</ul></div>"

		End If

		'scontenido=Favoritos(cantregN)
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
		Case 799
		'doc a vencer

		If session("UN") = 1 Then
			StrSql = ""
			StrSql = StrSql & "SELECT  a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente,a.aplic"
			StrSql = StrSql & " from objetos a "
			StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
			StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
			StrSql = StrSql & " where  a.objun=1 and b.fecvto > 0 order by b.fecvto "
		Else
			StrSql = ""
			StrSql = StrSql & "SELECT  a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
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
			CONT=1
			Do While Not rsx.EOF
				dfecvto = rsx(2)
				dfecvto = CLng(dfecvto)
				'RESPONSE.WRITE dfecvto
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
							CONT=CONT +1

						End If
					End If
				End If
				rsx.movenext

				If CONT > CANTREGN Then Exit Do

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
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
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

		StrSql = "SELECT distinct top " & cantregN & " a.codobjeto,a.tipobjeto,c.aplic FROM OBJPRTY a "
		StrSql = StrSql & " INNER JOIN  OBJOBSCMT B ON (b.codsol=9999 and B.tipo='CMD' AND B.ACTPEND is null  AND B.CODOBJETO=a.CODOBJETO AND B.TIPOBJETO=a.TIPOBJETO ) "
		StrSql = StrSql & " INNER JOIN  TIPOBJETO C ON (C.COMENTARIOS='S'  AND A.TIPOBJETO=C.TIPOBJETO ) "
		StrSql = StrSql & " WHERE  (a.ESTADO = 3 )"
		StrSql = StrSql & " And (a.autor='"& SESSION("usuario") &"')"
		'response.write strsql

		Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


		If  rsx.EOF Then
			scontenido=""
		Else
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
			Do While Not rsx.EOF
				sActcon = "ConsultarcMD.asp?PEND=S&Codigo=" & rsx(0) & "&Tipo=" & rsx(1)
				sHref = "&nbsp;"


				sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
				XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")
				'scontenido= scontenido & "<li class=documento-com><div class='fondo-icono'><i class='lar la-file-alt la-1x'></i></div>"  & XTSLNKM & "</li>"
				scontenido= scontenido & "<li class=lista-docs ><div class=documento-cards >" & armatarjeta(rsx,8)  & "</div></li>" 'armatarjeta(id,tipo,cn,nove)
				rsx.movenext

			Loop
			scontenido= scontenido & "</ul></div>"
		End If
		Call ofv.cerrarconsulta(rsx)

		Case 9 ' linea de negocios

		sql = "select distinct a.objid,a.descri,a.secuen from tipoproducto a"
		sql= sql & " inner join objetos b on (a.objid = b.tsegmento)  "
		sql= sql & " inner join objprty c on (b.codobjeto = c.codobjeto and b.tipobjeto=c.tipobjeto and c.estado=3)"
		sql= sql & " order by a.secuen"
		'sql= sql & "
		'response.write sql
		Set rsx = ofv.crearconsultaEx(sql,cn,1,parametros)
		If  rsx.EOF Then
			scontenido=""
		Else
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
			Do While Not rsx.EOF
				parametros = "MENU=PORLINEGO&NOMBRE=LINEA&NDAT=ESTADO&NVAL=2&seg=CLIMN4601201011384966908594&MNN=POR_LINEA_DE_NEGOCIO&id=" & rsx(0)

				sActcon = "SUBCONSULTA.asp?" & parametros
				sHref = "&nbsp;"


				sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
				XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, rsx(1) , "si")
				scontenido= scontenido & "<li " & xsFuncion & " class=lista-docs>" & "<a CLASS=nti  href='SUBCONSULTA.asp?" & parametros & "'> -" & rsx(1) & "</a>"  & "</li>"
				rsx.movenext

			Loop
			scontenido= scontenido & "</ul></div>"
		End If

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
		StrSql = StrSql & "SELECT top " & cantregN & " b1.CODOBJETO, max(b1.TIPOBJETO) as tipobjeto, max(finicio) as fecha  " 'agregar b.fecvigen as fecvig,
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
			scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'> <ul >"
			totalR=0
			Do While Not rsx.EOF
				xMAINAPL=rsx("aplic")
				sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
				sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
				sHref = ""


				sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
				XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, EsqnomPx(rsx(1), rsx(0), ofv), "si")
				'scontenido= scontenido & "<li class=documento-ultcons><div class='fondo-icono'><i class='lar la-file-alt la-1x'></i></div>"  & XTSLNKM & "</li>"
				scontenido= scontenido & "<li class=lista-docs ><div class=documento-cards >" & armatarjeta(rsx,11)  & "</div></li>" 'armatarjeta(id,tipo,cn,nove)
				rsx.movenext
				totalR=totalR + 1
			Loop
			scontenido= scontenido & "</ul></div>"
		End If
		Call ofv.cerrarconsulta(rsx)
		Case 12
		scontenido="<div style='padding:0px 5px;overflow-x: auto;overflow-y: auto;height:100%;width:100%'>"
		scontenido=scontenido & indice2()&"</div>"

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
				'sxlink = sxlink & "<i class='las la-book'></i>"
				sxlink = "ObjTreeSubTemaCons.asp?formato=N&"
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


				celda= ofv.GenCelda(""," xbgcolor='" & xTrColor & "' xstyle='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' ","","10%","","","",sImage & slink2,"si")

				Fila    = Fila & ofv.GenRow(""," ","","","20","","",Celda,"si")

			End If
			RS.MOVENEXT
		Loop

	End If






	Call OFV.cerrarconsulta(rs)




	INDICE2 = ofv.GenTabla("","aaab STYLE=PADDING-TOP:3PX; ","","100% ","","0","","0","0","","",Fila,"si")
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


			sli = sli & " <li style='list-style: none; margin-bottom: 10px;'>"
			sli = sli & armawkf(rsx)
			sli = sli & " </li>"

			rsx.movenext
			cant=cant +1
		Loop


	End If
	Call ofv.cerrarconsulta(rsx)




	xcontenedor="<div id=pfav data-cant='" & cant& "' title='Mis Autorizaciones'><div class='contenedor'  >"
	'  xcontenedor=xcontenedor &     "<div class='header'></div>"
	xcontenedor=xcontenedor &     "<div class='contenido'  >"
	xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "

	If cant=0 Then
		xcontenedor=xcontenedor & "<div class=Sininfo align='center' ><i class='las la-broadcast-tower' style='margin-right: 0.3rem;' ></i>Sin informacion</div>"
	Else
		xcontenedor=xcontenedor &  "<ul>" & sli &"</ul>"
	End If

	xcontenedor=xcontenedor &     "</div>"
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
			sli= sli & "<li class=documento-fav "& sFuncion & ">"  & XTSLNKM & "<div onclick='evtclick(event)' class='pill'>Fec. Vig.: "& fechadorCor(dfecha,"") &"</div></li> "

			rsfav.movenext
			totalR = totalR + 1
		Loop


	End If


	'xcontenedor="<div id=pfav2 data-cant='" & cant& "' title='FAVORITOS'><div class='contenedor'>"
	xcontenedor = "<div id='pfav2' title='FAVORITOS' style='overflow:hidden; padding: 0 10px;'><span style='float:left; font-weight:bold;'>FAVORITOS</span><span class='pill' style='float:right;'>" & totalR & " Documentos favoritos</span></div><div class='contenedor'>"
	xcontenedor=xcontenedor &     "<div class='contenido'  >"
	xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "
	xcontenedor=xcontenedor &  "<ul>" & sli &"</ul></div>"
	xcontenedor=xcontenedor &     "</div></div>"
	Favoritos = xcontenedor

End Function


Function DocumentosAvencer(cantregN)

	If session("UN") = 1 Then
		StrSql = ""
		StrSql = StrSql & "SELECT  a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente,a.aplic"
		StrSql = StrSql & " from objetos a "
		StrSql = StrSql & " inner join objprty b on (a.codobjeto = b.codobjeto and b.estado=3) "
		StrSql = StrSql & " inner join clienteplan c on (a.objun=c.id) "
		StrSql = StrSql & " where  a.objun=1 and b.fecvto > 0 order by b.fecvto "
	Else
		StrSql = ""
		StrSql = StrSql & "SELECT  a.codobjeto, a.tipobjeto, b.fecvto, a.aplic, a.objun, c.cliente "
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
		CONT=0
		Do While Not rsx.EOF
			dfecvto = rsx(2)
			dfecvto = CLng(dfecvto)
			'RESPONSE.WRITE dfecvto
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


						If CLng(dfecvto) = 0 Then
							dFecha = "-"
						Else
							dFecha = CDate(CLng(dfecvto) + CDate("01-01-1900"))
						End If
						scontenido= scontenido & "<li class=documento-vencer "& sFuncion & "><div class=noved-doc-icon-vencer><i class='lar la-file-alt la-1x'></i></div>"  & XTSLNKM & "<span class='noved-doc-tag-vencer'>Vencimiento de revisión: "& fechadorCor(dfecha,"")  &"</span></li>"
						CONT=CONT +1

					End If
				End If
				'End If
			End If
			rsx.movenext

			If CONT >= CANTREGN Then Exit Do

		Loop

	End If



	scontenido= scontenido & "</ul></div>"
	xcontenedor="<div id=pvencer data-cant2='" & CONT & "' title='Documentos a vencer'><div class='contenedor'  >"
	'  xcontenedor=xcontenedor &     "<div class='header'></div>"
	xcontenedor=xcontenedor &     "<div class='contenido'  >"
	'xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "
	'xcontenedor=xcontenedor &  "<ul>" & sli & "</ul></div>"
	If CONT=0 Then
		xcontenedor=xcontenedor & "<div class=Sininfo align='center' ><i class='las la-broadcast-tower' style='margin-right: 0.3rem;' ></i>No tienes documentos próximos a vencer.</div>"
	Else
		xcontenedor=xcontenedor & scontenido
	End If
	xcontenedor=xcontenedor &     "</div>"
	xcontenedor=xcontenedor &"  </div></div>"

	DocumentosAvencer=xcontenedor
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



			wmsgx = sTipobjeto & " "
			wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



			sImagen = wNomL

			If xarc = "NO" Then
				xarc = "&nbsp;"
			End If

			sLink = sImagen

			XTSLNKM = ofv.GenLink("", "nt  title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, wmsgx, "si")

			xtfechavig =   fechaArg(CDate(rsfav(3) + CLng(CDate("01-01-1900"))))
			'XTMSGSEND = "<B><b class='pill' font><font color='#AA0000' size=2 >&#0149;</font>" & xtfechavig & "</font></B>" 'comentado para traer sólo la desc 2/1/26  CG
			'XTMSGSEND = XTMSGSEND '& "<B>< font color='#0000AA' size=2 >&#0124;</font></B>"
			wmsgx = Trim(rsfav(5) & "")

			If Len(wmsgx) < 5 Then
				wmsgx = EsqnomPx(rsfav(9), rsfav(8), ofv) '& " VIGENTE A PARTIR DEL " & fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900")))) & ". "
			Else
				'if xmenu=q3 then
				' wmsgx = UCase(wmsgx) ' & "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
				' else'response.write "aa" & len(wmsgx) & rsx(5)
				wmsgx = UCase(wmsgx)'& ""& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
				' end if
			End If

			XTMSGSEND = "<font >" & wmsgx & "</font>"
			wmsgx=""
			If  session("APLMDCLI")="BGAL" Then
				XTSLNKM = XTSLNKM  '& XTMSGSEND
				XTSLNKM =  ofv.GenLink("", "nt title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, XTMSGSEND, "si")
			Else
				XTSLNKM = XTSLNKM & "<br><br>" '& XTMSGSEND
				XTSLNKM = "" & ofv.GenLink("", "nt title='" & TITULOLNK & "' ", sHref, "", "", sFuncion, XTMSGSEND, "si") '& "<br>"
			End If
			sImage = ofv.GenImage("", "", "", "", "Imagenes/Botdelx.jpg", "0", "")
			vlk = "../novedmenu.asp?tab=view3"
			sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=objmsg&volver=" & vlk
			sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

			CELDA = "" 'BotonEliminar("", sHref, "sus" &CLAVE)

			'sli = sli & "<li class=documento-link><i class=Ciclass></i>"  & XTSLNKM & "</li>"
			'sli = sli & "<li class='documento-link'><i class='lar la-file-alt la-2x '></i><table border='0' style='width:90%; table-layout:fixed;margin:15px;'>"
			' Apertura de la tabla y fila de encabezados
			'sli = sli & "  <thead>"
			'sli = sli & "    <tr style='background-color:" & Session("CMFPNTIT") & "; color: #fff;'>" ' Usa tu color de sesión
			'sli = sli & "      <th class='th-noti' style='padding left: 5px; width: 10%;'>Norma</th>"
			'sli = sli & "      <th class='th-noti' style=' width: 20%;'>Descripcion Norma</th>"
			'sli = sli & "      <th class='th-noti' style=' width: 20%;'>Descripcion Objeto</th>"
			'sli = sli & "      <th class='th-noti' style=' width: 20%;'>Tema</th>"
			'sli = sli & "      <th class='th-noti' style=' width: 20%;'>Subtema</th>"
			'sli = sli & "    </tr>"
			'sli = sli & "  </thead>"



			' Fila de datos
			'sli = sli & "    <tr>"

			' Columna Ex6
			'sli = sli & "      <td style='padding left: 5px; text-align: center; border: 0px solid #ddd;'>" & rsfav("Ex6") & "</td>"

			' Columna XTSKNKM
			'sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px; vertical-align: middle; word-wrap: break-word; max-width: 200px;'>" & XTSLNKM & "</td>"

			' Columna tobjdescri
			'sli = sli & "      <td style='padding: 5px; text-align: center	; border: 0px solid #ddd;'>" & rsfav("tobjdescri") & "</td>"

			' Columna temadescri
			'sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px solid #ddd;'>" & rsfav("temadescri") & "</td>"

			' Columna subtemadescri
			'sli = sli & "      <td style='padding: 5px; text-align: center; border: 0px solid #ddd;'>" & rsfav("subtemadescri") & "</td>"

			'sli = sli & "    </tr>"

			'sli = sli & "</table></li>"

			''''''''''''''''''Version en divs de suscripciones
			' Fila de datos convertida a Tarjeta
			'sli = sli & " <li style='list-style: none; margin-bottom: 10px;'>"
			' El contenedor principal usa space-between para separar texto de botón
			sli = sli & "   <div class='susc-item'  " & sFuncion & ">"

			' --- GRUPO IZQUIERDO: Todo el texto ---
			sli = sli & "     <div style='display: flex; flex-direction: column; gap: 6px;'>"

			' Descripción principal

			sli = sli & "       <div>" & XTSLNKM & "</div>"

			' Título
			'sli = sli & "       <span class='authb-pill' style='font-weight: bold;margin-left:5px;'>" & rsfav("Ex6") & "</span>"


			' Tags (Tema / Subtema)
			sli = sli & "       <div style='display: flex; align-items: center; gap: 10px; margin-top: 4px;margin-left:5px;'>"
			sli = sli & "         <span class='noved-doc-path'><i class='la la-folder'> </i> <span class='noved-doc-path'>" & rsfav("temadescri") & "</span></span>"
			sli = sli & "		  <span class='separator'>/</span>"
			sli = sli & "         <span class='noved-doc-path'>" & rsfav("subtemadescri") & "</span>"
			sli = sli & "       </div>"

			sli = sli & "     </div>"

			' --- GRUPO DERECHO: El botón ---
			' Ajustamos el estilo para que parezca un botón de acción claro
			sli = sli & "     <div style='flex-shrink: 0;'>"
			sli = sli & "<span class='noved-doc-tag'>#"& rsfav("Ex7")&"</span>"
			sli = sli & "<span class='noved-doc-tag'># "&rsfav("temaobjdesc")&"</span>" '''MSGOBJTIP
			'sli = sli & "       <button class='bietx' type='button'>Quitar suscripcion</button>"
			sli = sli & "     </div>"

			sli = sli & "   </div>"
			sli = sli & " </li>"
			rsfav.movenext
			totalR=totalR+1
		Loop


	End If




	'XTablaN = XTablaN & "</table></DIV>" & xblink




	'''SUSCRIPCIONES
	'SUSCRIPCION = XTablaN

	xcontenedor="<div id=psus data-cant2='" & totalR & "' title='Suscripciones'><div class='contenedor'  >"
	'  xcontenedor=xcontenedor &     "<div class='header'></div>"
	xcontenedor=xcontenedor &     "<div class='contenido'  >"
	xcontenedor=xcontenedor & "<div style='xgap:10px ;overflow-x: auto;overflow-y: auto;height:400px;width:100%'>  "

	If totalR=0 Then
		xcontenedor=xcontenedor & "<div class=Sininfo align='center' ><i class='las la-broadcast-tower' style='margin-right: 0.3rem;' ></i>No se registran novedades en sus documentos suscritos.</div>"
	Else
		xcontenedor=xcontenedor &  "<ul>" & sli &"</ul>"
	End If
	'	xcontenedor=xcontenedor &  "<ul>" & sli & "</ul></div>"
	xcontenedor=xcontenedor &  "</div>"
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





Function armatarjeta(rsdoc,noveid)
	xMAINAPL=rsdoc("aplic")
	If noveid=8 Then
		sActcon = "ConsultarcMD.asp?PEND=S&Codigo=" & rsdoc("codobjeto") & "&Tipo=" & rsdoc("tipobjeto")
	Else
		sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsdoc("codobjeto") & "&Tipo=" & rsdoc("tipobjeto") & "&MAINAPL=" & xMAINAPL
		sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
	End If
	sHref = ""

	sqlD="select temadescri,stemadescri,tobjdescri,fecvigen,fecvto,b.autor as Autor from CSRELPATHDESCRIex a "
	sqld=sqld  & " inner join objprty b on(b.codobjeto=a.cobj and b.tipobjeto=a.tobj) "
	sqld=sqld & " where a.cobj=0" & rsdoc("codobjeto") & " and a.tobj='" & rsdoc("tipobjeto") &  "'"
	'if noveid=6 then response.write sqld

	Set rsdocInf = ofv.crearconsultaEx(sqld, cn, 1, parametros)
	If Not rsdocInf.eof Then
		stipod=rsdocInf("tobjdescri")
		stipots=rsdocInf("temadescri") & "/" & rsdocInf("stemadescri")
		fecvigen=rsdocInf("fecvigen")
		fecvto=rsdocInf("fecvto")
		Autor= rsdocInf("Autor")
		xxxx = getInfoUser(Autor)
		Autor = SESSION("SGU_nombre")
	Else
		stipod="&nbsp;"
		stipots="&nbsp;"
		fecvto=0
		fecvigen=0
		Autor=0
	End If
	sFuncion = " onclick=" & Chr(34) & "AbrirVentana('" & sActcon & "'); return false;" & Chr(34)
	sdoc= EsqnomPxEx(rsdoc("tipobjeto"), rsdoc("codobjeto"), ofv,"1")
	XTSLNKM = ofv.GenLink("", "nt  title='" & sdoc & "' ", sHref, "", "", sFuncion,sdoc, "si")

	'armatarjeta="<div class='noved-doc-card'>"


	'armatarjeta=armatarjeta & "  <div class='noved-doc-item' data-doc-type='procedimiento'>"

	armatarjeta=armatarjeta & "    <div class='noved-doc-icon'>"
	armatarjeta=armatarjeta & "      <i class='la la-file-alt'></i>"
	armatarjeta=armatarjeta & "    </div>"

	armatarjeta=armatarjeta & "    <div class='noved-doc-body' " & sFuncion & ">"

	armatarjeta=armatarjeta & "      <div class='noved-doc-head'>"
	armatarjeta=armatarjeta & "        <div class='noved-doc-title' >"
	armatarjeta=armatarjeta &          XTSLNKM
	armatarjeta=armatarjeta & "        </div>"

	armatarjeta=armatarjeta & "        <span class='noved-doc-info'   onclick=event.stopPropagation(); "
	fecvigen=fechadorcor(CDate(CLng("0" & fecvigen + CLng(CDate("01-01-1900")))), "")
	fecvto=fechadorcor(CDate(CLng("0" & fecvto + CLng(CDate("01-01-1900")))), "")
	'armatarjeta=armatarjeta & "              data-info='Fecha de vigencia: " & fecvigen &" Fecha de vencimiento: " & fecvto & "'>"
	armatarjeta=armatarjeta & "              data-info='Fecha de vigencia: " & fecvigen &" Fecha vto. de Rev:  " & fecvto & " Autor: " & Autor & "'>"
	'armatarjeta=armatarjeta & "  data-info=' · Fecha de vencimiento :" & fecvto & "'>"
	armatarjeta=armatarjeta & "          <i class='la la-info-circle'></i>"
	armatarjeta=armatarjeta & "        </span>"
	armatarjeta=armatarjeta & "      </div>"

	armatarjeta=armatarjeta & "      <div class='noved-doc-meta'>"
	armatarjeta=armatarjeta & "        <span class='noved-doc-tag'>" &stipod &"</span>"

	armatarjeta=armatarjeta & "        <span class='noved-doc-path'"
	armatarjeta=armatarjeta & "              title='" & stipots& "'>"
	armatarjeta=armatarjeta & "          <i class='la la-folder' style='margin-right: 2px;'></i>"
	armatarjeta=armatarjeta & stipots
	armatarjeta=armatarjeta & "        </span>"
	armatarjeta=armatarjeta & "      </div>"

	armatarjeta=armatarjeta & "    </div>"

	'armatarjeta=armatarjeta & " </div>"

	'armatarjeta=armatarjeta & "</div>"

End Function

Function obtenerYoutubeID(url)
	Dim regEx, matches
	Dim result
	result = "" ' Valor por defecto '

	' Validar que la URL no esté vacía '
	If url <> "" Then
		Set regEx = New RegExp
		' Patrón actualizado que incluye shorts y live '
		regEx.Pattern = "(?:youtu\.be\/|youtube\.com\/(?:v\/|u\/\w\/|embed\/|watch\?v=|&v=|shorts\/|live\/))([^#&?\s]{11})"
		regEx.IgnoreCase = True
		regEx.Global = False

		Set matches = regEx.Execute(url)

		If matches.Count > 0 Then
			' El ID es el primer subgrupo capturado '
			result = matches(0).SubMatches(0)
		End If
	End If

	obtenerYoutubeID = result
End Function

Function ARMAWKF(RSx)



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

	XTSLNKM = ofv.GenLink("", "'nt nt2'  title='" & nomDoc& "' ", sHref, "", "", sFuncion, nomDoc, "si")


	ARMAWKF="<div class='authb-wrap'>"
	ARMAWKF=ARMAWKF & "  <div class='authb-grid'>"


	ARMAWKF=ARMAWKF & "    <div class='authb-list' id='authbList'>"

	ARMAWKF=ARMAWKF & "      <div class='authb-item'"
	ARMAWKF=ARMAWKF & "        data-url='/workflow/250'"
	ARMAWKF=ARMAWKF & "        data-user='Juan Pérez'"
	ARMAWKF=ARMAWKF & "        data-steps='20,25,30,40'"
	ARMAWKF=ARMAWKF & "        data-current='30' " & sFuncion &">"
	ARMAWKF=ARMAWKF & "        <div class='authb-item-inner'>"
	ARMAWKF=ARMAWKF & "          <div class='authb-main'>"
	ARMAWKF=ARMAWKF & "            <div>"
	ARMAWKF=ARMAWKF & XTSLNKM
	ARMAWKF=ARMAWKF & "            </div>"
	Select Case rsx("objtipo")
		Case "PD"
		ARMAWKF=ARMAWKF & "            <div class='authb-sub'>Pedido <b>#" & RSX("OBJCOD")& "</b></div>"
		Case "SN"
		ARMAWKF=ARMAWKF & "            <div class='authb-sub'>Solicitud <b>#" & RSX("OBJCOD")& "</b></div>"
	End Select

	ARMAWKF=ARMAWKF & "            <span class='noved-doc-tag'>"
	ARMAWKF=ARMAWKF & "              <i class='la la-sitemap'></i> " & rsx("cktdsc")
	ARMAWKF=ARMAWKF & "            </span>"
	ARMAWKF=ARMAWKF & "          </div>"
	ARMAWKF=ARMAWKF & "          <div class='authb-meta'>"
	ARMAWKF=ARMAWKF & "            <span class='authb-pill'><i class='la la-route'></i> Paso <b>" & rsx("paso") & "</b></span>"
	ARMAWKF=ARMAWKF & "            <span class='authb-pill'><i class='la la-clock'></i>" & rsx("comienzo") & "</span>"
	ARMAWKF=ARMAWKF & "            <span class='authb-pill'><i class='la la-hourglass-half'></i> <b>" & rsx("demora") &"</b> días</span>"
	ARMAWKF=ARMAWKF & "          </div>"
	ARMAWKF=ARMAWKF & "        </div>"

	ARMAWKF=ARMAWKF & "      </div>"
	'ARMAWKF=ARMAWKF & armawkfdetalle(rsx)



	ARMAWKF=ARMAWKF & "    </div>"




End Function


Function ArmaVerNovedad(rsx,valor)

	Select Case rsx("importancia")
		Case 0


		sgltit = "Prioridad Alta."

		Case 1


		sgltit = "Prioridad Normal."

		Case 2


		sgltit = "Prioridad baja."


	End Select

	ArmaVerNovedad="<div class=noticia-container  >"
	xxxx = getInfoUser(rsx("contacto"))
	Autor = SESSION("SGU_nombre")
	If valor="S" Then
		ArmaVerNovedad=ArmaVerNovedad &"   <div class=fechaycerrar>"
		fecinicio=fechadorcor(CDate(CLng("0" & rsx("fecinicio") + CLng(CDate("01-01-1900")))), "")

		ArmaVerNovedad=ArmaVerNovedad &"      <div class=fechapubli>Fecha de Publicación: " &fecinicio& "</div>"
		'ArmaVerNovedad=ArmaVerNovedad &"      <i class=bte>werwrw</i>"
		ArmaVerNovedad=ArmaVerNovedad &"   </div>"
		ArmaVerNovedad=ArmaVerNovedad &"   <div class=tituloycopete>"
		ArmaVerNovedad=ArmaVerNovedad &"      <div class=titulopubli>"& rsx("msgtitulo") & "</div>"
		ArmaVerNovedad=ArmaVerNovedad &"      <div class=copetepubli>" & rsx("msgcopete") & "</div>"
		ArmaVerNovedad=ArmaVerNovedad &"   </div>"
	End If
	If RSX("IMAGEN") <> "NO" Or rsx("msgimagen")=2 Then

		ArmaVerNovedad=ArmaVerNovedad &"   <div class=imagenovideo align='center'>"
		If rsx("msgimagen")=2 Then
			idyoutube=obtenerYoutubeID(rsx("msgurldsc"))
			imgtube="https://img.youtube.com/vi/" & idyoutube & "/0.jpg"
			ArmaVerNovedad= ArmaVerNovedad & "    <i class='las la-play' ></i><iframe  frameborder='0' allow='autoplay; encrypted-media; fullscreen'  allowfullscreen style='display:none' id='" &idyoutube & "' name="&idyoutube&"></iframe> <img class='videover' data-video-id='"& idyoutube & "' max-width:400px;  src='https://img.youtube.com/vi/" & idyoutube & "/0.jpg' alt='Video 1'>"
			'response.write rsx("msgurldsc")
		Else
			ArmaVerNovedad=ArmaVerNovedad &" <img style='max-width:400px;' src='" & Replace(RSX("IMAGEN") & "","///","/") & "'>"
		End If
		ArmaVerNovedad=ArmaVerNovedad &"</div>"
	End If
	If valor="S" Then
		ArmaVerNovedad=ArmaVerNovedad &"   <div class=cuerpopubli>" & rsx("mensaje") & "</div>"
		ArmaVerNovedad=ArmaVerNovedad &"   <div class=footerpubli>"
		ArmaVerNovedad=ArmaVerNovedad &"      <div class=prioridad align='center'>Autor: "& Autor &"" & xsgltit & "</div>"
		'ArmaVerNovedad=ArmaVerNovedad &"      <div class=contacto></div>"
		ArmaVerNovedad=ArmaVerNovedad &"   </div>"
	End If
	ArmaVerNovedad=ArmaVerNovedad &"</div>"

End Function

Function armawkfdetalle(rsx)
	armawkfdetalle = " <aside class='authb-preview'>"
	armawkfdetalle = armawkfdetalle  & "         <h3><i class='la la-stream'></i> Vista previa del workflow</h3>"
	armawkfdetalle = armawkfdetalle  & "          <div class='authb-preview-sub' id='previewSub'>"
	armawkfdetalle = armawkfdetalle  & "          <i class='la la-user'></i> Iniciado por <b>Juan Pérez</b></div>"
	armawkfdetalle = armawkfdetalle  & "          <div id='previewBody'><div class='authb-hsteps'><div class='authb-hstep done'>20</div>"
	armawkfdetalle = armawkfdetalle  & "          <div class='authb-hline'></div>"
	armawkfdetalle = armawkfdetalle  & "          <div class='authb-hstep done'>25</div>"
	armawkfdetalle = armawkfdetalle  & "          <div class='authb-hline'></div>"
	armawkfdetalle = armawkfdetalle  & "          <div class='authb-hstep current'>30</div>"
	armawkfdetalle = armawkfdetalle  & "          <div class='authb-hline'></div>"
	armawkfdetalle = armawkfdetalle  & "         <div class='authb-hstep next'>40</div>"
	armawkfdetalle = armawkfdetalle  & "        </div>"
	armawkfdetalle = armawkfdetalle  & "        <div class='authb-hint'>"
	armawkfdetalle = armawkfdetalle  & "          <i class='la la-mouse-pointer'></i> AAAClick para ingresar al workflow"
	armawkfdetalle = armawkfdetalle  & "        </div>"
	armawkfdetalle = armawkfdetalle  & "      </div>"
	armawkfdetalle = armawkfdetalle  & "     </aside>"
End Function

response.write shtml
%>



