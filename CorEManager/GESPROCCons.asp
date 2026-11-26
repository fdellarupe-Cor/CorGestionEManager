<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  response.write corstyler()

	set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))

  'Encabezados

  'Formulario=chr(13)&" <TR><TD   class=tt width=98% >PROCESO</TD></tr>" & chr(13)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if

  set cn = ofv.conectar(ofv.strconn2)

  Strsql =" Select * From TABPROCESOS01  where  nivel = '1'  Order By SECUEN"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  posiHeader = 0

	Do While (Not rs.EOF)
		CELDA=""
		
		sInput =  rs("CODEXTERNO") & " " &rs("DESCRIPCION")
		Celda =  ofv.gencelda("","VT ALIGN=LEFT ","","","","","",sInput,"si")
		Fila =  ofv.genrow("","","","","","","",Celda,"si") & chr(13)
		
		STRSQL="Select * From TABPROCESOS01  where  CODPROCESO01='"&RS("CODPROCESO")&"' AND NIVEL=2"
		
		set rsN2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		
		FilaZ=""
		Do While (Not rsN2.EOF)
						
			CELDAZ=""
			sInput =  "&nbsp;&nbsp;&nbsp;" &rsn2("CODEXTERNO")
			CeldaZ = CeldaZ & ofv.gencelda("","UT ALIGN=LEFT","","","","","",sInput,"si")
			sInput =  "&nbsp;&nbsp;&nbsp;" & rsn2("DESCRIPCION")
			CeldaZ = CeldaZ & ofv.gencelda("","UT ALIGN=LEFT","","","","","",sInput,"si")
			FilaZ = filaZ & ofv.genrow("","","","","","","",CeldaZ,"si") & chr(13)
		
			posiHeader = posiHeader + 1
			IF (posiHeader > 30) then
			Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
			Formulario = Formulario & "<DIV >&nbsp;</DIV> "
			Formulario = Formulario & oPrtG.RepHeader2(cstr(""), cstr("600"))
			FormularioD = ""
			posiHeader = 2
		END IF
			RSN2.MOVENEXT
		
		LOOP
		
		CALL OFV.CERRARCONSULTA(RSN2)
		XTABLA= ofv.GenTabla("","","","50% ","","0","","0","1","","",FILAZ,"si") &"<BR>"
		
		xtabla= ofv.gencelda("","UT ALIGN=LEFT","","","","","",xtabla,"si")
		FilaZ=ofv.genrow("","","","","","","",xtabla,"si") & chr(13)
		
		FormularioD = FormularioD  & fila & FilaZ

		IF (posiHeader > 30) then
			Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
			Formulario = Formulario & "<DIV >&nbsp;</DIV> "
			Formulario = Formulario & oPrtG.RepHeader2(cstr(""), cstr("600"))
			FormularioD = ""
			posiHeader = 2
		END IF
		
		posiHeader = posiHeader + 1
		posi = posi + 1
		rs.Movenext
	
	Loop
	
	IF posiHeader > 2 THEN Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
	sHTML = sHTML & Formulario
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>