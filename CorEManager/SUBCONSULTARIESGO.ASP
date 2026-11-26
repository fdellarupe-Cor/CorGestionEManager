<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

  session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

  sMenu = request.querystring("Menu")
  sNombre = request.querystring("Menu") 'request.querystring("Nombre")
  sTema = request.form("codtema")
  if sTema = "" or sTema = "nsnc" then sTema = 0
  sSubtema = request.form("codprodserv")
  if sSubtema = "" or sSubtema = "nsnc" then sSubtema = 0
  sEsquema = request.form("codobjeto")
  if sEsquema = "" or sEsquema = "nsnc" then sEsquema = 0
  xComp = ""
  xComp = request.form("XComplem")
  

  eACTIVIDAD = request.querystring("eACTIVIDAD")	
  eRIESGO = request.querystring("eRIESGO")
  ecodmaster = request.querystring("ecodmaster")
  etipmaster = request.querystring("etipmaster")


	'response.Write(chr(13) & chr(13) & chr(13) & eACTIVIDAD & chr(13) & eRIESGO & chr(13) & ecodmaster & chr(13) & etipmaster & chr(13) & chr(13))

  XXXX = NEWEMPRE("NO")

  sEMPRE = request.querystring("EMPRE")
  IF SEMPRE = "" THEN
     SEMPRE = ATREMP(0,0)
  ELSE
     XXXX = NEWEMPRE(SEMPRE)
  END IF

  sUN = SEMPRE

  if sUN = "" then sUN = 1
  session("UN") = sUN


  'response.write  smenu
  set cn = ofv.conectar(ofv.strconn2)

  sNombren = sNombre

  sNombre = "Consulta por " & sNombre
  snombre = ucase(snombre)

  sHTML = ofv.FormHeader(session("FormN")) & chr(13)

  sImagenM = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"


  Fila    = ""
  Celda   = ""
  sHref   = ""

  IF  Session("EMPMUN") = "S" THEN

  Fila    = ""

  SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

  Fila    = ""
  Celda   = ""
  sHref   = ""

  sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
  Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
  Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
  Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

  Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

  Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





  shtml=shtml & "<script language=javascript>" & chr(13)
  shtml=shtml & "function shwDem(){" & chr(13)
  shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
  shtml=shtml & "}" & chr(13)
  shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
  sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
  sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
  sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
  sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
  TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:150px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


  TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
  TABLAX = TablaX & "<tr><td  "
  TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
  TABLAX = TablaX & "<tr><td colspan=3 "
  TABLAX = TablaX & " >&nbsp;</td></tr>"

  FOR IAEMP=1 TO TAEMP

   '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

     IF ATREMP(IAEMP,2) = 1 THEN

      IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
         TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
      ELSE
         TABLAX = TablaX & "<tr>"
      END IF

      TABLAX = TablaX & "<td width=5 "
      IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
         TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
      ELSE
         TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
      END IF

      sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
      TABLAX = TablaX & "<td width=5 "
      TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

      TABLAX = TablaX & "<td class=vt  >"

      sHref = "SUBCONSULTARIESGO.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)

    '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
      sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
      TABLAX = TablaX & sLink
      TABLAX = TablaX & "</td>"
      TABLAX = TablaX & "</tr>"
     END IF

  NEXT

  TABLAX = TablaX & OFV.genFtabla()

  TABLA = Tabla & TABLAX

  TABLA = Tabla & "</DIV>"

  sHTML = sHTML & TABLA

  end if

  Fila = ""

     if session("ORPERIODO") = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  "
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")

        end if


      end if


  tt="&nbsp;"



  Fila       = ""
  formx      = ""
  Formulario = ""

 'Descripcion
  sInput = "<big>SELECCIONE LOS PARAMETROS PARA LA CONSULTA</big>"
  Celda  = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
  Fila   = ofv.genrow("","","","","30","","",Celda,"si")

  Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
  Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
  Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")



     scodmaster = "000"
     stipmaster = "PROCESO"


     IF clng("0" & request.form("PERIODO")) > 0 THEN
        XTPPER = clng("0" & request.form("PERIODO"))
     ELSE
        XTPPER = session("ORPERIODO")
     END IF

	 	SET ID_AR = request.querystring("ID_Actividad_Riesgo")
		
		if(clng("0" & request.form("RIESGOID")) <> 0) then
			XTTEVEN = request.form("RIESGOID")
		else
			XTTEVEN = request.querystring("ID")
		end if
		
		 ID_RIESGO = request.querystring("ID")		
		
		if (ID_RIESGO <> XTTEVEN) then
			sUpdate = "UPDATE ORACTIVRIESGO SET ITEMRIESGO = " & XTTEVEN & " WHERE OBJID = " & ID_AR
			call ofv.crearconsultaEx(sUpdate,cn,1,parametros)
			ID_RIESGO = XTTEVEN
			'response.Write(sUpdate & chr(13))
		end if
		
	
			
		
		Strsql = "SELECT " 
		Strsql = Strsql & "O1.OBJID AS 'OBJID_TEVENTO',	O2.OBJID AS 'OBJID_TCATEGORIA',	O3.OBJID AS 'OBJID_TRIESGO' "
		Strsql = Strsql & "FROM "
		Strsql = Strsql & "ORTIPOEVENTO O1 "
		Strsql = Strsql & "LEFT JOIN "
		Strsql = Strsql & "ORTIPOCATEGO O2 "
		Strsql = Strsql & "ON O1.OBJID = O2.TIPOEVENTO "
		Strsql = Strsql & "LEFT JOIN "
		Strsql = Strsql & "ORIMPRIESGO O3 "
		Strsql = Strsql & "ON O3.CATEGO = O2.OBJID "
		Strsql = Strsql & "AND O3.AGRUPA = O1.OBJID "
		Strsql = Strsql & "WHERE "
		Strsql = Strsql & "O3.OBJID = " & ID_RIESGO
		
		'response.Write(strsql & chr(13))
		
		
		set RS_Descri = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		
		if not RS_Descri.eof then
			XTPEVE = RS_Descri(0).value
			XTPROC = RS_Descri(1).value
			XTTEVEN = RS_Descri(2).value
		else
			XTPEVE = 0
			XTPROC = 0
			XTTEVEN = 0
		end if
		
		if(clng("0" & request.form("EVENTOID")) <> 0) then
			if (cstr(XTPEVE) <> cstr(request.form("EVENTOID"))) then
				XTTEVEN = 0
				'response.Write("CATEGORIAID_aaaa")
				'response.Write(request.form("EVENTOID") & "----" & XTPEVE)
			end if
			XTPEVE = request.form("EVENTOID")
		end if
		
		'response.Write(XTPEVE)
		
		
		if(clng("0" & request.form("CATEGORIAID")) <> 0) then
		if (cstr(XTPROC) <> cstr(request.form("CATEGORIAID"))) then
				XTTEVEN = 0
				'response.Write("CATEGORIAID_zzzz")
			end if
			XTPROC = request.form("CATEGORIAID")
			
		end if
		call ofv.cerrarconsulta(RS_Descri)

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","TIPO EVENTO","si")

      Strsql = "select OBJID AS 'EVENTOID', DESCRI from ORTIPOEVENTO"
      sInput = ofv.GenerarCombo("Subconsulta0","EVENTOID",XTPEVE,Strsql,cn,0,1,"")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","CATEGORIAS","si")


      Strsql = "select OBJID AS 'CATEGORIAID',TIPOEVENTO, DESCRI from ORTIPOCATEGO where TIPOEVENTO = " & XTPEVE
      sInput = ofv.GenerarCombo("Subconsulta0","CATEGORIAID",XTPROC,Strsql,cn,0,2,"")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","RIESGO","si")
	  
	  Strsql = "select OBJID AS 'RIESGOID', DESCRI from ORIMPRIESGO where catego= " & XTPROC
	  'response.Write(XTTEVEN)
      sInput = ofv.GenerarCombo("Subconsulta0","RIESGOID",XTTEVEN,Strsql,cn,0,1,"")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "SUBCONSULTARIESGO.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&ID_Actividad_Riesgo=" & ID_AR & "&ID=" & ID_RIESGO & "&eACTIVIDAD=" & eACTIVIDAD & "&eRIESGO=" & eRIESGO & "&ecodmaster=" & ecodmaster & "&etipmaster=" & etipmaster
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")
	  
	  Fila = ""
'-----------------------------


      if XTPROC  = 0 then XTPROC = "ALL"
      if XTTEVEN = 0 then XTTEVEN = "ALL"
      if XTCATE  = 0 then XTCATE = "ALL"
      if XTRSK   = 0 then XTRSK = "ALL"


      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila

  sHTML = sHTML & "<center><BR>"

  sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","60% ","","0","","5","0","","",Formulario,"si")

  sHTML = sHTML & "</center>"

  sHTML = sHTML & HTMgenlista




     sAccion = "ORCIRCUITODET.asp?ACTIVIDAD=" & eACTIVIDAD & "&RIESGO=" & eRIESGO & "&codmaster=" & ecodmaster & "&tipmaster=" & etipmaster
     Celda   = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "<script language=JavaScript>"
  sHTML = sHTML & " function Init() "
  sHTML = sHTML & "   { ff(); setTimeout (" & chr(34) & "Init()" & chr(34) & ", 350); }"
  sHTML = sHTML & " function ff() "
  sHTML = sHTML & "  {   "
  sHTML = sHTML & "  if (document.all.fnt01.color=='#f7aa22') { document.all.fnt01.color='#aa0000'; } "
  sHTML = sHTML & "  else { if (document.all.fnt01.color=='#aa0000') { document.all.fnt01.color='#f75500'; } "
  sHTML = sHTML & "  else {if (document.all.fnt01.color=='#f75500') { document.all.fnt01.color='#f7aa22'; } } } "
  sHTML = sHTML & "   } "


  sHTML = sHTML & "</script>"


  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>