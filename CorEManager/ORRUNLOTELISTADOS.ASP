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
  sNombre = request.querystring("Nombre")
  sTema = request.form("codtema")
  if sTema = "" or sTema = "nsnc" then sTema = 0
  sSubtema = request.form("codprodserv")
  if sSubtema = "" or sSubtema = "nsnc" then sSubtema = 0
  sEsquema = request.form("codobjeto")
  if sEsquema = "" or sEsquema = "nsnc" then sEsquema = 0
  xComp = ""
  xComp = request.form("XComplem")



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

      sHref = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)

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

  tt="&nbsp;"



  Fila       = ""
  formx      = ""
  Formulario = ""

 'Descripcion
  sInput = "<big>SELECCIONE PARAMETROS PARA LISTADO</big>"
  Celda  = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
  Fila   = ofv.genrow("","","","","30","","",Celda,"si")

  Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
  Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
  Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")


     scodmaster = "000"
     stipmaster = "PROCESO"


        XTPPER = clng("0" & request.form("TRIMESTRE"))

 '' sMenu = request.querystring("Menu")

      SORRIESGO = clng("0" & request.form("ORRIESGO"))
	  SORRANKING = clng("0" & request.form("ORRANKING"))

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PERIODO","si")

      Strsql = "SELECT TOP 8 OBJID, (CONVERT(NVARCHAR(4), ANO) + ' - TRIMESTRE ' + CONVERT(NVARCHAR(1), TRIMESTRE)) AS DETALLE "
      Strsql = Strsql & " FROM ORREGINFOBCRA WHERE estado <> 0 ORDER BY ano ASC, trimestre DESC "
      sInput = ofv.GenerarCombo("Subconsulta0","TRIMESTRE",XTPPER,Strsql,cn,0,1,"")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORRUNLOTELISTADOS.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")

      Fila    = ""
	  

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","LISTADO","si")

     SCOMSEP = chr(9)
	 scomK = ""
	 scomD = ""

         scomK = scomK & SCOMSEP & "1"
         scomD = scomD & SCOMSEP & "LINEA DE NEGOCIO"
		 scomK = scomK & SCOMSEP & "2"
         scomD = scomD & SCOMSEP & "CLASIFICACION TIPO EVENTO"
		 scomK = scomK & SCOMSEP & "3"
         scomD = scomD & SCOMSEP & "PROCESO"
		 scomK = scomK & SCOMSEP & "4"
         scomD = scomD & SCOMSEP & "PRODUCTO"
		 scomK = scomK & SCOMSEP & "5"
         scomD = scomD & SCOMSEP & "CARACTER DE EVENTO"
		 scomK = scomK & SCOMSEP & "6"
         scomD = scomD & SCOMSEP & "RANKING"

     sInput = ofv.gencomboFX("LISTADO","ORRIESGO",SORRIESGO,scomK,scomD,"")
	 sInput = sinput & ofv.GenerarInput("TRIMESTRE",XTPPER,"hidden","LISTADO","1","dt","")
	
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORRUNLOTELISTADOS.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("LISTADO","",sAccion,"","post","",Fila,"si")




      Fila    = ""
	  
	  Celda  = ofv.GenCelda("","tt","","","","","","RANKING","si")

     SCOMSEP = chr(9)
	 scomK = ""
	 scomD = ""

         scomK = scomK & SCOMSEP & "5"
         scomD = scomD & SCOMSEP & "5 PRIMEROS"
		 scomK = scomK & SCOMSEP & "10"
         scomD = scomD & SCOMSEP & "10 PRIMEROS"
		 scomK = scomK & SCOMSEP & "15"
         scomD = scomD & SCOMSEP & "15 PRIMEROS"
		 scomK = scomK & SCOMSEP & "20"
         scomD = scomD & SCOMSEP & "20 PRIMEROS"

     sInput = ofv.gencomboFX("RANKING","ORRANKING",SORRANKING,scomK,scomD,"")
	 sInput = sinput & ofv.GenerarInput("TRIMESTRE",XTPPER,"hidden","LISTADO","1","dt","")
	 sInput = sinput & ofv.GenerarInput("ORRIESGO",SORRIESGO,"hidden","RANKING","1","dt","")
	
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORRUNLOTELISTADOS.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
	  IF (clng("0" & request.form("ORRIESGO")) = 6) THEN
	      Formulario = Formulario & ofv.GenForm("RANKING","",sAccion,"","post","",Fila,"si")
	  END IF

	  Fila    = ""
	  
      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sAccion  = "ORRUNLOTELISTADOSLIST.asp?TRIMESTRE=" & XTPPER & "&LISTADO=" & SORRIESGO  & "&CANTREGISTROS=" & SORRANKING
      sTARGET  = "_BLANK"

      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
      sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")

    ''sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
    Formulario = Formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post","",formx,"si")


  sHTML = sHTML & "<center><BR>"

       sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","60% ","","0","","5","0","","",Formulario,"si")

  sHTML = sHTML & "</center>"

  sHTML = sHTML & HTMgenlista

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML
%>