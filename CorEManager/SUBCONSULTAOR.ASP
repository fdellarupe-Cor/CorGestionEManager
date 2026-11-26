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

  if sNombre = "CONSULTA POR CATEGORIA" then

     scodmaster = "000"
     stipmaster = "PROCESO"


     IF clng("0" & request.form("PERIODO")) > 0 THEN
        XTPPER = clng("0" & request.form("PERIODO"))
     ELSE
        XTPPER = session("ORPERIODO")
     END IF

 '' sMenu = request.querystring("Menu")

      XTPROC = clng("0" & request.form("PROC"))
      XTTEVEN = clng("0" & request.form("TEVEN"))
      XTCATE = clng("0" & request.form("CATE"))
      XTRSK = clng("0" & request.form("RSK"))


'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PERIODO","si")

      Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
      Strsql = Strsql & " ESTADO IN ('OK','CL')  "
      Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
      sInput = ofv.GenerarCombo("Subconsulta0","PERIODO",XTPPER,Strsql,cn,0,1,"")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")

      Fila    = ""


'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PROCESO","si")

      Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      Strsql = Strsql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "
    '  reponse.write strsql
      sInput = ofv.GenerarCombo("Subconsulta1","PROC",XTPROC,Strsql,cn,0,3,"")

      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta1","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""
'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","TIPO DE RIESGO","si")

      Strsql = "SELECT * FROM ORTIPOEVENTO ORDER BY DESCRI"

      sInput = ofv.GenerarCombo("Subconsulta2","TEVEN",XTTEVEN,Strsql,cn,0,1,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")
      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta2","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

      Fila    = ""

'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","CATEGORIA","si")

      Strsql = "SELECT * FROM ORTIPOCATEGO where tipoevento=0" & XTTEVEN & "ORDER BY DESCRI"


      sInput = ofv.GenerarCombo("Subconsulta3","CATE",XTCATE,Strsql,cn,0,2,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta3","1","dt","")
      sInput = sinput & ofv.GenerarInput("TEVEN",XTTEVEN,"hidden","Subconsulta3","1","dt","")
      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta3","1","dt","")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      if XTTEVEN = 0 then
       xstyle=" style='display:none;'"
      else
       xstyle=""
      end if
       Fila   = Fila & ofv.GenRow("aaa " & xstyle,"","","","30","","",Celda,"si")


      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta3","",sAccion,"","post","",Fila,"si")

      Fila    = ""

'-------------------------

      Celda = ofv.GenCelda("","tt","","","","","","RIESGO","si")

      Strsql = "SELECT * FROM ORIMPRIESGO WHERE AGRUPA = '" & XTTEVEN & "' AND CATEGO = '" & XTCATE & "'  ORDER BY DESCRI"

      sInput = ofv.GenerarCombo("Subconsulta4","RSK",XTRSK,Strsql,cn,0,2,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta4","1","dt","")
      sInput = sinput & ofv.GenerarInput("TEVEN",XTTEVEN,"hidden","Subconsulta4","1","dt","")
      sInput = sinput & ofv.GenerarInput("CATE",XTCATE,"hidden","Subconsulta4","1","dt","")
      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta4","1","dt","")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta4","",sAccion,"","post","",Fila,"si")

      Fila    = ""

      if XTPROC  = 0 then XTPROC = "ALL"
      if XTTEVEN = 0 then XTTEVEN = "ALL"
      if XTCATE  = 0 then XTCATE = "ALL"
      if XTRSK   = 0 then XTRSK = "ALL"


      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sAccion  = "ORMATRIXIMPACTO.asp?PERIODO=" & XTPPER & "&FILSEL=CAT&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & "&SUBFILTRO2=" & XTCATE & "&SUBFILTRO3=" & XTRSK & ""
      sTARGET  = "_BLANK"


    elseif sNombre = "CONSULTA POR SECTOR" then


     scodmaster = "000"
     stipmaster = "PROCESO"


 '' sMenu = request.querystring("Menu")

     IF clng("0" & request.form("PERIODO")) > 0 THEN
        XTPPER = clng("0" & request.form("PERIODO"))
     ELSE
        XTPPER = session("ORPERIODO")
     END IF


      XTPROC = clng("0" & request.form("PROC"))
      XTTEVEN = clng("0" & request.form("TEVEN"))
      XTCATE = clng("0" & request.form("CATE"))
      XTRSK = clng("0" & request.form("RSK"))

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PERIODO","si")

      Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
      Strsql = Strsql & " ESTADO IN ('OK','CL')  "
      Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
      sInput = ofv.GenerarCombo("Subconsulta0","PERIODO",XTPPER,Strsql,cn,0,1,"")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")

      Fila    = ""


'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PROCESO","si")

      Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      Strsql = Strsql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "

      sInput = ofv.GenerarCombo("Subconsulta1","PROC",XTPROC,Strsql,cn,0,3,"")
      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta1","1","dt","")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""
'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","AREA/SECTOR","si")

      Strsql = "SELECT * FROM ECO ORDER BY DESCRI"

      sInput = ofv.GenerarCombo("Subconsulta2","TEVEN",XTTEVEN,Strsql,cn,0,1,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")
      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta2","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

      Fila    = ""

'-----------------------------



      if XTPROC  = 0 then XTPROC = "ALL"
      if XTTEVEN = 0 then XTTEVEN = "ALL"
      if XTCATE  = 0 then XTCATE = "ALL"
      if XTRSK   = 0 then XTRSK = "ALL"


      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sAccion  = "ORMATRIXIMPACTO.asp?PERIODO=" & XTPPER & "&FILSEL=SEC&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & ""
      sTARGET  = "_BLANK"




  elseif sNombre = "CONSULTA POR CIRCUITO" then

     scodmaster = "000"
     stipmaster = "PROCESO"


 '' sMenu = request.querystring("Menu")

     IF clng("0" & request.form("PERIODO")) > 0 THEN
        XTPPER = clng("0" & request.form("PERIODO"))
     ELSE
        XTPPER = session("ORPERIODO")
     END IF


      XTPROC = clng("0" & request.form("PROC"))
      XTTEVEN = clng("0" & request.form("TEVEN"))
      XTCATE = clng("0" & request.form("CATE"))
      XTRSK = clng("0" & request.form("RSK"))

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PERIODO","si")

      Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
      Strsql = Strsql & " ESTADO IN ('OK','CL')  "
      Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
      sInput = ofv.GenerarCombo("Subconsulta0","PERIODO",XTPPER,Strsql,cn,0,1,"")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta0","",sAccion,"","post","",Fila,"si")

      Fila    = ""


'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PROCESO","si")

      Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      Strsql = Strsql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "
' response.write strsql
      sInput = ofv.GenerarCombo("Subconsulta1","PROC",XTPROC,Strsql,cn,0,3,"")
      sInput = sinput & ofv.GenerarInput("PERIODO",XTPPER,"hidden","Subconsulta1","1","dt","")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "subconsultaOR.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""
'-----------------------------




      if XTPROC  = 0 then XTPROC = "ALL"
      if XTTEVEN = 0 then XTTEVEN = "ALL"
      if XTCATE  = 0 then XTCATE = "ALL"
      if XTRSK   = 0 then XTRSK = "ALL"


      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sAccion  = "ORMATRIXIMPACTO.asp?PERIODO=" & XTPPER & "&FILSEL=CRT&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & ""
      sTARGET  = "_BLANK"



'Listas de Complementos
  elseif sNombre = "CONSULTA POR COMPONENTE" then




'Fin preparacion
  end if


'''' Armado del Form de despacho '''''''''''''''''''''''

    if formx <> "" then


      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
      sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")

   '' sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
    Formulario = Formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post","",formx,"si")

    end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''

  sHTML = sHTML & "<center><BR>"

  if sNombre = "CONSULTA POR TEXTO" then

      sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","80% ","","0","","5","0","","",Formulario,"si")

  else

      sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","60% ","","0","","5","0","","",Formulario,"si")

  end if

  sHTML = sHTML & "</center>"

  sHTML = sHTML & HTMgenlista

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  if session("SMPLMENU") = "S" then
     sAccion = "novedmenu.asp"
     Celda   = ofv.BotonVolver(sAccion)
  end if

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

  if sNombre = "CONSULTA POR TEXTO" then

     sHTML = sHTML & " Init();"

  end if

  sHTML = sHTML & "</script>"


  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

function GenListaDoc(HTgrp,HTcod,HTtip,HTnam)

htdata = ""




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' SELECCION ACTUAL '''''''''''''''''''''''''''''''''''''''''''''''


       if HTgrp = "M" then
          descrit = "TODOS LOS DOCUMENTOS"
       else
          descrit = "SIN DESCRIPCION CONOCIDA"
          if HTgrp = "T" then

              Strsql2 = " select descri from tema where codtema = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)


          elseif HTgrp = "S" then

              Strsql2 = " select descri from PRODSERV where codPRODSERV = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

          else

              Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

          end if

       end if

      sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput3  = ofv.GenerarInput("ptextot",textot,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("ptgrp",HTgrp,"hidden","","1","dt","")
      sInput5  = ofv.GenerarInput("ptcod",HTcod,"hidden","","1","dt","")
      sInput6  = ofv.GenerarInput("pttip",HTtip,"hidden","","1","dt","")
      sInput7  = ofv.GenerarInput("ptnam",HTnam,"hidden","","1","dt","")

      sInputb  = ofv.GenerarInput("bot","OK","button","","1","bt onclick='clselecc();' ","")

      HTdatos = sInput1 & sINPUT2 & sINPUT3 & sINPUT4 & sINPUT5 & sINPUT6  & sINPUT7 & sInputb


      sImagen2 = "<B><FONT FACE=corflaticon COLOR=#00aa00 SIZE=3>&#097;</FONT></B>"
      sImagen  = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>4</FONT></B>"
      sImagen3 = "<B><FONT FACE=corflaticon COLOR=#99ccee SIZE=1>&#060;</FONT></B>"

      celdasl = ""
      Celdasl = celdasl & ofv.GenCelda("","uts style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;' ","","1%","","","",sImagen2,"si")
      Celdasl = celdasl & ofv.GenCelda("","ttt","","","","","","&nbsp;" & descrit,"si")
      Celdasl = celdasl & ofv.GenCelda("","ut","","2%","","","",HTdatos,"si")
      Filasl  = ofv.GenRow("","","","","","","",Celdasl,"si")

     sAccion = "SubconsultaOR.asp?frm=4&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
     Filasl  = ofv.GenForm("Subconsulta4","",sAccion,"","post","",Filasl,"si")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   STYLOTE = " style='border-style:solid;border-width:1;border-color:#336699;background-color:#ffffff;' "
   STYLOT  = ""

   dim ATABLVL(20,2)
   FilaslT  = ""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' LISTA '''''''''''''''''''''''''''''''''''''''''''''''


'' 1ERA '''''''''''''''''''''''''''''''''''''''''''''''

       descrit = "TODOS LOS DOCUMENTOS"

       HTLEVEL = 1

       if HTgrp = "M" then

          STYLOT  = STYLOTE
       end if

      sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('M','99999','TODOS','TODOS');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

      celdasl = ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
      Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=15 " & STYLOT & " ","","","","","",sINPUTT,"si")
      FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' 2DA '''''''''''''''''''''''''''''''''''''''''''''''


   STYLOT  = ""

       if HTgrp = "T" OR HTgrp = "S" OR HTgrp = "D" then

          HTLEVEL = 2

          descrit = "TEMAS"


          descrit = "SIN DESCRIPCION CONOCIDA"
          if HTgrp = "T" then

              Hxcod = HTcod

              Strsql2 = " select descri from tema where codtema = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

              STYLOT  = STYLOTE


          elseif HTgrp = "S" then

              Strsql2 = " select codtema from PRODSERV where codPRODSERV = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 Hxcod = rsr(0)

                 Strsql2 = " select descri from tema where codtema = 0" & rsr(0)
                 set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                 if not rsr.eof then
                    descrit=rsr(0)
                 end if

              end if
              call ofv.cerrarconsulta(rsr)

          else



              Strsql2 = " select tema from CsRelPath where cobj = 0" & HTcod & " and tobj = '" & HTtip & "' "
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 Hxcod = rsr(0)


                 Strsql2 = " select descri from tema where codtema = 0" & rsr(0)
                 set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                 if not rsr.eof then
                    descrit=rsr(0)
                 end if

              end if
              call ofv.cerrarconsulta(rsr)






          end if

        ATABLVL(2,1) = Hxcod
        ATABLVL(2,2) = "TEMA"

         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('T','" & Hxcod & "','TEMA','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=14 " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'' 3erA '''''''''''''''''''''''''''''''''''''''''''''''




   STYLOT  = ""



       if HTgrp = "S" OR HTgrp = "D" then

          descrit = "SUBTEMAS"

          HTLEVEL = 3


          descrit = "SIN DESCRIPCION CONOCIDA"

         if HTgrp = "S" then

              Hxcod = HTcod
              Hxtip = HTtip

              Strsql2 = " select descri from PRODSERV where codPRODSERV = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

              STYLOT  = STYLOTE


          else



              Strsql2 = " select stema,stematipo from CsRelPath where cobj = 0" & HTcod & " and tobj = '" & HTtip & "' "
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 Hxcod = rsr(0)
                 Hxtip = rsr(1)


                 Strsql2 = " select descri from PRODSERV where codPRODSERV = 0" & rsr(0)
                 set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                 if not rsr.eof then
                    descrit=rsr(0)
                 end if

              end if
              call ofv.cerrarconsulta(rsr)






          end if

        ATABLVL(3,1) = Hxcod
        ATABLVL(3,2) = Hxtip


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('S','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=13 " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' 4rtA '''''''''''''''''''''''''''''''''''''''''''''''

   STYLOT  = ""

       descrit = "DOCUMENTOS"

       if HTgrp = "D" then


          descrit = "SIN DESCRIPCION CONOCIDA"


              Strsql2 = " select * from CsRelPath where cobj = 0" & HTcod & " and tobj = '" & HTtip & "' "
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 iii2 = 0
                 iii4 = 0
                 for iii = 1 to 13 step 2

                     httemp = rsr(3 + iii) & ""

                     if  httemp  <> "" then
                         iii2 = iii2 + 1
                         if iii2 > 1 then
                            Hxcod = rsr(3 + iii) & ""
                            Hxtip = rsr(4 + iii) & ""

                            HTLEVEL = HTLEVEL + 1

                            ATABLVL(HTLEVEL,1) = Hxcod
                            ATABLVL(HTLEVEL,2) = Hxtip

                            Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & Hxcod
                            set rsr2 = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                            if not rsr2.eof then
                               descrit=rsr2(0)
                            end if
                            call ofv.cerrarconsulta(rsr2)



         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('D','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL-1)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (16 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")


                         end if
                     end if


                 next



              end if
              call ofv.cerrarconsulta(rsr)



        Hxcod = HTcod
        Hxtip = HTtip

        HTLEVEL = HTLEVEL + 1

        ATABLVL(HTLEVEL,1) = Hxcod
        ATABLVL(HTLEVEL,2) = Hxtip

        Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & HTcod
        set rsr2 = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
        if not rsr2.eof then
           descrit=rsr2(0)
        end if
        call ofv.cerrarconsulta(rsr2)

         STYLOT  = STYLOTE


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('D','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL-1)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (16 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")







       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' linea de datos '''''''''''''''''''''''''''''''''''''''''''''''

   STYLOT  = ""
       descrit = "DATOS"

       if HTLEVEL = 1 then



          Strsql2 = " select codtema,descri from tema ORDER BY INDA, DESCRI "
          set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
          DO UNTIL rsr.eof
             Hxcod   = rsr(0)
             HxTIP   = "TEMA"
             descrit = rsr(1)


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('T','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen3,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (15 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

          RSR.MOVENEXT
          LOOP
          call ofv.cerrarconsulta(rsr)



       ELSEif HTLEVEL = 2 then

          Strsql2 = " select codPRODSERV,MARCAPS,DESCRI from PRODSERV where codtema = 0" & HTcod
          set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
          DO UNTIL rsr.eof
             Hxcod   = rsr(0)
             HxTIP   = rsr(1)
             descrit = rsr(2)


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('S','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen3,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (15 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

          RSR.MOVENEXT
          LOOP
          call ofv.cerrarconsulta(rsr)


     '''  ELSEif HTLEVEL = 3 then

       ELSE

         '' response.write "aaa" & HTLEVEL

          HTARMAWHERE = ""

          HTLEVELW = HTLEVEL

          IF HTLEVELW > 10 THEN HTLEVELW = 10

          III4 = 8

          FOR III3=3 TO HTLEVELW

              III2 = HTLEVELW - III3 + 3
              III4 = III4 - 1

              HTARMAWHERE = HTARMAWHERE & " AND MC" & III4 & " = " & ATABLVL(III2,1)
              HTARMAWHERE = HTARMAWHERE & " AND MT" & III4 & " = '" & ATABLVL(III2,2) & "' "


          NEXT



          descrit = "SIN DESCRIPCION CONOCIDA"


              Strsql2 = " select COBJ,TOBJ from CsRelPath where cobj > 0 " & HTARMAWHERE
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              DO UNTIL rsr.eof

                            Hxcod = rsr(0) & ""
                            Hxtip = rsr(1) & ""

                            Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & Hxcod
                            set rsr2 = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                            if not rsr2.eof then
                               descrit=rsr2(0)
                            end if
                            call ofv.cerrarconsulta(rsr2)


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('D','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen3,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (15 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")





              RSR.MOVENEXT
              LOOP
              call ofv.cerrarconsulta(rsr)











       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''












''''''''''''''''''''''
 htdata = htdata &  "<DIV class='DINDM' ID='HTX1'  NAME='HTX1'  style='display:none;position:absolute;top:140;left:220;HEIGHT=298px;WIDTH=558px;' >"

 htdata = htdata &  "<DIV class='DINDC' ID='HTX1C' NAME='HTX1C' style='position:relative;top:+2px;left:+2;HEIGHT=290px;WIDTH=550px;' >"
 htdata = htdata &  "<CENTER>"

 htdata = htdata &  "<table id=""MTINDX"" class='TWIND' width=100% cellspacing=0 cellpadding=2 >"
 htdata = htdata &  "<tr>"
 htdata = htdata &  "<td CLASS=mbt width=99%  >&nbsp;SELECCION DE CARPETAS</td>"
 htdata = htdata &  "<td CLASS=mbt width=1% onclick=""clselecc();"" title='Ocultar seleccion.' >"
 htdata = htdata &  "<b><font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1   >r</font></b>"
 htdata = htdata &  "</td>"
 htdata = htdata &  "</tr>"
 htdata = htdata &  "</table>"

 htdata = htdata &  ofv.GenTabla("TSEL1","aaa style='display:block;' ","","95%","","0","0","2","0","","",FilaSL,"si")

 htdata = htdata &  "<DIV class='DINDC' ID='HTX1d' NAME='HTX1d' style='position:relative;HEIGHT=220px;WIDTH=546px;overflow:auto;' >"
 htdata = htdata &  "<CENTER>"
 htdata = htdata &  ofv.GenTabla("TSEL2","aaa style='display:block;' ","","95%","","0","0","2","0","","",FilaSLT,"si")
 htdata = htdata &  "</CENTER>"
 htdata = htdata &  "</div>"

 htdata = htdata &  "</CENTER>"
 htdata = htdata &  "</div>"
 htdata = htdata &  "</div>"

 htdata = htdata & "<script language=javascript >"
 htdata = htdata & " function clselecc() { " & chr(13)
 htdata = htdata & "  if  (document.all.HTX1.style.display=='none')   " & chr(13)
 htdata = htdata & "      { document.all.HTX1.style.display='block';} " & chr(13)
 htdata = htdata & "  else                                             " & chr(13)
 htdata = htdata & "      { document.all.HTX1.style.display='none';}  " & chr(13)
 htdata = htdata & " } " & chr(13)
 htdata = htdata & " function clactfrm(fngrp,fncod,fntip,fnnam) { " & chr(13)
 htdata = htdata & "    document.all.ptgrp.value=fngrp; " & chr(13)
 htdata = htdata & "    document.all.ptcod.value=fncod; " & chr(13)
 htdata = htdata & "    document.all.pttip.value=fntip; " & chr(13)
 htdata = htdata & "    document.all.ptnam.value=fnnam; " & chr(13)
 htdata = htdata & "    document.all.ptextot.value=document.all.textot.value; " & chr(13)
' htdata = htdata & "    document.Subconsulta4.submit(); " & chr(13)
' htdata = htdata & "    return false; " & chr(13)
 htdata = htdata & " } " & chr(13)



 if  frmsen = "S" then
     htdata = htdata & " clselecc(); " & chr(13)

 end if


 htdata = htdata & "</script>"

' Whtml = Whtml & "<script language=javascript >"
' Whtml = Whtml & "document.all.mt1.style.display='none';"
' Whtml = Whtml & "document.all.mt2.style.display='block';"
' Whtml = Whtml & " DCEXRCLC();"
' Whtml = Whtml & "</script>"

''''''''''''''''''''''

   GenListaDoc = htdata

end function

%>