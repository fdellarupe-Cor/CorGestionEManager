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

      sHref = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)

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

  select case UCase(sNombreN)
    case "ESTADO"
      Strsql = "Select * From Estados order by estado"

    case "SECTOR"

    case "FECHA"
      Strsql = "SELECT DISTINCT Format(fecemis,'Short Date'), fecemis AS Fecha "
      Strsql = Strsql & "FROM CSEMISION WHERE Str(fecemis) Is Not Null "
      Strsql = Strsql & " order by fecemis "
    case "INSTRUMENTO"
      Strsql = "SELECT * from instrumen order by descri "

    case "PRODUCTO"

      Strsql = "SELECT distinct objtsegagr.tsegmento,TIPOSEGMENTO.descri  FROM objtsegagr, TIPOSEGMENTO,OBJETOS"

      if SESSION("EMPMUN") = "S" THEN
         if SESSION("GLBMUN") = 2 THEN
            Strsql = Strsql & " ,OBJMEMPALC "
         END IF
      END IF

      Strsql = Strsql & " where objtsegagr.tsegmento = TIPOSEGMENTO.objid "
      Strsql = Strsql & " AND OBJETOS.CODOBJETO = OBJTSEGAGR.CODOBJETO AND "

      if SESSION("EMPMUN") <> "S"  THEN
         Strsql = Strsql & " OBJETOS.OBJUN = 1"
      ELSE

         if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
            Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
         ELSE
            Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
            Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
            Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
            Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
            Strsql = Strsql & " ) "
      END IF

     end if

      Strsql = Strsql & " order by TIPOSEGMENTO.descri "

    '  RESPONSE.WRITE "<FONT COLOR=#000000>" & STRSQL & "<BR>"
    '  RESPONSE.WRITE SESSION("EMPMUN") & "<BR>"
    '  RESPONSE.WRITE SESSION("GLBMUN") & "<BR>"

    case "NUMEROCIRCULAR"

      Strsql = "SELECT distinct objtsegagr.tagrupa,TIPOagrupa.descri  FROM objtsegagr, TIPOagrupa, OBJETOS"

      if SESSION("EMPMUN") = "S" THEN
         if SESSION("GLBMUN") = 2 THEN
            Strsql = Strsql & " ,OBJMEMPALC "
         END IF
      END IF

      Strsql = Strsql & " where objtsegagr.tagrupa = TIPOagrupa.objid "
      Strsql = Strsql & " AND OBJETOS.CODOBJETO = OBJTSEGAGR.CODOBJETO AND "

      if SESSION("EMPMUN") <> "S"  THEN
         Strsql = Strsql & " OBJETOS.OBJUN = 1"
      ELSE

      if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
         Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
      ELSE
         Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
         Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
         Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
         Strsql = Strsql & " ) "
      END IF

      END IF

      Strsql = Strsql & " order by TIPOagrupa.descri "

  end select

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

    if sNombre = "CONSULTA POR PENDIENTES" then

     scodmaster = "000"
     stipmaster = "PROCESO"


 '' sMenu = request.querystring("Menu")
      XTPROC = request.form("PROC")
      XTTEVEN = request.form("TEVEN")

      IF XTPROC = "" THEN XTPROC = "0"
      IF XTTEVEN = "" THEN XTTEVEN = "0"

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PROCESO","si")

      SCOMSEP = ""
      scomK   = ""
      scomD   = ""
      sseld   = "0"
      ssels   = ","

         scomK = scomK & SCOMSEP & "0"
         scomD = scomD & SCOMSEP & "TODOS"
         SCOMSEP = chr(9)

  Strsql = "     select B.CODOBJETO,B.DESCRI, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " where "
  Strsql = Strsql & " B.APLIC = 'RG' AND "
  Strsql = Strsql & " A.MC7 = 0 "
  Strsql = Strsql & " ORDER BY B.DESCRI "
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      do until rsX1.eof

         scomK = scomK & SCOMSEP & RSX1(0)
         scomD = scomD & SCOMSEP & RSX1(1)
         SCOMSEP = chr(9)


      RSX1.MOVENEXT
      LOOP
      call ofv.cerrarconsulta(rsX1)

      sInput = ofv.gencomboFX("Subconsulta1","PROC",XTPROC,scomK,scomD,"")


      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""
'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","SUBPROCESO","si")


      SCOMSEP = chr(9)
      scomK   = ""
      scomD   = ""


         scomK = scomK & SCOMSEP & "0"
         scomD = scomD & SCOMSEP & "TODOS"
         SCOMSEP = chr(9)

  Strsql = ""
  Strsql = Strsql & " select DISTINCT B.codobjeto, B.DESCRI, B.PRIORIDAD "
  Strsql = Strsql & " from csrelpathdescri a   "
  Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "
  Strsql = Strsql & " inner join objPRTY b2 on (b2.codobjeto = b.codobjeto)   "
  Strsql = Strsql & " inner join orrskcircuito w on (a.cobj = w.codobjeto)   "
  Strsql = Strsql & " inner join orperiodos x on (w.periodo = x.objid)  "
  Strsql = Strsql & " where  B2.ESTADO <> 4 AND "
  Strsql = Strsql & "  ("
  Strsql = Strsql & " A.MC1 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC2 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC3 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC4 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC5 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC6 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC7 = " & XTPROC & " OR "
  Strsql = Strsql & " A.COBJ = " & XTPROC & "  "
  Strsql = Strsql & " ) "
  Strsql = Strsql & " ORDER BY B.DESCRI "

''RESPONSE.WRITE STRSQL & "<BR>"
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      do until rsX1.eof

         scomK = scomK & SCOMSEP & RSX1(0)
         scomD = scomD & SCOMSEP & RSX1(1)
         SCOMSEP = chr(9)


      RSX1.MOVENEXT
      LOOP
      call ofv.cerrarconsulta(rsX1)


 '    scomK = "0" & SCOMSEP & "1" & SCOMSEP & "2" & SCOMSEP & "3" & SCOMSEP & "4" & SCOMSEP & "5" & SCOMSEP & "10" & SCOMSEP & "99"
 '     scomD = "TODOS" & SCOMSEP & "1 DIA" & SCOMSEP & "2 DIAS" & SCOMSEP & "3 DIAS" & SCOMSEP & "4 DIAS" & SCOMSEP & "5 DIAS" & SCOMSEP & "10 DIAS" & SCOMSEP & "MAS DE 10 DIAS"


      sInput = ofv.gencomboFX("Subconsulta2","TEVEN",XTTEVEN,scomK,scomD,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

      Fila    = ""



      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sAccion  = "ORLISTAWKFPENDIENTES.asp?FILSEL=CAT&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & "&SUBFILTRO2=0&SUBFILTRO3=0&Nombre=" & sNombren & "&Menu=" & sMenu
      sTARGET  = ""


    ELSEif sNombre = "CONSULTA POR NOVEDADES" then

     scodmaster = "000"
     stipmaster = "PROCESO"


 '' sMenu = request.querystring("Menu")
      XTPROC = request.form("PROC")
      XTTEVEN = request.form("TEVEN")

      IF XTPROC = "" THEN XTPROC = "0"
      IF XTTEVEN = "" THEN XTTEVEN = "0"

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","PROCESO","si")

      SCOMSEP = ""
      scomK   = ""
      scomD   = ""
      sseld   = "0"
      ssels   = ","

         scomK = scomK & SCOMSEP & "0"
         scomD = scomD & SCOMSEP & "TODOS"
         SCOMSEP = chr(9)

  Strsql = "     select B.CODOBJETO,B.DESCRI, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " where "
  Strsql = Strsql & " B.APLIC = 'RG' AND "
  Strsql = Strsql & " A.MC7 = 0 "
  Strsql = Strsql & " ORDER BY B.DESCRI "
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      do until rsX1.eof

         scomK = scomK & SCOMSEP & RSX1(0)
         scomD = scomD & SCOMSEP & RSX1(1)
         SCOMSEP = chr(9)


      RSX1.MOVENEXT
      LOOP
      call ofv.cerrarconsulta(rsX1)

      sInput = ofv.gencomboFX("Subconsulta1","PROC",XTPROC,scomK,scomD,"")


      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""
'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","SUBPROCESO","si")


      SCOMSEP = chr(9)
      scomK   = ""
      scomD   = ""


         scomK = scomK & SCOMSEP & "0"
         scomD = scomD & SCOMSEP & "TODOS"
         SCOMSEP = chr(9)

  Strsql = ""
  Strsql = Strsql & " select DISTINCT B.codobjeto, B.DESCRI, B.PRIORIDAD "
  Strsql = Strsql & " from csrelpathdescri a   "
  Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "
  Strsql = Strsql & " inner join objPRTY b2 on (b2.codobjeto = b.codobjeto)   "
  Strsql = Strsql & " inner join orrskcircuito w on (a.cobj = w.codobjeto)   "
  Strsql = Strsql & " inner join orperiodos x on (w.periodo = x.objid)  "
  Strsql = Strsql & " where  B2.ESTADO <> 4 AND "
  Strsql = Strsql & "  ("
  Strsql = Strsql & " A.MC1 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC2 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC3 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC4 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC5 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC6 = " & XTPROC & " OR "
  Strsql = Strsql & " A.MC7 = " & XTPROC & " OR "
  Strsql = Strsql & " A.COBJ = " & XTPROC & "  "
  Strsql = Strsql & " ) "
  Strsql = Strsql & " ORDER BY B.DESCRI "

''RESPONSE.WRITE STRSQL & "<BR>"
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      do until rsX1.eof

         scomK = scomK & SCOMSEP & RSX1(0)
         scomD = scomD & SCOMSEP & RSX1(1)
         SCOMSEP = chr(9)


      RSX1.MOVENEXT
      LOOP
      call ofv.cerrarconsulta(rsX1)


 '    scomK = "0" & SCOMSEP & "1" & SCOMSEP & "2" & SCOMSEP & "3" & SCOMSEP & "4" & SCOMSEP & "5" & SCOMSEP & "10" & SCOMSEP & "99"
 '     scomD = "TODOS" & SCOMSEP & "1 DIA" & SCOMSEP & "2 DIAS" & SCOMSEP & "3 DIAS" & SCOMSEP & "4 DIAS" & SCOMSEP & "5 DIAS" & SCOMSEP & "10 DIAS" & SCOMSEP & "MAS DE 10 DIAS"


      sInput = ofv.gencomboFX("Subconsulta2","TEVEN",XTTEVEN,scomK,scomD,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

      Fila    = ""



      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sAccion  = "ORLISTAWKFPENDIENTES.asp?FILSEL=CAT&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & "&SUBFILTRO2=0&SUBFILTRO3=0&Nombre=" & sNombren & "&Menu=" & sMenu
      sTARGET  = ""



    elseif sNombre = "CONSULTA POR PRODUCTO" then


 '' sMenu = request.querystring("Menu")
      xtseg = clng("0" & request.form("segm"))
      xtagr = clng("0" & request.form("agrm"))
      xtPRO = clng("0" & request.form("PROm"))

'--------------------

      Celda  = ofv.GenCelda("","tt","","","","","","Segmento","si")
      sInput = ofv.GenerarCombo("Subconsulta1","segm",xtseg,Strsql,cn,0,1,"")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""
'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","Grupo","si")

      Strsql = "SELECT distinct objtsegagr.tagrupa,tipoagrupa.descri  FROM objtsegagr, tipoagrupa, OBJETOS"
      if SESSION("EMPMUN") = "S" THEN
         if SESSION("GLBMUN") = 2 THEN
            Strsql = Strsql & " ,OBJMEMPALC "
         END IF
      END IF
      Strsql = Strsql & " where objtsegagr.tagrupa = tipoagrupa.objid and "
      Strsql = Strsql & " objtsegagr.tsegmento = " & xtseg  & " AND "


      if SESSION("EMPMUN") <> "S"  THEN
         Strsql = Strsql & " OBJETOS.OBJUN = 1"
      ELSE

      if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
         Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
      ELSE
         Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
         Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
         Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
         Strsql = Strsql & " ) "
      END IF

      end if

      Strsql = Strsql & " order by tipoagrupa.descri "

      sInput = ofv.GenerarCombo("Subconsulta2","agrm",xtagr,Strsql,cn,0,1,"")
      sInput = sinput & ofv.GenerarInput("segm",xtseg,"hidden","Subconsulta2","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

      Fila    = ""

'-----------------------------

      Celda = ofv.GenCelda("","tt","","","","","","PRODUCTO","si")

      Strsql = "SELECT distinct objetos.tsegmento,tipoproducto.descri "
      Strsql = Strsql & " FROM objtsegagr, objetos, tipoproducto "

      if SESSION("EMPMUN") = "S" THEN
         if SESSION("GLBMUN") = 2 THEN
            Strsql = Strsql & " ,OBJMEMPALC "
         END IF
      END IF

      Strsql = Strsql & " where objtsegagr.tagrupa = " & xtagr & " and "
      Strsql = Strsql & " objtsegagr.tsegmento = " & xtseg & " and "
      Strsql = Strsql & " objtsegagr.codobjeto = objetos.codobjeto and "
      Strsql = Strsql & " objetos.tsegmento = tipoproducto.objid AND "


      if SESSION("EMPMUN") <> "S"  THEN
         Strsql = Strsql & " OBJETOS.OBJUN = 1"
      ELSE

      if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
         Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
      ELSE
         Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
         Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
         Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
         Strsql = Strsql & " ) "
      END IF

      end if


      Strsql = Strsql & "  order by tipoproducto.descri "
      sInput = ofv.GenerarCombo("Subconsulta3","PROm",xtPRO,Strsql,cn,0,1,"")
      sInput = sinput & ofv.GenerarInput("SEGM",xtseg,"hidden","Subconsulta3","1","dt","")
      sInput = sinput & ofv.GenerarInput("AGRM",xtagr,"hidden","Subconsulta3","1","dt","")
      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta3","",sAccion,"","post","",Fila,"si")

      Fila    = ""

'-------------------------

      Celda = ofv.GenCelda("","tt","","","","","","TIPO","si")

      Strsql = "SELECT distinct TIPOBJETO.DESCRI  FROM objtsegagr, objetos, tipoproducto,TIPOBJETO "

      if SESSION("EMPMUN") = "S" THEN
         if SESSION("GLBMUN") = 2 THEN
            Strsql = Strsql & " ,OBJMEMPALC "
         END IF
      END IF

      Strsql = Strsql & " where TIPOBJETO.TIPOBJETO = OBJETOS.TIPOBJETO AND "
      Strsql = Strsql & " objtsegagr.tagrupa = " & xtagr & " and objtsegagr.tsegmento = " & xtseg
      Strsql = Strsql & " and objtsegagr.codobjeto = objetos.codobjeto and "
      Strsql = Strsql & " objetos.tsegmento = tipoproducto.objid  AND tipoproducto.objid = " & XTPRO & " AND "


      if SESSION("EMPMUN") <> "S"  THEN
         Strsql = Strsql & " OBJETOS.OBJUN = 1"
      ELSE

      if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
         Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
      ELSE
         Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
         Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
         Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
         Strsql = Strsql & " ) "
      END IF

     end if

      sInput = ofv.GenerarCombox3(strSQL, cn, 0,0, "SFILTRO")

      if xtseg = 0 then xtseg = "nsnc"
      if xtagr = 0 then xtagr = "nsnc"
      if xtPRO = 0 then xtPRO = "nsnc"

      sInput = sinput & ofv.GenerarInput("ListaS",xtseg,"hidden","Subconsulta","1","dt","")
      sInput = sinput & ofv.GenerarInput("ListaG",xtagr,"hidden","Subconsulta","1","dt","")
      sInput = sinput & ofv.GenerarInput("ListaP",xtPRO,"hidden","Subconsulta","1","dt","")

      Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx  = Fila


      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren



    elseif sNombre = "CONSULTA POR NUMEROCIRCULAR" then

      xtagr   = clng("0" & request.form("agrm"))

      Celda   = ofv.GenCelda("","tt ","","","","","","Grupo","si")
      sInput  = ofv.GenerarCombo("Subconsulta1","agrm",xtagr,Strsql,cn,0,1,"")
      Celda   = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila    = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

      Fila    = ""

       Celda = ofv.GenCelda("","tt","","","","","","Numero","si")

                      Strsql = "SELECT distinct objetos.codiden, "
                      strsql = strsql & " INDI = CASE "
                      strsql = strsql & " WHEN ISNUMERIC(objetos.codiden)= 1 "
                      strsql = strsql & " THEN REPLACE(CONVERT(NVARCHAR(5), SPACE(5 - (DATALENGTH(OBJETOS.CODIDEN)/2)) + objetos.codiden ),' ','0') "
                      strsql = strsql & " ELSE CONVERT(NVARCHAR(8), objetos.codiden ) "
                      strsql = strsql & " END "
                      strsql = strsql & " FROM objtsegagr, objetos "

                      if SESSION("EMPMUN") = "S" THEN
                         if SESSION("GLBMUN") = 2 THEN
                            Strsql = Strsql & " ,OBJMEMPALC "
                         END IF
                      END IF

                      strsql = strsql & " where objtsegagr.tagrupa = " & xtagr & " and "
                      strsql = strsql & " objtsegagr.codobjeto = objetos.codobjeto AND "


      if SESSION("EMPMUN") <> "S"  THEN
         Strsql = Strsql & " OBJETOS.OBJUN = 1"
      ELSE

      if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
         Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
      ELSE
         Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
         Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
         Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
         Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
         Strsql = Strsql & " ) "
      END IF

      end if

      strsql = strsql & " order by INDI, objetos.codiden "

      sInput   = ofv.GenerarCombox3(strSQL, cn, 0,1, "ListaN")
      if xtagr = 0 then xtagr = "nsnc"
      sInput   = sinput & ofv.GenerarInput("ListaG",xtagr,"hidden","Subconsulta","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


   elseif sNombre = "CONSULTA POR INSTRUMENTO" then

    Celda      = ofv.GenCelda("","tt","","","","","","Instrumento","si")

    sInput     = ofv.GenerarCombox2(strSQL, cn, 0,1, "Lista")
    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    Celda      = ofv.GenCelda("","tt","","","","","","Identificacion","si")
    sInput     = ofv.GenerarInput("buscar","","agregar","","60","dt","")
    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


  elseif sNombre = "CONSULTA POR IDENTIFICACION" then


    Celda      = ofv.GenCelda("","tt","","","","","","Identificacion","si")
    sInput     = ofv.GenerarInput("buscar","","agregar","","60","dt","")
    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren

  elseif sNombre = "CONSULTA POR TEXTO" then

      sInput  = ""
      sInput2  = ""
      sInput3  = ""
      sInput4  = ""
      sInput5  = ""
      sInput6  = ""
      sInput7  = ""

      xstdoc   = "0"
      xstsel   = "0"
      xstsec   = "0"







 




''response.write "<br><br><br>---------------------------"'
''for each x in request.form
'   response.write "-- " & x & " : " & request.form(x) & " --<br>"
'next
''response.write "---------------------------"


      xx1 = request.form("Ntipodoc") & ""
      xx2 = request.form("Ntcod") & ""
      xx3 = request.form("Nlista") & ""
      xx4 = request.form("textot") & ""
      xx5 = request.form("tsel") & ""

      ttcc = request.form("ttcc") & ""




      stipodoc    = request.form("Ntipodoc") & ""
      if stipodoc = "" then
      '   stipodoc = "99999"
         xstdoc   = "0"
      else
         xstdoc   = "1"
      end if



      tsel     = request.form("tsel") & ""

      if tsel  = "" then
         tsel  = "C"
         tgrp  = "T"
         tcod  = "99999"
         ttip  = "TODOS"
         tnam  = "TEMA"
         xstsel   = "0"
      else
         xstsel   = "1"

         if tsel  = "" then
            tgrp  = "T"
         end if

      end if

      frmsen = "N"

      if tsel  = "T" then
         tgrp  = "M"
         tcod  = "99999"
         ttip  = "TODOS"
         tnam  = "TODOS"

      end if

      sSectort = request.form("Nlista") & ""

      sSectort    = ofv.convertircar(SsectorT,"_"," ")
      if ssectort = "" then
      '   ssectort = "99999"
         xstsec   = "0"
      else
         xstsec   = "1"
      end if

      textot = request.form("textot") & ""




    formul = request.querystring("frm")

    if len(formul) = 0 then

         tgrp  = request.form("tgrp") & ""
         tcod  = request.form("Ntcod") & ""
         ttip  = request.form("ttip") & ""
         tnam  = request.form("tnam") & ""


    elseif formul = 4 then

      if tsel  <> "T" then
         tgrp   = request.form("ptgrp") & ""
         tcod   = request.form("ptcod") & ""
         ttip   = request.form("pttip") & ""
         tnam   = request.form("ptnam") & ""
         textot = request.form("ptextot") & ""
      end if

    else

      if tsel  <> "T" then
         tgrp  = request.form("tgrp") & ""
         tcod  = request.form("Ntcod") & ""
         ttip  = request.form("ttip") & ""
         tnam  = request.form("tnam") & ""

      end if


    end if

    if  tgrp <> "M" and formul=4 then frmsen = "S"

''''''''' introduccion '''''''''''''''''''''''''''''''''''''


      sInputxx1 = ofv.GenerarInput("Ntipodoc","99999","hidden","","1","dt","")
      sInputxx2 = ofv.GenerarInput("Ntcod","99999","hidden","","1","dt","")
      sInputxx3 = ofv.GenerarInput("Nlista","99999","hidden","","1","dt","")
      sInputxx5 = ofv.GenerarInput("tsel","C","hidden","","1","dt","")
      sInputxx6  = ofv.GenerarInput("tgrp","T","hidden","","1","dt","")
      sInputxx7  = ofv.GenerarInput("ttip","TODOS","hidden","","1","dt","")
      sInputxx8  = ofv.GenerarInput("tnam","TODOS","hidden","","1","dt","")

      sInputxx9b  = ofv.GenerarInput("ttcc","BASIC","hidden","","1","dt","")

      sInputxx9a  = ofv.GenerarInput("ttcc","AVANZADA","hidden","","1","dt","")


      IF TTCC="BASIC" OR TTCC="" THEN
            sInputxb1 = ofv.GenerarInput("bx1","AVANZADA","BUTTON","","1","dt","document.all.Subconsultax1.submit();")
            sInputxb2=""
           ELSE
           sInputxb1=""
           sInputxb2 = ofv.GenerarInput("bx2","BASICA","BUTTON","","1","dt","document.all.Subconsultax2.submit();")
     END IF

    '  sInputxb1 = ofv.GenerarInput("bx1","AVANZADA","BUTTON","","1","dt  ","document.all.Subconsultax1.submit();")
     ' sInputxb2 = ofv.GenerarInput("bx2","BASICA","BUTTON","","1","dt","document.all.Subconsultax2.submit();")



       sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formuln2 = ofv.GenForm("Subconsultax2","",sAccion,"","post","",sInputxx1 & sInputxx2 & sInputxx3 & sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9B,"si")


       sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formuln1 = ofv.GenForm("Subconsultax1","",sAccion,"","post","",sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9A,"si")


       sHTML = sHTML & Formuln2 & Formuln1


      celdan = ""

      celdan = celdan & ofv.GenCelda("","ut ","","","","","",sInputxb1,"si") ' avanzada
      celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb2,"si") 'basica
      Filan  = ofv.GenRow("","","","","","","",celdan,"si")
      Filan  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;display:BLOCK;' ","","20% ","","0","","0","0","","",Filan,"si")
      'Filan  = ofv.GenTabla("","aaa ' ","","20% ","","0","","0","0","","",Filan,"si")

      celda = ofv.GenCelda("","ut COLSPAN=2  ","RIGHT","","","","",Filan,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



      sglmar = "<font id='fnt01' size=4 color='#f7aa22' face='corflaticon' onclick='alert(""Seleccione los parametros para la busqueda...."");' >&#0125;</font>&nbsp;"

      sglpnt = "<font size=2 color='#00aa00' face='corflaticon' >4</font>&nbsp;"

      sinput = sglmar & "Se recomienda acotar la búsqueda mediante la utilización de los filtros, "
      sinput = sinput & "a efecto de hacerla mas rápida, precisa y eficiente."

         sintab = "<script language='javascript' > "
         sintab = sintab & " function opsel(dtfld) {document.all[dtfld + 'tab1'].style.display='none';document.all[dtfld + 'sl'].style.display='block';} "
         sintab = sintab & " function setsel(dtfld) {"
         sintab = sintab & "   document.all.Subconsultax[dtfld].value=document.all.Subconsultax[dtfld + 'sl'].value;"
         sintab = sintab & "   document.all[dtfld +'td1'].innerHTML=document.all.Subconsultax[dtfld + 'sl'].options[document.all.Subconsultax[dtfld + 'sl'].selectedIndex].text;"
         sintab = sintab & "   document.all[dtfld + 'tab1'].style.display='block';document.all[dtfld + 'sl'].style.display='none';"
         sintab = sintab & "   document.all.Subconsultax.submit();"
         sintab = sintab & "                    } "
         sintab = sintab & "</script>"

      sinput = sinput & sintab

      celda = ofv.GenCelda("","utg COLSPAN=2 style='text-align:justify;' ","","","","","",sinput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

      if ttcc=""  then
	     TTCC=SESSION("TTCC")
	     tcod ="99999"
	     sSectort="99999"
	     stipodoc="99999"
	     tsel="C"
	     tgrp="T"
	     ttip="TODOS"
	     tnam="TODOS"
	     xstdoc   = "1"
	     xstsel   = "1"
	     xstsec   = "1"


    end if

    'response.write ttcc
 if ttcc <> "BASIC" THEN



''''''''' tipo de documento '''''''''''''''''''''''''''''''''''''

      sinputtt = "¿Qu&eacute; tipo de objeto est&aacute; Ud. buscando ?"

      celda = ofv.GenCelda("","ttT ","","50%","","","",sinputtt,"si")


      Strsql = "SELEct * FROM TIPOBJETO WHERE TIPOCPT = 'D' and FILTRO = 'S' and aplic='"& WKAPLIC &"' order by ICONO,descri "

      scomK = ""
      scomD = "SELECCIONE DOCUMENTO"

      sinputdtK = ""
      sinputdtd = "SELECCIONE DOCUMENTO"
      sintab2   = ""

      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & rsr(0)
         scomD = scomD & chr(9) & rsr(1)

         sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


         if stipodoc  = cstr(rsr(0)) then
            sinputdtd = rsr(1)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(1) & "</option>"

      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)

         scomK = scomK & chr(9) & "99999"
         scomD = scomD & chr(9) & "* CUALQUIER TIPO DE OBJETO"

         sintab2 = sintab2 & "<option value='99999' "

         if stipodoc  = "99999" then
            sinputdtd = "* CUALQUIER TIPO DE OBJETO"
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >* CUALQUIER TIPO DE OBJETO</option>"


         sintab = "<table id='Ntipodoctab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Ntipodoctd1' width='100%' class='tdcm1' onclick=""opsel('Ntipodoc');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Ntipodoctd2' width='1'    class='tdcm2' onclick=""opsel('Ntipodoc');"" >"
         sintab = sintab & "<font class='btic'  >6<font></td>"
         sintab = sintab & "</table>"


         sintab3 = "<select name='Ntipodocsl' size='5' onchange=""setsel('Ntipodoc');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Ntipodoc",stipodoc,"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK

      Celda = Celda & ofv.GenCelda("","ut  ","center","50%","","","",sInput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")





'''''''''''' carpetas ''''''''''''''''''''''''''''''''''''''''''''''''

  if  xstdoc   = "1" then

      sinputtt = "¿En qu&eacute; carpeta desea buscarlo ?"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")

      Strsql = "SELEct * FROM tema where PSTIPO = 'C' order by inda, descri "

      scomK = ""
      scomD = "SELECCIONE CARPETA"

      sinputdtK = ""
      sinputdtd = "SELECCIONE CARPETA"
      sintab2   = ""

      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & rsr(0)
         scomD = scomD & chr(9) & rsr(1)

         sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


         if tcod  = cstr(rsr(0)) then
            tnam       = replace(rsr(1)," ","_")
            sinputdtd  = rsr(1)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(1) & "</option>"


      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)

         scomK = scomK & chr(9) & "99999"
         scomD = scomD & chr(9) & "* TODAS LAS CARPETAS"

         sintab2 = sintab2 & "<option value='99999' "

         if tcod  = "99999" then
            tnam       =  "TODAS_LAS_CARPETAS"
            sinputdtd  = "* TODAS LAS CARPETAS"
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >* TODAS LAS CARPETAS</option>"


         sintab = "<table id='Ntcodtab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Ntcodtd1' width='100%' class='tdcm1' onclick=""opsel('Ntcod');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Ntcodtd2' width='1'    class='tdcm2' onclick=""opsel('Ntcod');"" >"
         sintab = sintab & "<font class='btic' >6<font></td>"
         sintab = sintab & "</table>"


         sintab3 = "<select name='Ntcodsl' size='5' onchange=""setsel('Ntcod');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Ntcod",tcod,"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK


      sInput1  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
      sInput3  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")

      Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

  end if



'''''''''' sector '''''''''''''''''''''''''''''''''''''''''''''''''


  if  xstdoc   = "1" and   xstsel   = "1" then

 if  stipodoc = "99999" or stipodoc = "PCD" then


      Strsql = "SELECT DISTINCT ECO.descri FROM OBJSECTOR, ECO,  OBJETOS "


      if SESSION("EMPMUN") = "S" THEN
         if SESSION("GLBMUN") = 2 THEN
            Strsql = Strsql & " ,OBJMEMPALC "
         END IF
      END IF

      Strsql = Strsql & " where OBJSECTOR.codsector = ECO.eco AND ECO.ECOSITE = " & SEMPRE
      Strsql = Strsql & " AND OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO  "

      if SESSION("EMPMUN") = "S" then
         Strsql = Strsql & " AND "
         if SESSION("GLBMUN") = 1 THEN
            Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
         ELSE
            Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
            Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
            Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
            Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
            Strsql = Strsql & " ) "
         END IF
      END IF


      Strsql  = Strsql & " ORDER BY ECO.descri"


      scomK = ""
      scomD = "SELECCIONE SECTOR"

      sinputdtK = ""
      sinputdtd = "SELECCIONE SECTOR"
      sintab2   = ""


      scomK = scomK & chr(9) & "99999"
      scomD = scomD & chr(9) & "* TODOS LOS SECTORES"

         sintab2 = sintab2 & "<option value='99999' "

         if ofv.convertircar(SsectorT," ","_")  = "99999" then
            sinputdtd = "* TODOS LOS SECTORES"
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >* TODOS LOS SECTORES</option>"


      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & ofv.convertircar(rsr(0)," ","_")
         scomD = scomD & chr(9) & rsr(0)


         sintab2 = sintab2 & "<option value='" & ofv.convertircar(rsr(0)," ","_") & "' "


         if ofv.convertircar(SsectorT," ","_")  = ofv.convertircar(rsr(0)," ","_") then
            sinputdtd = rsr(0)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(0) & "</option>"


      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)


         sintab = "<table id='Nlistatab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Nlistatd1' width='100%' class='tdcm1' onclick=""opsel('Nlista');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Nlistatd2' width='1'    class='tdcm2' onclick=""opsel('Nlista');"" >"
         sintab = sintab & "<font class='btic' >6<font></td>"
         sintab = sintab & "</table>"


         sintab3 = "<select name='Nlistasl' size='5' onchange=""setsel('Nlista');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Nlista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK

      sinputtt = "¿Desea acotar la b&uacute;squeda a alg&uacute;n sector en particular ?"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")
      Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

 else

      SsectorT = "99999"

      xstsec   = "1"

 end if

 else

      SsectorT = "99999"

 end if

 END IF

 

'''''''''''' texto a buscar '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if  xstdoc   = "1" and   xstsel   = "1" and   xstsec   = "1"   then

      if textot = "" then
         textot = "INGRESE EL TEXTO A BUSCAR"
      end if

      sinputtt = "Ingrese el texto a buscar"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")


 end if

       sInputTTCC  = ofv.GenerarInput("TTCC",ttcc,"hidden","","1","dt","")


       sAccion = "ORSUBCONSULTA.ASP?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"","post","",Fila & sInputTTCC ,"si")



       Fila    = ""

         if tgrp  = "" then
            tgrp  = "T"
         end if

      sInput0  = ofv.GenerarInput("tdoc",stipodoc,"hidden","","1","dt","")
      sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
      sInput5  = ofv.GenerarInput("tcod",tcod,"hidden","","1","dt","")
      sInput6  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
      sInput7  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")




    sInput3     = ofv.GenerarInput("buscar",textot,"agregar","","50","dt","")





    if  textot <> "" then

        sInput     = "&nbsp;"
        Celda      = Celda & ofv.GenCelda("","UT","center","","","","",sInput0 & sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7,"si")
        Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    end if





      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren




  elseif sNombre = "CONSULTA POR BUSCADOR" then






 



      if textot = "" then
         textot = "TEXTO A BUSCAR"
      end if

      sinputtt = "INGRESE EL TEXTO A BUSCAR EN"

      sInput3     = ofv.GenerarInput("buscar",textot,"agregar","","50","dt","")

      Celda  = ofv.GenCelda("","vt","CENTER","30%","","","",sinputtt,"si")
      Celda  = Celda & ofv.GenCelda("","vti","","70%","","","",sInput3,"si")
      Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

       sInputTTCC   = ofv.GenerarInput("","Buscar","submit","","10","bt","")

  sImage  = ofv.GenImage("","","","","IMAGENES\GOOGLE.jpg","0","")

      Celda  = ofv.GenCelda("","vt","CENTER","30%","","","",sImage,"si")
      Celda  = Celda & ofv.GenCelda("","vti","","70%","","","",sInputTTCC,"si")
      Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

       sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"'_blank'","post","",Fila ,"si")

      Fila  = ""
       sAccion = ""
      formx  = ""




  elseif sNombre = "CONSULTA POR CATEGORIAS" then

      sInput  = ""
      sInput2  = ""
      sInput3  = ""
      sInput4  = ""
      sInput5  = ""
      sInput6  = ""
      sInput7  = ""

      xstdoc   = "0"
      xstsel   = "0"
      xstsec   = "0"

''response.write "<br><br><br>---------------------------"
'for each x in request.form
 '   response.write "-- " & x & " : " & request.form(x) & " --<br>"
'next
''response.write "---------------------------"


      xx1 = request.form("Ntipodoc") & ""
      xx2 = request.form("Ntcod") & ""
      xx3 = request.form("Nlista") & ""
      xx4 = request.form("textot") & ""
      xx5 = request.form("tsel") & ""

      ttcc = request.form("ttcc") & ""





      stipodoc    = request.form("Ntipodoc") & ""
      if stipodoc = "" then
      '   stipodoc = "99999"
         xstdoc   = "0"
      else
         xstdoc   = "1"
      end if



      tsel     = request.form("tsel") & ""

      if tsel  = "" then
         tsel  = "C"
         tgrp  = "T"
         tcod  = "99999"
         ttip  = "TODOS"
         tnam  = "TEMA"
         xstsel   = "0"
      else
         xstsel   = "1"

         if tsel  = "" then
            tgrp  = "T"
         end if

      end if

      frmsen = "N"

      if tsel  = "T" then
         tgrp  = "M"
         tcod  = "99999"
         ttip  = "TODOS"
         tnam  = "TODOS"

      end if

      sSectort = request.form("Nlista") & ""

      sSectort    = ofv.convertircar(SsectorT,"_"," ")
      if ssectort = "" then
      '   ssectort = "99999"
         xstsec   = "0"
      else
         xstsec   = "1"
      end if

      textot = request.form("textot") & ""








    formul = request.querystring("frm")

    if len(formul) = 0 then

         tgrp  = request.form("tgrp") & ""
         tcod  = request.form("Ntcod") & ""
         ttip  = request.form("ttip") & ""
         tnam  = request.form("tnam") & ""


    elseif formul = 4 then

      if tsel  <> "T" then
         tgrp   = request.form("ptgrp") & ""
         tcod   = request.form("ptcod") & ""
         ttip   = request.form("pttip") & ""
         tnam   = request.form("ptnam") & ""
         textot = request.form("ptextot") & ""
      end if

    else

      if tsel  <> "T" then
         tgrp  = request.form("tgrp") & ""
         tcod  = request.form("Ntcod") & ""
         ttip  = request.form("ttip") & ""
         tnam  = request.form("tnam") & ""

      end if


    end if




    if  tgrp <> "M" and formul=4 then frmsen = "S"

''''''''' introduccion '''''''''''''''''''''''''''''''''''''


      sInputxx1 = ofv.GenerarInput("Ntipodoc","99999","hidden","","1","dt","")
      sInputxx2 = ofv.GenerarInput("Ntcod","99999","hidden","","1","dt","")
      sInputxx3 = ofv.GenerarInput("Nlista","99999","hidden","","1","dt","")
      sInputxx5 = ofv.GenerarInput("tsel","C","hidden","","1","dt","")
      sInputxx6  = ofv.GenerarInput("tgrp","T","hidden","","1","dt","")
      sInputxx7  = ofv.GenerarInput("ttip","TODOS","hidden","","1","dt","")
      sInputxx8  = ofv.GenerarInput("tnam","TODOS","hidden","","1","dt","")




      sInputxx9b  = ofv.GenerarInput("ttcc","BASIC","hidden","","1","dt","")

      sInputxx9a  = ofv.GenerarInput("ttcc","AVANZADA","hidden","","1","dt","")


      sInputxb1 = ofv.GenerarInput("bx1","AVANZADA","BUTTON","","1","dt","document.all.Subconsultax1.submit();")
      sInputxb2 = ofv.GenerarInput("bx2","BASICA","BUTTON","","1","dt","document.all.Subconsultax2.submit();")






       sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formuln2 = ofv.GenForm("Subconsultax2","",sAccion,"","post","",sInputxx1 & sInputxx2 & sInputxx3 & sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9B,"si")


       sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formuln1 = ofv.GenForm("Subconsultax1","",sAccion,"","post","",sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9A,"si")


       sHTML = sHTML & Formuln2 & Formuln1


      celdan = ""
      celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb1,"si")
      celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb2,"si")
      Filan  = ofv.GenRow("","","","","","","",celdan,"si")
      Filan  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;display:none;' ","","20% ","","0","","0","0","","",Filan,"si")


      celda = ofv.GenCelda("","ut COLSPAN=2  ","RIGHT","","","","",Filan,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



      sglmar = "<font id='fnt01' size=4 color='#f7aa22' face='corflaticon' onclick='alert(""Seleccione los parametros para la busqueda...."");' >&#0125;</font>&nbsp;"

      sglpnt = "<font size=2 color='#00aa00' face='corflaticon' >4</font>&nbsp;"

      sinput = ""


         sintab = "<script language='javascript' > "
         sintab = sintab & " function opsel(dtfld) {document.all[dtfld + 'tab1'].style.display='none';document.all[dtfld + 'sl'].style.display='block';} "
         sintab = sintab & " function setsel(dtfld) {"
         sintab = sintab & "   document.all.Subconsultax[dtfld].value=document.all.Subconsultax[dtfld + 'sl'].value;"
         sintab = sintab & "   document.all[dtfld +'td1'].innerHTML=document.all.Subconsultax[dtfld + 'sl'].options[document.all.Subconsultax[dtfld + 'sl'].selectedIndex].text;"
         sintab = sintab & "   document.all[dtfld + 'tab1'].style.display='block';document.all[dtfld + 'sl'].style.display='none';"
         sintab = sintab & "   document.all.Subconsultax.submit();"
         sintab = sintab & "                    } "
         sintab = sintab & "</script>"

      sinput = sinput & sintab

      celda = ofv.GenCelda("","utg COLSPAN=2 style='text-align:justify;' ","","","","","",sinput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")





 if ttcc <> "BASIC" THEN



''''''''' tipo de RIESGO (SITUACION) '''''''''''''''''''''''''''''''''''''

      sinputtt = "SITUACION"

      celda = ofv.GenCelda("","ttT ","","50%","","","",sinputtt,"si")


      Strsql = "SELEct * FROM RSKSITUACION  order by descri "

      scomK = ""
      scomD = "SELECCIONE SITUACION"

      sinputdtK = ""
      sinputdtd = "SELECCIONE SITUACION"
      sintab2   = ""

      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & rsr(0)
         scomD = scomD & chr(9) & rsr(1)

         sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


         if stipodoc  = cstr(rsr(0)) then
            sinputdtd = rsr(1)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(1) & "</option>"

      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)

   '      scomK = scomK & chr(9) & "99999"
   '      scomD = scomD & chr(9) & "* CUALQUIER TIPO DE OBJETO"

   '      sintab2 = sintab2 & "<option value='99999' "

   '      if stipodoc  = "99999" then
   '         sinputdtd = "* CUALQUIER TIPO DE OBJETO"
   '         sintab2    = sintab2 & " selected "
   '      end if

   '      sintab2 = sintab2 & " >* CUALQUIER TIPO DE OBJETO</option>"


         sintab = "<table id='Ntipodoctab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Ntipodoctd1' width='100%' class='tdcm1' onclick=""opsel('Ntipodoc');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Ntipodoctd2' width='1'    class='tdcm2' onclick=""opsel('Ntipodoc');"" >"
         sintab = sintab & "<font class='btic'  >6<font></td>"
         sintab = sintab & "</table>"


         sintab3 = "<select name='Ntipodocsl' size='5' onchange=""setsel('Ntipodoc');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Ntipodoc",stipodoc,"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK

      Celda = Celda & ofv.GenCelda("","ut  ","center","50%","","","",sInput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")





'''''''''''' CATEGORIAS ''''''''''''''''''''''''''''''''''''''''''''''''

  if  xstdoc   = "1" then

      sinputtt = "CATEGORIA"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")

      Strsql = "SELEct * FROM RSKCATEGO WHERE SITUACION = 0" & stipodoc & " order by descri "

      scomK = ""
      scomD = "SELECCIONE CATEGORIA"

      sinputdtK = ""
      sinputdtd = "SELECCIONE CATEGORIA"
      sintab2   = ""

      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & rsr(0)
         scomD = scomD & chr(9) & rsr(2)

         sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


         if tcod  = cstr(rsr(0)) then
            tnam       = replace(rsr(2)," ","_")
            sinputdtd  = rsr(2)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(2) & "</option>"


      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)

     '    scomK = scomK & chr(9) & "99999"
     '    scomD = scomD & chr(9) & "* TODAS LAS CARPETAS"

     '    sintab2 = sintab2 & "<option value='99999' "

     '    if tcod  = "99999" then
     '       tnam       =  "TODAS_LAS_CARPETAS"
     '       sinputdtd  = "* TODAS LAS CARPETAS"
     '       sintab2    = sintab2 & " selected "
     '    end if

     '    sintab2 = sintab2 & " >* TODAS LAS CARPETAS</option>"


         sintab = "<table id='Ntcodtab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Ntcodtd1' width='100%' class='tdcm1' onclick=""opsel('Ntcod');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Ntcodtd2' width='1'    class='tdcm2' onclick=""opsel('Ntcod');"" >"
         sintab = sintab & "<font class='btic' >6<font></td>"
         sintab = sintab & "</table>"


         sintab3 = "<select name='Ntcodsl' size='5' onchange=""setsel('Ntcod');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Ntcod",tcod,"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK


      sInput1  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
      sInput3  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")

      Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

  end if



'''''''''' sector '''''''''''''''''''''''''''''''''''''''''''''''''


  if  xstdoc   = "1" and   xstsel   = "1" then

' if  stipodoc = "99999" or stipodoc = "PCD" then

      Strsql = "SELEct * FROM RSKEVENTO WHERE CATEGO = 0" & tcod & " order by descri "

      scomK = ""
      scomD = "SELECCIONE EVENTO"

      sinputdtK = ""
      sinputdtd = "SELECCIONE EVENTO"
      sintab2   = ""





      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & ofv.convertircar(rsr(0)," ","_")
         scomD = scomD & chr(9) & rsr(3)


         sintab2 = sintab2 & "<option value='" & ofv.convertircar(rsr(0)," ","_") & "' "


         if ofv.convertircar(SsectorT," ","_")  = ofv.convertircar(rsr(0)," ","_") then
            sinputdtd = rsr(3)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(3) & "</option>"


      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)


         sintab = "<table id='Nlistatab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Nlistatd1' width='100%' class='tdcm1' onclick=""opsel('Nlista');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Nlistatd2' width='1'    class='tdcm2' onclick=""opsel('Nlista');"" >"
         sintab = sintab & "<font class='btic' >6<font></td>"
         sintab = sintab & "</table>"


         sintab3 = "<select name='Nlistasl' size='5' onchange=""setsel('Nlista');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Nlista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK

      sinputtt = "EVENTO DE PERDIDAS"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")
      Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



 else

      SsectorT = "99999"

 end if

 END IF

'''''''''''' texto a buscar '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if  xstdoc   = "1" and   xstsel   = "1" and   xstsec   = "1"  then

      if textot = "" then
         textot = " "
      end if

      sinputtt = "CONFIRMA SELECCION DE BUSQUEDA?"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")


 end if


       sAccion = "ORSUBCONSULTA.ASP?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"","post","",Fila,"si")

'' END IF

       Fila    = ""

         if tgrp  = "" then
            tgrp  = "T"
         end if

      sInput0  = ofv.GenerarInput("tdoc",stipodoc,"hidden","","1","dt","")
      sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
      sInput5  = ofv.GenerarInput("tcod",tcod,"hidden","","1","dt","")
      sInput6  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
      sInput7  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")




    sInput3     = ofv.GenerarInput("buscar",textot,"HIDDEN","","50","dt","")





    if  textot <> "" then

        sInput     = "&nbsp;"
        Celda      = Celda & ofv.GenCelda("","UT","center","","","","",sInput0 & sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7,"si")
        Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    end if



      formx    = Fila



      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


  elseif sNombre = "CONSULTA POR FECHA" then


    Celda      = ofv.GenCelda("","tt","","","","","","tipo","si")

    scomK = "Emision" & chr(9) & "Vigencia"  & chr(9) & "Publicacion"  & chr(9) & "Vencimiento" & chr(9) & "Eliminacion"
    scomD = "Emision" & chr(9) & "Vigencia"  & chr(9) & "Publicacion"  & chr(9) & "Revision" & chr(9) & "Eliminacion"

    sInput = ofv.gencomboFX("subc","fecha","Emision",scomK,scomD,"no")



    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    Celda      = ofv.GenCelda("","tt","","","","","","Desde","si")
    sInput     = ofv.GenerarInput("desde","","agregar","","20","dt","")
    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    Celda      = ofv.GenCelda("","tt","","","","","","Hasta","si")
    sInput     = ofv.GenerarInput("hasta","","agregar","","20","dt","")
    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren

  elseif sNombre = "CONSULTA POR TEMA" then

   response.redirect "OBJtreeTema.asp?Tipo=Consulta&Nombre=TEMAS&Menu=Consulta&EMPRE=" & SEMPRE

'Listas de Complementos
  elseif sNombre = "CONSULTA POR COMPONENTE" then


    Celda      = ofv.GenCelda("","tt","","","","","","Descripcion","si")
    sInput     = ofv.GenerarInput("buscar","","agregar","","60","dt","")
    Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
    Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      sInput   = ofv.GenerarInput("","Consultar","submit","","10","bt","")
      sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      Fila     = Fila & ofv.GenRow("","","","","40","","",Celda,"si")

      Celda    = ofv.GenCelda("","ttt colspan=2 ","","","","","","&nbsp;","si")
      Fila     = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

    sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0" & "&Nombre=" & sNombren
    Formulario = Formulario & ofv.GenForm("buscar1","",sAccion,"","post","",Fila,"si")

      Fila    = ""


      SQL      = "Select descri,tipo From Tipocomp Order By descri"
      Celda    = ofv.GenCelda("","tt","","","","","","Tipo","si")
      sInput   = ofv.GenerarComboSP("radio","XComplem",xComp,SQL,cn,1,0,"XComplem")
      Celda    = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


       sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
    Formulario = Formulario  & ofv.GenForm("radio","",sAccion,"","post","",Fila,"si")

      Fila    = ""



      Celda    = ofv.GenCelda("","tt","","","","","","Componente","si")

      Strsql   = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
      Strsql   = Strsql & "COMPLEMENTOS.descri FROM COMPLEMENTOS, "
      Strsql   = Strsql & " objlnk where COMPLEMENTOS.codcomp = objlnk.codobjetol and objlnk.ltipo = 'KL' "
      Strsql   = Strsql & "and Complementos.tipocomp = '" & xComp & "' "
      Strsql   = Strsql & "ORDER BY COMPLEMENTOS.descri"

      sInput   = ofv.GenerarLista(Strsql,cn,2)
      Celda    = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
     'response.write strsql

      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


'Listas de Sectores
    elseif sNombre = "CONSULTA POR SECTOR" then

      sSectort = request.form("lista") & ""
      sSectort = ofv.convertircar(SsectorT,"_"," ")

      Strsql = "SELECT DISTINCT ECO.descri FROM OBJSECTOR, ECO,  OBJETOS "


                      if SESSION("EMPMUN") = "S" THEN
                         if SESSION("GLBMUN") = 2 THEN
                            Strsql = Strsql & " ,OBJMEMPALC "
                         END IF
                      END IF

      Strsql = Strsql & " where OBJSECTOR.codsector = ECO.eco AND ECO.ECOSITE = " & SEMPRE
      Strsql = Strsql & " AND OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO  "

      if SESSION("EMPMUN") = "S" then
         Strsql = Strsql & " AND "
         if SESSION("GLBMUN") = 1 THEN
            Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
         ELSE
            Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
            Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
            Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
            Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
            Strsql = Strsql & " ) "
         END IF
      END IF

      Strsql2 = Strsql & " AND ECO.DESCRI = '" &  sSectort & "' "
      Strsql  = Strsql & " ORDER BY ECO.descri"
     '' response.write strsql

      set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
      if not rsr.eof then
         sSectort=rsr(0)
      end if
      call ofv.cerrarconsulta(rsr)


      sInput   = ofv.GenerarComboSP("Subconsulta1","lista",sSectort,Strsql,cn,0,0," ")


         celda = ofv.GenCelda("","tt","","","","","","Sector","si")
         Celda = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
         Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

       sAccion = "ORSUBCONSULTA.ASP?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
    Formulario = Formulario  & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

       Fila    = ""


      x2sql="SELECT OBJSECTOR.objid FROM OBJSECTOR, ECO, OBJETOS "


                      if SESSION("EMPMUN") = "S" THEN
                         if SESSION("GLBMUN") = 2 THEN
                            x2sql=x2sql & " ,OBJMEMPALC "
                         END IF
                      END IF

      x2sql=x2sql & " where OBJSECTOR.codsector = ECO.eco "
      x2sql=x2sql & " AND OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO "

      if SESSION("EMPMUN") = "S" then
         x2sql=x2sql & " AND "

         if SESSION("GLBMUN") = 1 THEN
            x2sql=x2sql & " OBJETOS.OBJUN = " & SEMPRE
         ELSE
            x2sql=x2sql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
            x2sql=x2sql & "   ( OBJETOS.OBJUN = 1 AND "
            x2sql=x2sql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
            x2sql=x2sql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
            x2sql=x2sql & " ) "
        END IF
      END IF
      x2sql=x2sql & " and ECO.descri='" & ofv.convertircar(SsectorT,"_"," ") & "'"
   ''   response.write x2sql & "<BR>"
      set rs = ofv.crearconsultaEx(x2sql,cn,1,parametros)
      if not rs.eof then
         sSector=rs(0)

          xsql="SELECT tipoperfil,TIPOPERFIL.descri "
          xsql =xsql & " FROM Objsector, ECO, ECOPERFIL, TIPOPERFIL "
          xsql=xsql & " where Eco.eco = Objsector.codsector and "
          xsql=xsql &  "ECO.eco = ECOPERFIL.codsector and "
          xsql=xsql & "ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
          xsql=xsql & "and objsector.objid =" & rs(0)
          call ofv.cerrarconsulta(rs)

          sInput = ofv.GenerarCombo("Subconsulta","perfil","nsnc",Xsql,cn,0,1,"no")

      else

          sInput = "&nbsp;" & ofv.GenerarInput("perfil","nsnc","hidden","","1","dt","")

      end if

      sInput   = sinput & ofv.GenerarInput("lista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")
      celda    = ofv.GenCelda("","tt","","","","","","Perfil","si")
      Celda    = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
      Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


'Listas de Estados
  else
        Strsql = "Select * From Estados order by estado"
        sInput = ofv.GenerarCombo2("lista","1",Strsql,cn,0,1,"sb")
        celda  = ofv.GenCelda("","tt","","","","","","Estado","si")
        Celda  = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
        Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


      formx    = Fila

      sTARGET  = ""
      sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


'Fin preparacion
  end if


'''' Armado del Form de despacho '''''''''''''''''''''''

 '   if formx <> "22" then
 '     Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
 '     sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
 '     sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
 '     Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
 '     formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")
 '   sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
 '   Formulario = Formulario & ofv.GenForm("Subconsulta","",sAccion,"","post","",formx,"si")
 '   end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    if formx <> "" then

       IF sAccion = "" THEN
          sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
       END IF


       Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
       sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
       sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
       Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
       formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")


       Formulario = Formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post","",formx,"si")

    end if



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

     sAccion = "ORSUBCONSULTA.ASP?frm=4&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
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