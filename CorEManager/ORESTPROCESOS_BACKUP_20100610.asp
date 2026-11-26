<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

response.expires = 0
Server.ScriptTimeout = 120

sModal  = request.querystring("modal")
sTipo   = request.querystring("Tipo")
sNombre = request.querystring("Nombre")
sNombre = ofv.convertircar(sNombre,"_"," ")
sMenu   = request.querystring("Menu")
sMETODO   = request.querystring("METODO")
sACTIVIDAD   = request.querystring("ACTIVIDAD")
SACCIONEJEC   = request.querystring("ACCIONEJEC")
SACCIONEJECCOD   = request.querystring("ACCIONEJECCOD")
SACCIONEJECTIP   = request.querystring("ACCIONEJECTIP")


if sModal = "" then sModal = "LISTA"



session("ModMain") = "ED"


xapli = 1

  on error resume next
  set cn = ofv.conectar(ofv.strconn1)
  Strsql = "Select ID From CorAplic where codapli = '" & Trim(Session("CorApli")) & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     xapli = rs(0)
  end if

  Strsql = "Select menu From menux where menuatr = '" & snombre & "' and aplicacion = " & xapli
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sNombre = rs(0)
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



''''''''''''''''''''''''''''

  session("xorigen") = "tema"

  XXXX = NEWEMPRE("NO")

  sEMPRE = request.querystring("EMPRE")

  IF SEMPRE = "" THEN
     SEMPRE = ATREMP(0,0)
  ELSE
     recarga = NEWEMPRE(SEMPRE)
  END IF


  sUN = SEMPRE

  if sUN = "" then sUN = 1
  session("UN") = sUN


  LNombre = ofv.convertircar(sNombre," ","_")

  if sMenu <> "" then Session("Menu") = sMenu

  sID = request.querystring("ID")
  if sNombre = "" Then sNombre = UCase(sTipo)
  sitem      = ofv.convertircar(request.querystring("sitem"),"_"," ")
  stitem     = ofv.convertircar(request.querystring("stitem"),"_"," ")
  sticono    = request.querystring("sticono")
  stnivel    = request.querystring("stnivel")




  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "Select estado from estados where objid = " & session("sEstadon")
 ' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     ssTipo = rs(0)
  end if
  call ofv.cerrarconsulta(rs)




  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

  FUNCION = "  onload="  & Chr(34)

  if  recarga = "S" AND sTipo <> "Consulta" then
      FUNCION = FUNCION & "document.recarga.submit();"
  end if

  FUNCION = FUNCION & "rszdx4();"  & Chr(34)
 if sModal <> "LISTA" THEN
  FUNCION = FUNCION & " onmousedown = " & Chr(34) & "Abre();" & Chr(34) & " "
 end if
  FUNCION = FUNCION & " onresize = " & Chr(34) & "rszdx4();" & Chr(34) & " "

  sHTML = ofv.MENUheader("",funcion)




           shtml=shtml & "<script language=javascript>" & chr(13)
           shtml=shtml & "function mostrarmenuD(){" & chr(13)
           shtml=shtml & "document.all.Dx1.style.display='block';" & chr(13)
           shtml=shtml & "}" & chr(13)
           shtml=shtml & "function rszdx4(){" & chr(13)
           shtml=shtml & "document.all.DX2.style.display='block';" & chr(13)

           shtml=shtml & "}" & chr(13)
           shtml=shtml & "</script>" & chr(13)

 if sModal <> "LISTA" THEN

	           shtml=shtml  & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop2.js' ></script>"
 END IF

  sHTML = sHTML & ofv.CorFormHeader

  sNombre = ofv.convertircar(sNombre," ","_")

  HrfAsp = "EMDETAPROCPROPER.asp?"
  HrfTrg = "fesquemas"


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

      sHref = "OBJtree.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
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
  Tabla = ""


  sHref = ""
  sImage = ofv.GenImage("","","","","Iconos\ICorMMenu.jpg","0","")

  sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")

  sLink = ofv.GenLink("","nt",sHref,"ftemas","","",sImage,"si")
  Celda = ofv.GenCelda("","aaa colspan=3 rowspan=2  style='" & Session("CBRDICO") & "'","center","2%","","","",sLink,"si")



  xxx = "<b><u>" & Ucase(sstipo) & "</u></b>"

  sLink = ofv.GenLink("","cc",sHref,"ftemas","","","<b><u>" & ATREMP(0,4) & "</u></b>","si")
  Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2","","","","","",sLink,"si")
  Celda = Celda & ofv.GenCelda("","","","","","","",SINPUT3,"si")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' set de permisos
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




'''''''''''''''''''''''''''''''''''''''''''
            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = ""
                WKSPSDKGRP = ""

                for XMFI = 0 to ubound(XMFTGRP)
                    WKSPSDKGRP = WKSPSDKGRP  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = ","
                next

            else
                WKSPSDKGRP = "'GRP-" & Session("Seggrp") & "'"
            end if


''''''''''''''''''''''''''''''''''''''''''



  stitem = "Tema"
  sIcono = "ICorTema.jpg"
  sNuevoTipo = "SUBTEMAS"
  nNivel = 0

  sID = 0
  nv = 0
  cborderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("APLSHW") & ";"
  cborderl = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("APLSHW") & ";"
  borderB = ""
  borderl = ""


'''''''''''''''''''''''''''''''''''''''''''
 ''' session("menu") = 3
  scodmaster = request.querystring("codmaster") & ""
  stipmaster = request.querystring("tipmaster") & ""
  sCODNIVEL  = request.querystring("CODNIVEL")

  if scodmaster = "" OR stipmaster = "PROCESO" then
     scodmaster = "000"
     stipmaster = "PROCESO"
     sCODNIVEL  = 1

     SESSION("OROBJ") = 0




  end if




   Session("ASP")     = "vuelveprocesos.asp"
   Session("ASPSTR")  = wkproccod
   Session("ASPSTRD") = Session("ASPSTR")








''''''''''''''''''''''''''''''''''''''''''''

 sImageCRT = "<font face='corflaticon' size=2 color='#0000ff' style='cursor:pointer;border-style:OUTSET;border-width:1;' title='Tiene circuito definido.' >&#061;</font>"

 SESSION("ORPERIODOST") = XFORPERIODOST
 SESSION("ORRSKST") = 1

 wspTIPO = "SP"
 wkproccodL = "&nbsp;"
 wkproccodL2 = "&nbsp;"
 wkproccodN = "&nbsp;"



  DIM TABSPPERM(100,3)
  IPERM = 0
  TPERM = 0


  DIM TABMAPA(10,10)

  TTABMAPA = 0

  STRSQL = ""
  STRSQL = STRSQL & " SELECT * FROM ORPERIODOS "
  STRSQL = STRSQL & " WHERE ESTADO = 'OK' AND APLIC = 'OR'  "
  STRSQL = STRSQL & " ORDER BY INICIO DESC, OBJID DESC "
 ' response.write strsql
      set rsX5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX5.eof then

         DO WHILE NOT RSX5.EOF

            TTABMAPA = TTABMAPA + 1
            TABMAPA(TTABMAPA,00) = RSX5(0)
            TABMAPA(TTABMAPA,01) = RSX5(1)
            TABMAPA(TTABMAPA,02) = RSX5(2)
            TABMAPA(TTABMAPA,03) = 0
            TABMAPA(TTABMAPA,04) = 0
            TABMAPA(TTABMAPA,05) = 0


         RSX5.MOVENEXT
         LOOP



      END IF


    wkprocSPW = "&nbsp;"
    WKSPPERMISOS = "&nbsp;"
    WKSPTIPOPR = "NO"

      WKSPMASTCT = "S"
      WKSPMASTMO = "S"
      WKSPMASTCO = "S"
      WKSPMASTCE = "S"


      WKSPCIRCCT = "N"
      WKSPCIRCMO = "N"
      WKSPCIRCCO = "N"
      WKSPCIRCCE = "N"





    WKPROCDESC = "PRINCIPAL DEL NEGOCIO&nbsp;(NIVEL&nbsp;" & sCODNIVEL & ")"


    wkproccod = "&nbsp;"
    wkproccod2 = "&nbsp;"
    wspTIPO = "SP"

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

''WKSPPERMISOS





                        sInput = "PROCESO"
                        wspTIPOCLC = "#993333"




    sImagePRC = "<font face='corflaticon' size=2 color='" & wspTIPOCLC & "'  >&#060;</font>"

    WKPROCDESC = WKDIVSECUEN & sInput & ":&nbsp;" & WKPROCDESC


    Celda = ""

   Celda = Celda & ofv.GenCelda("","Ut style='background-color:#f0f0f0;color:#990000;' ","CENTER","1%","","","",sImagePRC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","",WKPROCDESC,"si")
   'Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","","hOLA A TODOS ","si")

   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = ofv.GenTabla("","aaa style='background-color:#f0eedd;border-style:solid;border-width:1;border-color:#000000;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

''''''''''''''''''''''''''''''''''''''''''''
   TablaVH = ""



''''''''''''''''''''''''''''''''''''''''''''



   Fila = ""

          wspcolor1  = "#990000"
          wspcolor2  = "#009900"
          sImageO1   = "&#0253;"
          sImageO2   = "&#0254;"







      Strsql = "Select A.* From OBJETOS A"
      Strsql = Strsql & " INNER JOIN relaobj B ON ( B.CODOJASOC = A.CODOBJETO )"


      Strsql = Strsql & " where  "
      Strsql = Strsql & "  B.codobjmaster = '" & scodmaster & "' and B.tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & "  AND (A.CODVINOBJ = 'SP' or A.CODVINOBJ = 'SPI') "
      Strsql = Strsql & " ORDER BY A.TSEGMENTO, A.TAGRUPA, A.INDOBJ "


' response.write strsql & "<br>" & chr(13)




'  response.write strsql & "<br>" & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO HAY PROCESOS, ACTIVOS O ACTIVIDADES ASOCIADAS","si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
     TablaN = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

     Tabla = "<CENTER><BR><BR>" & Tabla & "<BR><BR>" & TablaN & "</CENTER>"

  else





    xTrColor = "#FFFFFF"

    wkcant = 1
    wkleft = 5
    wktop  = 5

    DHOY = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))


       Celda     = ""
       Celda = Celda & ofv.GenCelda("","utI COLSPAN=2 STYLE='COLOR:#FFFFFF;' ","","","","","","PROCESO","si")
       Celda = Celda & ofv.GenCelda("","UtI STYLE='COLOR:#FFFFFF;' ","","","","","","FRECUENCIA","si")
       Celda = Celda & ofv.GenCelda("","UtI COLSPAN=2 STYLE='COLOR:#FFFFFF;' ","","","","","","ULTIMA EVALUACION","si")

       FilaN = filaN & ofv.GenRow("","AAA bgcolor='#888888' ","","","20","","",Celda,"si")



    do until rs.eof

       wkcoddelete = "N"
       wktitdelete = " "

''''''''''''''''''''''''''''''

     XFNOVECRT = ""
     XFNOVEPROC = ""



     WKSPESTADO = 3


          wspcolor = "#993333"
          sImageP    = "&#0167;"



       wspcolor2 = wspcolor

       sImageP    =  sImagePER & "<font face='corflaticon' size=2 color='" & wspcolor2 & "'>" & sImageP & "</font>"



       wkproccod = rs(3)

       wkproccodFC = CLNG("0" & rs(10))


       SELECT CASE(RS(9) & "")

              CASE "1"

                    wkproccodF = "<B>" & wkproccodFC & "&nbsp;MES/ES" & "</B>"

                    WSPCANTDIAS = wkproccodFC * 30


              CASE "2"

                    wkproccodF = "<B>" & wkproccodFC & "&nbsp;AÑO/S" & "</B>"

                    WSPCANTDIAS = wkproccodFC * 360


              CASE ELSE

                    wkproccodF = "<I>SIN DEFINIR</I>"

                    WSPCANTDIAS = 0


       END SELECT

      wkproccodPER = "<I>SIN EVALUACIONES</I>"
      WSPDIF       = 9999
      WSPEVAL      = "N"

      Strsql = ""
      Strsql = Strsql & " Select DISTINCT C.PERIODO, D.INICIO, D.PERIODO, D.ESTADO "
      Strsql = Strsql & "  From CSRELPATHDESCRI A "
      Strsql = Strsql & " INNER JOIN OBJETOS B ON (A.COBJ = B.CODOBJETO) "
      Strsql = Strsql & " INNER JOIN ORRSKCIRCUITO C ON (A.COBJ = C.CODOBJETO) "
      Strsql = Strsql & " INNER JOIN ORPERIODOS D ON (C.PERIODO = D.OBJID ) "
      Strsql = Strsql & " WHERE C.APLIC = 'OR' AND "
      Strsql = Strsql & " ( "
      Strsql = Strsql & " (A.MC1 = 0 AND A.MC2 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC2 = 0 AND A.MC3 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC3 = 0 AND A.MC4 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC4 = 0 AND A.MC5 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC5 = 0 AND A.MC6 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC6 = 0 AND A.MC7 = " & rs(0) & ")  "

      Strsql = Strsql & " ) "
      Strsql = Strsql & " ORDER BY D.INICIO DESC, C.PERIODO DESC "

''  response.write strsql & "<br>" & chr(13)
      set rsX5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX5.eof then

         IF (RSX5(3) & "") = "OK" THEN WSPEVAL      = "S"


             wkproccodPER = "<B>" & RSX5(2) & "&nbsp;-&nbsp;" & GTFECHADORSH(RSX5(1)) & "</B>"
             WSPDIF       = DHOY - RSX5(1)


      END IF

        WSPDISFRES = WSPDIF - WSPCANTDIAS


        IF WSPDISFRES < 20 THEN

          IF WSPDISFRES < 1 THEN
             WSPBGCL = " BGCOLOR='#00DD00' "
          ELSE
             WSPBGCL = " BGCOLOR='#F7AA22' "
          END IF


        ELSE

           WSPBGCL = " BGCOLOR='#DD0000' "


        END IF

        WSPDISFRES = "&nbsp;"

       Celda     = ""
   '''    Celda = Celda & ofv.GenCelda("","ut style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","2%","","","",sImageP,"si")
       Celda = Celda & ofv.GenCelda("","UtI style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' COLSPAN=2  ","","50%","","","",wkproccod,"si")
       Celda = Celda & ofv.GenCelda("","Ut style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ","","10%","","","",wkproccodF,"si")
       Celda = Celda & ofv.GenCelda("","Ut style='color:#000000;border-style:solid;border-width:1;border-color:#EEEEEE;' " & WSPBGCL,"","1%","","","",WSPDISFRES,"si")
       Celda = Celda & ofv.GenCelda("","UtI style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ","","30%","","","",wkproccodPER,"si")

       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


    IF  TTABMAPA > 0 THEN

      Strsql = ""


      Strsql = Strsql & " Select  A.COBJ, W.OBJID AS PERID, W.PERIODO AS PERDESCRI, W.INICIO AS PERINICIO, "
      Strsql = Strsql & " W.PERESTADO, C.PERIODO AS RSKPERIODO,  c.ESTADO AS RSKESTADO, "
      Strsql = Strsql & " W.codobjeto, W.descri, W.codvinobj, W.ESTADO AS SPROESTADO "
      Strsql = Strsql & " From CSRELPATHDESCRI A "
      Strsql = Strsql & " INNER JOIN ( "
      Strsql = Strsql & "    select z.objid, z.periodo, z.estado AS PERESTADO, Z.INICIO, x.*  "
      Strsql = Strsql & "    from orperiodos z, "
      Strsql = Strsql & "    (Select    b.codobjeto, b.descri, b.codvinobj, E.ESTADO , a.tsec "
      Strsql = Strsql & "     From CSRELPATHDESCRI A "
      Strsql = Strsql & "     INNER JOIN OBJETOS B ON (A.COBJ = B.CODOBJETO and (CODVINOBJ IN ('VN','AI','CR'))) "
      Strsql = Strsql & "     INNER JOIN OBJPRTY E ON (A.COBJ = E.CODOBJETO AND E.ESTADO <> 4) ) x "
      Strsql = Strsql & "     where z.estado = 'OK' AND Z.APLIC = 'OR' "
      Strsql = Strsql & "     ) W "
      Strsql = Strsql & "   ON (A.COBJ = W.CODOBJETO) "
      Strsql = Strsql & " LEFT JOIN ORRSKCIRCUITO C "
      Strsql = Strsql & "    ON (A.COBJ = C.CODOBJETO and W.OBJID = C.PERIODO AND c.aplic = 'OR' AND c.ESTADO <> 4 ) "

      Strsql = Strsql & " WHERE  "
      Strsql = Strsql & " ( "
      Strsql = Strsql & " (A.MC1 = 0 AND A.MC2 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC2 = 0 AND A.MC3 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC3 = 0 AND A.MC4 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC4 = 0 AND A.MC5 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC5 = 0 AND A.MC6 = " & rs(0) & ") OR "
      Strsql = Strsql & " (A.MC6 = 0 AND A.MC7 = " & rs(0) & ")  "

      Strsql = Strsql & " ) "


      Strsql = Strsql & " ORDER BY W.INICIO DESC, W.OBJID DESC "


 'response.write strsql & "<br>" & chr(13)


      set rsX5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX5.eof then

         DO WHILE NOT RSX5.EOF


            FOR II=1 TO TTABMAPA

                IF RSX5(1) = TABMAPA(II,00) THEN

                   IF (RSX5(5) & "") = "" THEN
                      TABMAPA(II,03) = TABMAPA(II,03) + 1    'PENDIENTE
                   ELSEIF (RSX5(6) & "") <> "3" THEN
                          TABMAPA(II,04) = TABMAPA(II,04) + 1    'EN CURSO
                   ELSE
                          TABMAPA(II,05) = TABMAPA(II,05) + 1    'FINALIZADO
                   END IF



                   EXIT FOR
                END IF

            NEXT



         RSX5.MOVENEXT
         LOOP

         FilaX = ""

         WSPCANTPEN = 0
         WSPCANTCUR = 0
         WSPCANTFIN = 0

                Celda     = ""
                Celda = Celda & ofv.GenCelda("","utI","","70%","","","","PERIODO","si")
                Celda = Celda & ofv.GenCelda("","UtI","","10%","","","","PENDIENTES","si")
                Celda = Celda & ofv.GenCelda("","UtI","","10%","","","","EN CURSO","si")
                Celda = Celda & ofv.GenCelda("","UtI","","10%","","","","FINALIZADOS","si")

                FilaX = filaX & ofv.GenRow("","AAA BGCOLOR='#E0E0E0' ","","","","","",Celda,"si")

         IF TTABMAPA > 1 THEN

            Celda     = ""
            Celda = Celda & ofv.GenCelda("","Ut COLSPAN=4  ","","","","","","<U>PENDIENTES PERIODOS ANTERIORES</U>","si")
            FilaX = filaX & ofv.GenRow("","","","","","","",Celda,"si")


            FOR II=2 TO TTABMAPA


                WSPCANTPEN = TABMAPA(II,03) & "&nbsp;"
                WSPCANTCUR = TABMAPA(II,04) & "&nbsp;"
                WSPCANTFIN = TABMAPA(II,05) & "&nbsp;"

                Celda     = ""
                Celda = Celda & ofv.GenCelda("","utI","","70%","","","","&nbsp;-&nbsp;" & TABMAPA(II,01),"si")
                Celda = Celda & ofv.GenCelda("","Ut","RIGHT","10%","","","",WSPCANTPEN,"si")
                Celda = Celda & ofv.GenCelda("","Ut","RIGHT","10%","","","",WSPCANTCUR,"si")
                Celda = Celda & ofv.GenCelda("","Ut","RIGHT","10%","","","",WSPCANTFIN,"si")

                FilaX = filaX & ofv.GenRow("","AAA bgcolor='#F0F0F0' ","","","","","",Celda,"si")

            NEXT



         END IF

         II = 1

         WSPCANTPEN = TABMAPA(II,03) & "&nbsp;"
         WSPCANTCUR = TABMAPA(II,04) & "&nbsp;"
         WSPCANTFIN = TABMAPA(II,05) & "&nbsp;"

            Celda     = ""
            Celda = Celda & ofv.GenCelda("","Ut COLSPAN=4  ","","","","","","<U>PERIODO ACTUAL</U>","si")
            FilaX = filaX & ofv.GenRow("","","","","","","",Celda,"si")

            Celda     = ""
            Celda = Celda & ofv.GenCelda("","utI","","70%","","","","&nbsp;-&nbsp;" & TABMAPA(II,01),"si")
            Celda = Celda & ofv.GenCelda("","Ut","RIGHT","10%","","","",WSPCANTPEN,"si")
            Celda = Celda & ofv.GenCelda("","Ut","RIGHT","10%","","","",WSPCANTCUR,"si")
            Celda = Celda & ofv.GenCelda("","Ut","RIGHT","10%","","","",WSPCANTFIN,"si")

            FilaX = filaX & ofv.GenRow("","AAA bgcolor='#F0F0F0' ","","","","","",Celda,"si")

       TablaX = ofv.GenTabla("","AAA bgcolor='#F5F5F5' style='color:#000000;border-style:solid;border-width:1;border-color:#aaaaaa;' ","","80%","","0","","0","0","","",FilaX,"si")


       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut","CENTER","2%","","","","&nbsp;","si")
       Celda = Celda & ofv.GenCelda("","Ut COLSPAN=4 ","RIGHT","","","","",TablaX,"si")

       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")



      END IF

    END IF


            FOR II=1 TO TTABMAPA

                TABMAPA(II,03) = 0
                TABMAPA(II,04) = 0
                TABMAPA(II,05) = 0

            NEXT






       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF











'''''''''''''''''''''''''''''''''



       if wkcant = 4 then
          wkcant = 1
          wkleft = wkleft + 10 + 130
          wktop  = 5
       else
          wkcant = wkcant + 1
          wktop  = wktop + 5 + 90
       end if


      rs.movenext
    loop
	'RESPONSE.Write(CHR(13) & CHR(13))
	'RESPONSE.Write(SMODAL)
	'RESPONSE.Write(CHR(13) & CHR(13))
     IF SMODAL = "LISTA" THEN

       Tabla = "<BR>" & Tabla & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")

     END IF

  end if





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = ""






      Celda = Celda & TablaH

      Celda = Celda & TablaVH
      Celda = Celda & TABLA




  sHTML = sHTML & Celda



  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no","","no","")


  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML





%>

