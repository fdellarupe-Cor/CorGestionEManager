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


if sModal = "" then sModal = "GRAF"

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
        '   shtml=shtml & "function mostrarmenuD(){" & chr(13)
        '   shtml=shtml & "document.all.Dx1.style.display='block';" & chr(13)
        '   shtml=shtml & "}" & chr(13)
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








     IF session("ORPERIODO") = "" THEN

        Strsql2 = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
        Strsql2 = Strsql2 & " ESTADO IN ('OK','CL') and RGVALEAUTOEV = 'S' "
        Strsql2 = Strsql2 & " ORDER BY XFFINAL DESC ,INICIO DESC , ESTADO DESC"
         set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rs.eof then
            session("ORPERIODO")     = RS(0)
          ''  session("ORPERIODODESC") = RS(1)
          ''  session("ORPERIODOMAP")  = RS(8)
         ELSE
            session("ORPERIODO")     = ""
          ''  session("ORPERIODODESC") = ""
          ''  session("ORPERIODOMAP")  = "0"
         END IF


     END IF


     IF (REQUEST.FORM("PERIODO") & "") <> "" THEN

         Strsql2 = "Select OBJID,PERIODO From ORPERIODOS  where OBJID = 0" & REQUEST.FORM("PERIODO")
         set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rs.eof then
            session("ORPERIODO")     = RS(0)
         '''   session("ORPERIODODESC") = RS(1)
         END IF

     END IF


     IF session("ORPERIODO") <> "" THEN

         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & session("ORPERIODO")
         set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rs.eof then
            session("ORPERIODO")     = RS(0)
            session("ORPERIODODESC") = RS(1)
            session("ORPERIODOMAP")  = RS(2)
         ELSE
            session("ORPERIODO")     = ""
            session("ORPERIODODESC") = ""
            session("ORPERIODOMAP")  = "0"
         END IF


     END IF


     WSPcodmaster = session("ORPERIODOMAP")
     WSPtipmaster = "PROCESO"




'''''''''''''''''''''''''''''''''''''''''''
 ''' session("menu") = 3
  scodmaster = request.querystring("codmaster") & ""
  stipmaster = request.querystring("tipmaster") & ""
  sCODNIVEL  = request.querystring("CODNIVEL")

  if scodmaster = "" OR stipmaster = WSPtipmaster then
    '' scodmaster = "000"
    '' stipmaster = "PROCESO"

     scodmaster = WSPcodmaster
     stipmaster = WSPtipmaster


     sCODNIVEL  = 1

     IF WKAPLIC = "EM" THEN session("ORPERIODO")     = 0



     SESSION("OROBJ") = 0

  ''   IF session("ORPERIODO") <> "" THEN
  ''       Strsql2 = "Select OBJID,PERIODO From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & session("ORPERIODO")
  ''       set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
  ''       if NOT rs.eof then
  ''          session("ORPERIODO")     = RS(0)
  ''          session("ORPERIODODESC") = RS(1)
  ''       ELSE
  ''          session("ORPERIODO")     = ""
  ''          session("ORPERIODODESC") = ""
  ''       END IF
  ''   END IF



  ''   IF session("ORPERIODO") = "" THEN
  ''      Strsql2 = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
  ''      Strsql2 = Strsql2 & " ESTADO IN ('OK','CL')  "
  ''      Strsql2 = Strsql2 & " ORDER BY XFFINAL DESC ,INICIO DESC , ESTADO DESC"
  ''       set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
  ''       if NOT rs.eof then
  ''          session("ORPERIODO")     = RS(0)
  ''         session("ORPERIODODESC") = RS(1)
  ''       ELSE
  ''          session("ORPERIODO")     = ""
  ''          session("ORPERIODODESC") = ""
  ''       END IF
  ''   END IF



  end if


   wkproccod = "emdetaproc.asp?Nombre=" & request.querystring("Nombre")
   wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & stipmaster
   wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL
   wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

   Session("ASP")     = "vuelveprocesos.asp"
   Session("ASPSTR")  = wkproccod
   Session("ASPSTRD") = Session("ASPSTR")

 IF SACCIONEJEC = "CANCEL" THEN
    XXXX = GETWKFMANAGER(SACCIONEJECCOD,SACCIONEJECTIP,4,"ORRSKCIRCUITO","WKF")
    RESPONSE.REDIRECT Session("ASPSTR")
 ELSEIF SACCIONEJEC = "VIGEN" THEN
    XXXX = GETWKFMANAGER(SACCIONEJECCOD,SACCIONEJECTIP,3,"ORRSKCIRCUITO","WKF")
    RESPONSE.REDIRECT Session("ASPSTR")
 ELSE

 END IF


''''''''''''''''''''''''''''''''''''''''''''
''' ESPECIFICACION DE PERIODOS PARA OR  ''''
''''''''''''''''''''''''''''''''''''''''''''

  wktabperiodo = ""
  Celda        = ""





   ''  IF (REQUEST.FORM("PERIODO") & "") <> "" THEN
   ''      Strsql2 = "Select OBJID,PERIODO From ORPERIODOS  where OBJID = 0" & REQUEST.FORM("PERIODO")
   ''      set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
   ''      if NOT rs.eof then
   ''         session("ORPERIODO")     = RS(0)
   ''         session("ORPERIODODESC") = RS(1)
   ''      END IF
   ''  END IF

     sForm = "formPER"

     sAccion = Session("ASPSTR")

     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL') and RGVALEAUTOEV = 'S' "
     Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"

     WKFSPPERSEN = "S"



     if session("ORPERIODO") = "" then


        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)

           if scodmaster = WSPcodmaster then
              sInput = ofv.GenerarCombo(sForm,"PERIODO",session("ORPERIODO"),Strsql,cn,0,1,"")
           else
              sinput = "&nbsp;"
           end if



        ELSE

           WKFSPPERSEN = "N"



        END IF

     ELSE

           if scodmaster = WSPcodmaster then
              sInput = ofv.GenerarCombo(sForm,"PERIODO",session("ORPERIODO"),Strsql,cn,0,1,"")
           else
              sinput = "&nbsp;"
           end if



     end if



     wkspimpreper =  session("ORPERIODODESC")
     wkspimprebkg =  "#888888"


         Strsql2 = "Select * From ORPERIODOS  where OBJID = 0" & session("ORPERIODO")
         set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rs.eof then
            XFORPERIODOFD     = RS(2)
            XFORPERIODOFH     = RS(3)
            XFORPERIODOST     = RS(4)

            IF XFORPERIODOFH     = 0 THEN XFORPERIODOFH     = 99999

            IF XFORPERIODOST     = "CL" THEN

               wkspimpreper = wkspimpreper & "&nbsp;(FINALIZADO)"
               wkspimprebkg = "#993333"

            end if


         END IF





  if wkaplic = "OR" then


     if session("ORPERIODO") = "" then



           Celda = Celda & ofv.GenCelda("","vti ","CENTER","100%","","",""," NO EXISTEN PERIODOS DE EVALUACION ACTIVOS!!! ","si")
           Fila  = ofv.GenRow("","","","","","","",Celda,"si")



     ELSE





           Celda = Celda & ofv.GenCelda("","uti style='background-color:" & wkspimprebkg & ";color:#ffffff;' ","","10%","","","","PERIODO:","si")
           Celda = Celda & ofv.GenCelda("","uti ","","65%","","","",wkspimpreper,"si")
           Celda = Celda & ofv.GenCelda("","ut  ","","20%","","","",sInput,"si")
           Fila  = ofv.GenRow("","","","","","","",Celda,"si")
           Fila =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

     end if

     wktabperiodo = ofv.GenTabla("","","","100% ","100%","0","","-1","1","","",Fila,"si")


  end if

''''''''''''''''''''''''''''''''''''''''''''
''' ESPECIFICACION DE PERIODOS PARA PL  ''''
''''''''''''''''''''''''''''''''''''''''''''



  if wkaplic = "PL" then


     if session("ORPERIODO") = "" then


           Celda = Celda & ofv.GenCelda("","vti ","CENTER","100%","","",""," NO EXISTEN PLANIFICACIONES DE CONTINUIDAD DE NEGOCIOS ACTIVOS!!! ","si")
           Fila  = ofv.GenRow("","","","","","","",Celda,"si")


     ELSE



           Celda = Celda & ofv.GenCelda("","uti style='background-color:" & wkspimprebkg & ";color:#ffffff;' ","","10%","","","","PLANIFICACION:","si")
           Celda = Celda & ofv.GenCelda("","uti ","","65%","","","",wkspimpreper,"si")
           Celda = Celda & ofv.GenCelda("","ut  ","","20%","","","",sInput,"si")
           Fila  = ofv.GenRow("","","","","","","",Celda,"si")
           Fila =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

     end if

     wktabperiodo = ofv.GenTabla("","","","100% ","100%","0","","-1","1","","",Fila,"si")


  end if



''''''''''''''''''''''''''''''''''''''''''''

 sImageCRT = "<font face='corflaticon' size=2 color='#0000ff' style='cursor:pointer;border-style:OUTSET;border-width:1;' title='Tiene circuito definido.' >&#061;</font>"

 SESSION("ORPERIODOST") = XFORPERIODOST
 SESSION("ORRSKST") = 1

 wspTIPO = "SP"
 wkproccodL = "&nbsp;"
 wkproccodL2 = "&nbsp;"
 wkproccodN = "&nbsp;"

  if wkaplic = "OR" then

       sImageL = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0143;</font>"

       wkproccodL = "<a class=nt href='ORMATRIXIMPACTO.asp?PROCESO=" & scodmaster & "&FILTRO=PROCESO&PROCMOD=ACT&FILSEL=CRT' "
       wkproccodL = wkproccodL & " TITLE='MATRIZ DE IMPACTO.' target='_blank'  >"
       wkproccodL = wkproccodL & sImageL
       wkproccodL = wkproccodL & "</a>"

       sImageL = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >&#0165;</font>"

     '  wkproccodL2 = "<a class=nt href='ORMATRIXSITUACION.asp?PROCESO=" & scodmaster & "&FILTRO=PRO&TIPO=PROCESO&PROCMOD=ACT' "
     '  wkproccodL2 = wkproccodL2 & " TITLE='SITUACION DEL PROCESO.' target='_blank'  >"
     '  wkproccodL2 = wkproccodL2 & sImageL
     '  wkproccodL2 = wkproccodL2 & "</a>"


  end if

  DIM TABSPPERM(100,3)
  IPERM = 0
  TPERM = 0



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




 if stipmaster = WSPTIPmaster then

    ' ofv.ObtenerAtributos Session("Form"),""
    ' bagregar   = ofv.Aagregar
    ' bModificar = ofv.AModificar
    ' beliminar  = ofv.Aeliminar

     AEstado = True
     AConsultar = True
     AEliminar = True
     AModificar = True
     AAgregar = True

  xxxx = ObtenerAtributos( Session("Form"),"","","")

  bAgregar = cbool(AAgregar)

  bModificar = cbool(AModificar)
  bSuprimir = cbool(AEliminar)
  bConsultar = cbool(AConsultar)

  'shtml=shtml & verdata()

''''''''''''''''
      Strsql = "Select A.* "
      Strsql = Strsql & " From RGAMBIENTES A"
      Strsql = Strsql & " where A.CODAMBIENTE = '" & scodmaster & "' and A.TIPOAMBIENTE = '" & stipmaster & "'  "
 '''  response.write strsql & "<br>" & chr(13)
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then


      IF RS("SEGCE") THEN




      END IF

      IF RS("SEGMO") THEN



      END IF


      IF RS("SEGCT") THEN


      END IF




       WKPROCDESC = RS(3) & "&nbsp;(NIVEL&nbsp;" & sCODNIVEL & ")"



  END IF

''''''''''''''''




 ''   WKPROCDESC = "PRINCIPAL DEL NEGOCIO&nbsp;(NIVEL&nbsp;" & sCODNIVEL & ") "
'' & bagregar & "-" & bModificar & "-" & bSuprimir & "-" & Session("Form")

 ''' & bagregar & "-" & bModificar & "-" & beliminar & "-" & Session("Form")


    wkproccod = "&nbsp;"
    wkproccod2 = "&nbsp;"
    wkproccodN = "&nbsp;"
    wspTIPO = "SP"

       session("WKSPTPERMCT") = "N-S"
       session("WKSPTPERMMO") = "N-S"
       session("WKSPTPERMCE") = "N-S"

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

    if bAgregar then

       wkproccodN = "<a class=nt href='javascript:void(0);' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar.' onclick=" & chr(34) & "AbreNew('" & wspTIPO & "');" & chr(34) & "  >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"

       session("WKSPTPERMCT") = "S-S"
       session("WKSPTPERMMO") = "S-S"
       session("WKSPTPERMCE") = "S-S"

    end if




 ''''' response.write session("WKSPTPERMCT") & "<br>" & chr(13)



 else

    wkproccod = "&nbsp;"
    wkproccod2 = "&nbsp;"


    WCODNIVEL  = sCODNIVEL  - 1


    WKSPTPERMCT = session("WKSPTPERMCT")
    WKSPTPERMMO = session("WKSPTPERMMO")
    WKSPTPERMCE = session("WKSPTPERMCE")


    TWKSPTPERMCT = SPLIT(WKSPTPERMCT,"-")
    TWKSPTPERMMO = SPLIT(WKSPTPERMMO,"-")
    TWKSPTPERMCE = SPLIT(WKSPTPERMCE,"-")



    FOR IPERM=1 TO UBOUND(TWKSPTPERMCT)
        TABSPPERM(IPERM,1) = TWKSPTPERMCT(IPERM)
    NEXT

    FOR IPERM=1 TO UBOUND(TWKSPTPERMMO)
        TABSPPERM(IPERM,2) = TWKSPTPERMMO(IPERM)
    NEXT

    FOR IPERM=1 TO UBOUND(TWKSPTPERMCE)
        TABSPPERM(IPERM,3) = TWKSPTPERMCE(IPERM)
    NEXT



      WKSPMASTCT = TABSPPERM(wCODNIVEL,1)
      WKSPMASTMO = TABSPPERM(wCODNIVEL,2)
      WKSPMASTCE = TABSPPERM(wCODNIVEL,3)






      Strsql = "Select A.*, C.SEGCT, C.SEGCO, C.SEGMO, C.SEGCE, c.SEGOWN, D.PERMITECT, E.PERMITEMO, F.PERMITECE From OBJETOS A"
      Strsql = Strsql & " INNER JOIN OBJSDK  C ON ( C.CODOBJETO = A.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITECT FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 1 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    D ON ( A.CODOBJETO = D.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITEMO FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 2 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    E ON ( A.CODOBJETO = E.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITECE FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 4 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    F ON ( A.CODOBJETO = E.CODOBJETO )"
      Strsql = Strsql & " where  "
      Strsql = Strsql & "  A.CODOBJETO = '" & scodmaster & "' and A.TIPOBJETO = '" & stipmaster & "'  "


 '''  response.write strsql & "<br>" & chr(13)
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then




      IF RS("SEGCE") THEN

         IF (RS("PERMITECE") & "") <> "OK" THEN

            WKSPMASTCE = "N"

         ELSE

            WKSPMASTCE = "S"

         END IF


      END IF

      IF RS("SEGMO") THEN

         IF (RS("PERMITEMO") & "") <> "OK" THEN

            WKSPMASTMO = "N"

         ELSE

            WKSPMASTMO = "S"

         END IF


      END IF


      IF RS("SEGCT") THEN

         IF (RS("PERMITECT") & "") <> "OK" THEN

            WKSPMASTCT = "N"

         ELSE

            WKSPMASTCT = "S"
            WKSPMASTCE = "S"
            WKSPMASTMO = "S"

         END IF


      END IF


      IF (RS("SEGOWN") & "") = SESSION("USUARIO") THEN

         WKSPMASTCT = "S"
         WKSPMASTMO = "S"
         WKSPMASTCO = "S"
         WKSPMASTCE = "S"

      END IF




        TABSPPERM(SCODNIVEL,1) = WKSPMASTCT
        TABSPPERM(SCODNIVEL,2) = WKSPMASTMO
        TABSPPERM(SCODNIVEL,3) = WKSPMASTCE


        WKSEP = "-"

       session("WKSPTPERMCT") = "S"
       session("WKSPTPERMMO") = "S"
       session("WKSPTPERMCE") = "S"


    FOR IPERM=1 TO SCODNIVEL

       session("WKSPTPERMCT") = session("WKSPTPERMCT") & WKSEP & TABSPPERM(IPERM,1)
       session("WKSPTPERMMO") = session("WKSPTPERMMO") & WKSEP & TABSPPERM(IPERM,2)
       session("WKSPTPERMCE") = session("WKSPTPERMCE") & WKSEP & TABSPPERM(IPERM,3)

        WKSEP = "-"

    NEXT


  ''''''response.write session("WKSPTPERMCT") & "<br>" & chr(13)


       WKPROCDESC = RS(3) & "&nbsp;(NIVEL&nbsp;" & sCODNIVEL & ")"
       WKPROCVINC = CLNG("0" & RS(15))


           XFSENPERIODO = "N"
           XFSENPERANT  = "N"
           XFSENDOCUANT = "N"
           XFSENESTADOP = 1
           XFSENESTADOD = 3

''          IF XFORPERIODOST <> "OK" OR XFSENESTADOP = 4 THEN

           Strsql = "SELECT ESTADO From OBJPRTY where codobjeto = 0" & RS(0)
           set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsX.eof then
               XFSENESTADOD = RSX(0)
           end if



       if (rs(17) & "") = "SP" OR (rs(17) & "") = "SPI" then

           SESSION("OROBJ") =  RS(0)
           wspTIPO = (rs(17) & "")

           IF WSPTIPO = "SPI" THEN

            '    wkproccodL = "<a class=nt href='ORMATRIXIMPACTO.asp?PROCESO=" & scodmaster & "&FILTRO=PROCESO&PROCMOD=OBJ' "
            '    wkproccodL = wkproccodL & " TITLE='MATRIZ DE IMPACTO.' target='_blank'  >"
            '    wkproccodL = wkproccodL & sImageL
            '    wkproccodL = wkproccodL & "</a>"

                sImageL = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0143;</font>"

            '    wkproccodL2 = "<a class=nt href='ORMATRIXSITUACION.asp?PROCESO=" & scodmaster & "&FILTRO=PROCESO&PROCMOD=OBJ' "
            '    wkproccodL2 = wkproccodL2 & " TITLE='SITUACION DEL PROCESO.' target='_blank'  >"
             '   wkproccodL2 = wkproccodL2 & sImageL
            '    wkproccodL2 = wkproccodL2 & "</a>"

           END IF

'''''''''''

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

       wkproccodN = "<a class=nt href='javascript:void(0);' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='Agregar.' onclick=" & chr(34) & "AbreNew('" & wspTIPO & "');" & chr(34) & "   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"


''''''''''''

       elseif (rs(17) & "") = "VN" then
           wspTIPO = "VN"
           wkproccodN = "&nbsp;"

                           wkproccodL = "<a class=nt href='ORMATRIXIMPACTO.asp?PROCESO=" & scodmaster & "&FILTRO=PROCESO&PROCMOD=ACT&FILSEL=CRT' "
		                   wkproccodL = wkproccodL & " TITLE='MATRIZ DE IMPACTO.' target='_blank'  >"
		                   wkproccodL = wkproccodL & sImageL
		                   wkproccodL = wkproccodL & "</a>"

		                   sImageL = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0143;</font>"


       elseif (rs(17) & "") = "AI" then
           wspTIPO = "AI"
           sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

           wkproccodN = "<a class=nt href='" & Session("ASPSTR") & "&ACTIVIDAD=ADD" & "' id='NewElem' "
           wkproccodN = wkproccodN & " TITLE='Agregar activos.'   >"
           wkproccodN = wkproccodN & sImage
           wkproccodN = wkproccodN & "</a>"

       elseif (rs(17) & "") = "CR" then
           wspTIPO = "CR"

           sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

           wkproccodN = "<a class=nt href='" & Session("ASPSTR") & "&ACTIVIDAD=ADD" & "' id='NewElem' "
           wkproccodN = wkproccodN & " TITLE='Agregar actividad.'   >"
           wkproccodN = wkproccodN & sImage
           wkproccodN = wkproccodN & "</a>"

                wkproccodL = "<a class=nt href='ORMATRIXIMPACTO.asp?PROCESO=" & scodmaster & "&FILTRO=PROCESO&PROCMOD=ACT&FILSEL=CRT' "
                wkproccodL = wkproccodL & " TITLE='MATRIZ DE IMPACTO.' target='_blank'  >"
                wkproccodL = wkproccodL & sImageL
                wkproccodL = wkproccodL & "</a>"

                sImageL = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0143;</font>"



       else
           SESSION("OROBJ") =  RS(0)
           wspTIPO = "SP"
           wkproccodN = "&nbsp;"
       end if





       if (rs(17) & "") = "CR" or (rs(17) & "") = "AI" then

          IF XFSENESTADOD = 4 THEN wkproccodN = "&nbsp;"

       else
           if WKSPMASTCT <> "S" then wkproccodN = "&nbsp;"
       end if




    end if



   if ((rs(17) & "") = "VN" OR (rs(17) & "") = "CR" OR (rs(17) & "") = "AI")   then

         if wkaplic = "OR" then

           WKSPTABCIR = "CIRCUITO NO INICIADO PARA EL PERIODO."

         ELSEif wkaplic = "PL" then

           WKSPTABCIR = "CIRCUITO NO INICIADO PARA ESTA PLANIFICACION."

         else

           WKSPTABCIR = "OPERACION NO PREVISTA."

         end if

           WKSPTIPOPR = "SI"

           WKSPCIRCCT = WKSPMASTCT
           WKSPCIRCMO = WKSPMASTMO
           WKSPCIRCCO = WKSPMASTCO
           WKSPCIRCCE = WKSPMASTCE

 '''  response.write WKSPMASTCT & WKSPMASTMO  & WKSPMASTcO & WKSPMASTce & "<br>" & chr(13)


           Strsql = "SELECT A.*, D.PERMITECT, E.PERMITEMO, F.PERMITECE From ORRSKCIRCUITO A "

           Strsql = Strsql & " LEFT  JOIN "
           Strsql = Strsql & "       (SELECT DISTINCT codRSK, 'OK' AS PERMITECT FROM ORRSKCIRCUITOSDKDET "
           Strsql = Strsql & "        WHERE TIPOSEG = 1 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
           Strsql = Strsql & "        ) "
           Strsql = Strsql & "                    D ON ( A.OBJID = D.codRSK )"
           Strsql = Strsql & " LEFT  JOIN "
           Strsql = Strsql & "       (SELECT DISTINCT codRSK, 'OK' AS PERMITEMO FROM ORRSKCIRCUITOSDKDET "
           Strsql = Strsql & "        WHERE TIPOSEG = 2 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
           Strsql = Strsql & "        ) "
           Strsql = Strsql & "                    E ON ( A.OBJID = E.codRSK )"
           Strsql = Strsql & " LEFT  JOIN "
           Strsql = Strsql & "       (SELECT DISTINCT codRSK, 'OK' AS PERMITECE FROM ORRSKCIRCUITOSDKDET "
           Strsql = Strsql & "        WHERE TIPOSEG = 4 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
           Strsql = Strsql & "        ) "
           Strsql = Strsql & "                    F ON ( A.OBJID = E.codRSK )"

           strsql = strsql & " where A.codobjeto = 0" & RS(0) & " AND "
           strsql = strsql & " A.APLIC = '" & wkaplic & "' AND A.PERIODO = 0" & session("ORPERIODO")
   'response.write strsql & "<br>" & chr(13)
           set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsX.eof then

               WKSPCIRCCOD = rsX(0)
               'response.write "circuito "& WKSPCIRCCOD

               IF rsX("SEGCE") THEN

                  IF (rsX("PERMITECE") & "") <> "OK" THEN

                     WKSPCIRCCE = "N"

                  ELSE

                     WKSPCIRCCE = "S"

                  END IF


               END IF

               IF rsX("SEGMO") THEN

                  IF (rsX("PERMITEMO") & "") <> "OK" THEN

                     WKSPCIRCMO = "N"

                  ELSE

                     WKSPCIRCMO = "S"

                  END IF


               END IF


               IF rsX("SEGCT") THEN

                  IF (rsX("PERMITECT") & "") <> "OK" THEN

                     WKSPCIRCCT = "N"

                  ELSE

                     WKSPCIRCCT = "S"
                     WKSPCIRCCE = "S"
                     WKSPCIRCMO = "S"

                  END IF


               END IF


               IF (rsX("OWNER") & "") = SESSION("USUARIO") THEN

                     WKSPCIRCCT = "S"
                     WKSPCIRCCE = "S"
                     WKSPCIRCMO = "S"
                     WKSPCIRCCO = "S"
               END IF








               XFSENPERIODO = "S"
               XFSENESTADOP = RSX(8)
               SESSION("ORRSKST") = XFSENESTADOP





               if (rsx(12)  or rsx(13)  or rsx(14)  or rsx(15) ) then
                  sImagePER = "<font face='corflaticon' size=3 color='#f7aa22'>&#0207;</font>"
               ELSE
                  sImagePER = "<font face='corflaticon' size=3 color='#00aa00'>&#0208;</font>"
               END IF

               IF WKSPCIRCCT = "S" THEN

                  if XFSENESTADOP = 4 then
                     WKSPPERMISOS = "&nbsp;" & sImagePER & "&nbsp;"

                  else
                     WKSPPERMISOS = "<a style='color:#990000;cursor:pointer;' "
                     WKSPPERMISOS = WKSPPERMISOS & " href='edobjsdk.asp?Tipobjeto=" & rsx(2) & "&Codobjeto="
                     WKSPPERMISOS = WKSPPERMISOS & rsx(1) & "&MODALIDAD=PRGCOR' title='Permisos del circuito.' >"
                     WKSPPERMISOS = WKSPPERMISOS & "&nbsp;" & sImagePER & "&nbsp;PERMISOS&nbsp;</a>"
                  end if

               ELSE

                  WKSPPERMISOS = "&nbsp;" & sImagePER & "&nbsp;RESTRINGIDO&nbsp;"

               END IF


              Celda = ""

             ' Celda = Celda & ofv.GenCelda("","Vti style='Xbackground-color:#993333;Xcolor:#ffffff;color:#990000;' ","","20%","","","","CIRCUITO DE AUTOEVALUACION","si")
              Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","15%","","","","ESTADO:","si")
              Celda = Celda & ofv.GenCelda("","utI  ","","20%","","","",GTESTADODSC(RSX(8)),"si")
              Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","15%","","","","CREADO:","si")
              Celda = Celda & ofv.GenCelda("","utI  ","","20%","","","",GTFECHADORSH(RSX(5)),"si")
              Celda = Celda & ofv.GenCelda("","Uti style='background-color:#888888;color:#ffffff;' ","","15%","","","","CONFIRMADO:","si")
              Celda = Celda & ofv.GenCelda("","utI  ","","20%","","","",GTFECHADORSH(RSX(6)),"si")

              Fila  = ofv.GenRow("","","","","","","",Celda,"si")
              TablaW = ofv.GenTabla("","AAA style='background-color:#ddddff;border-style:SOLID;border-width:1;Border-color:#000000;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

              WKSPTABCIR = TablaW


           else

               WKSPPERMISOS = "&nbsp;"

           end if


               IF XFSENESTADOD = 4 THEN
                     WKSPCIRCCT = "N"
                     WKSPCIRCCE = "N"
                     WKSPCIRCMO = "N"
                  '   XFSENESTADOP = 4
               END IF


           if WKSPCIRCMO <> "S" then wkproccodN = "&nbsp;"



       IF  XFORPERIODOST = "OK"  AND XFSENESTADOD <> 4 THEN

         IF  XFSENPERIODO = "N"  THEN



           Strsql = "SELECT A.OBJID,A.PERIODO,B.PERIODO From ORRSKCIRCUITO A, ORPERIODOS B where A.codobjeto = 0" & RS(0) & " AND "
           strsql = strsql & " A.APLIC = '" & wkaplic & "' AND A.PERIODO = B.OBJID "
        '   strsql = strsql & " AND A.PERIODO <> 0" & session("ORPERIODO")

           set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsX2.eof then

               XFSENPERANT = "S"


           else


           end if

        END IF




         IF ((rs(17) & "") = "VN" OR (rs(17) & "") = "CR" OR (rs(17) & "") = "AI") AND XFSENPERIODO = "N" AND XFSENPERANT  = "N" THEN


           strsql = "SELECT B.FECBAJA, B.ESTADO, A.OBJREM FROM OBJETOS A INNER JOIN OBJPRTY B   "
           strsql = strsql & " ON A.CODOBJETO = B.CODOBJETO "
           strsql = strsql & " WHERE A.CODOBJETO = 0" & RS(0)
       '''   response.write strsql & "<br>"
           set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if NOT rsN.eof then
              IF RSN(1) = 3 THEN

                 IF RSN(2) > 0 THEN

                    Strsql = "SELECT A.periodo, A.CODOBJETO, A.TIPOBJETO, B.DESCRI, C.PERIODO AS PERIDSC From ORRSKCIRCUITO A "
                    strsql = strsql & " INNER JOIN OBJETOS B ON (A.APLIC = '" & wkaplic & "' AND A.CODOBJETO = B.CODOBJETO) "
                    strsql = strsql & " INNER JOIN ORPERIODOS C ON (A.PERIODO = C.OBJID) "
                    strsql = strsql & " where A.codobjeto = 0" & RSN(2) & "  "
                    strsql = strsql & "  order by a.objid desc "
                    set rsX32 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    if  not rsX32.eof then
                        XFSENDOCUANT     = "S"
                        XFSENDOCUANTP    = rsX32(0)
                        XFSENDOCUANTPDSC = rsX32(4)
                        XFSENDOCUANTCO   = rsX32(1)
                        XFSENDOCUANTTO   = rsX32(2)
                    end if

                 END IF

              END IF
           END IF




        END IF

       END IF

       XFDX6DSP = "none"


       if  (XFSENPERIODO = "N" and XFORPERIODOST = "OK" ) OR SACCIONEJEC = "ABREESTAD" then

           IF SESSION("OROBJ") <> RS(0) THEN
              SESSION("OROBJ") =  RS(0)
          '''''    if WKSPCIRCCT = "S" then XFDX6DSP = "BLOCK"
              if SACCIONEJEC = "ABREESTAD" then XFDX6DSP = "BLOCK"
           END IF
       end if


          sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0120;</font>"

          wkproccod3 = "<a class=nt onclick=""document.all.DX6.style.display='none';"" href='javascript:void(0);' "
          wkproccod3 = wkproccod3 & " ><b>"
          wkproccod3 = wkproccod3 & sImage
          wkproccod3 = wkproccod3 & "</b></a>"


          Celda = ""
          Celda = Celda & ofv.GenCelda("","uti STYLE='COLOR:#990000;' ","","98%","","","","ESTADO DEL CIRCUITO","si")
          Celda = Celda & ofv.GenCelda("","ut ","","2%","","","",wkproccod3,"si")
          FilaN = ofv.GenRow("","","","","","","",Celda,"si")
          FilaW1  = ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' ","","100%","","0","","0","1","","",FilaN,"si")



          if  XFSENPERIODO = "S" then

                    XFSENWKFAUTOSET = 0
                    XFSENWKFOBSSET  = 0



                    Strsql = "SELECT * From ORRSKCONFIG where APLIC = '" & wkaplic & "' "
                    'response.write strsql
                    set rsX3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    if  not rsX3.eof then
                        XFSENWKFAUTOSET = rsX3(3)
                        XFSENWKFOBSSET  = rsX3(4)
                    end if



              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti style='background-color:#336699;color:#ffffff;' ","","80%","","","","CIRCUITO ACTUAL","si")

          session("TCliente") = "RSK"
          session("NCliente") = RSX(0)


          if XFORPERIODOST <> "OK" or RSX(8) = 4 THEN

          ELSE

'''' CANCELA - CIRCUITO DE AUTOEVALUACION

              sImage = "<font face='corflaticon' size=3 color='#993333' style='border-style:OUTSET;border-width:2;' >&#0114;</font>"

              SACCIONF = Session("ASPSTR") & "&ACCIONEJEC=CANCEL&ACCIONEJECCOD=" & RSX(0) & "&ACCIONEJECTIP=RSK"


              sformobj = "rskcancel"
              sAccion     = SACCIONF
              formobj1    = ofv.genform(sformobj,"",sAccion,"","post"," style='display:none;' ",sInput,"si")

              wkprocOPT1 = "<B>" & sImage & "</B>" & formobj1

              Celda = Celda & ofv.GenCelda("","ut title='Cancelar circuito.' style='border-style:OUTSET;border-width:1;color:#AA0000;' onclick=""disp_confirm('Confirma la cancelacion del circuito???','rskcancel');""  ","","2%","","","",wkprocOPT1,"si")

          END IF

       '  response.write  XFORPERIODOST & "<br>" & XFSENWKFAUTOSET & "<br>" & XFSENWKFOBSSET  & "<br>" &  RSX(8) & "<br>" & XFSENWKFOBSSET & "<br>" & RSX(8)& "<br>" & RSX(8)


             if XFORPERIODOST <> "OK" or XFSENWKFAUTOSET = 2 OR (XFSENWKFOBSSET  = 2 AND RSX(8) <> 6) OR (XFSENWKFOBSSET  = 1 AND RSX(8) <> 1 AND RSX(8) <> 6) THEN

          ' if XFORPERIODOST <> "OK" or XFSENWKFAUTOSET  = 0 OR (XFSENWKFOBSSET  = 0 AND RSX(8) <> 6) OR (XFSENWKFOBSSET  = 1 AND RSX(8) <> 1 AND RSX(8) <> 6) THEN



          else

'''' CONFIRMA - CIRCUITO DE AUTOEVALUACION

              sImage = "<font face='corflaticon' size=3 color='#009900' style='border-style:OUTSET;border-width:2;' >&#097;</font>"

              SACCIONF = Session("ASPSTR") & "&ACCIONEJEC=VIGEN&ACCIONEJECCOD=" & RSX(0) & "&ACCIONEJECTIP=RSK"
              sformobj = "rskvigen"
              sAccion     = SACCIONF
              formobj1    = ofv.genform(sformobj,"",sAccion,"","post"," style='display:none;' ",sInput,"si")

              wkprocOPT1 = "<B>" & sImage & "</B>" & formobj1

              Celda = Celda & ofv.GenCelda("","ut title='Confirmar circuito.' style='border-style:OUTSET;border-width:1;color:#AA0000;' onclick=""disp_confirm('Confirma la puesta en vigencia del circuito???','rskvigen');""  ","","2%","","","",wkprocOPT1,"si")

          END IF


          if XFORPERIODOST <> "OK" or XFSENWKFAUTOSET = 0 OR (XFSENWKFOBSSET  = 2 AND RSX(8) <> 6)  OR (XFSENWKFOBSSET  = 1 AND RSX(8) <> 1 AND RSX(8) <> 6) THEN

          ELSE


'''' GESTIONA CIRCUITO DE AUTORIZACION - CIRCUITO DE AUTOEVALUACION

              sImage   = "<FONT FACE=corflaticon COLOR='#0000ff' SIZE=3 style='border-style:OUTSET;border-width:2;'  >&#0130;</FONT>"



              SACCIONF = "edobjwflor.asp?PASO=AUTO&Tipobjeto=RSK&Codobjeto=" & RSX(0) & "&GLBSOLN=ESPWFL"

              sformobj = "rskautor"
              sAccion     = SACCIONF
              formobj1    = ofv.genform(sformobj,"",sAccion,"","post"," style='display:none;' ",sInput,"si")

              wkprocOPT1 = "<B>" & sImage & "</B>" & formobj1

              Celda = Celda & ofv.GenCelda("","ut title='Configuracion del workflow.' style='border-style:OUTSET;border-width:1;color:#AA0000;' onclick=""disp_confirm('Confirma ingresar a la configuracion del workflow???','rskautor');""  ","","2%","","","",wkprocOPT1,"si")


          END IF


          if XFORPERIODOST <> "OK" or (XFSENWKFOBSSET  = 0 )  OR (XFSENWKFOBSSET  = 2 AND RSX(8) <> 6) OR (XFSENWKFOBSSET  = 1 AND RSX(8) <> 1 AND RSX(8) <> 6) THEN

          ELSE

'''' GESTIONA CIRCUITO DE OBSERVACION - CIRCUITO DE AUTOEVALUACION

              sImage   = "<FONT FACE=corflaticon COLOR='#ff0000' SIZE=3 style='border-style:OUTSET;border-width:2;'  >&#094;</FONT>"


              SACCIONF = "edobjwflor.asp?PASO=OBS&Tipobjeto=RSK&Codobjeto=" & RSX(0) & "&GLBSOLN=ESPWFL"

              sformobj = "rskobser"
              sAccion     = SACCIONF
              formobj1    = ofv.genform(sformobj,"",sAccion,"","post"," style='display:none;' ",sInput,"si")

              wkprocOPT1 = "<B>" & sImage & "</B>" & formobj1

              Celda = Celda & ofv.GenCelda("","ut title='Circuito de comentarios.' style='border-style:OUTSET;border-width:1;color:#AA0000;' onclick=""disp_confirm('Confirma el inicio del circuito de comentarios???','rskobser');""  ","","2%","","","",wkprocOPT1,"si")

          END IF


          '''    sImage  = "<FONT FACE=corflaticon COLOR='#f7aa22' SIZE=3 style='border-style:OUTSET;border-width:2;'  >&#095;</FONT>"
              sImage  = "<FONT FACE=corflaticon COLOR='#993333' SIZE=3   >&#095;</FONT>"

              SACCIONF = "coobshist.asp?Tipobjeto=RSK&Codobjeto=" & RSX(0)

              sformobj = "rskhobs"
              sAccion     = SACCIONF
              formobj1    = ofv.genform(sformobj,"",sAccion,"_blank","post"," style='display:none;' ",sInput,"si")

              wkprocOPT1 = "<B>" & sImage & "</B>" & formobj1

            ''  Celda = Celda & ofv.GenCelda("","ut title='Historial de comentarios.'  style='border-style:OUTSET;border-width:1;color:#AA0000;' onclick=""disp_confirm('Desea acceder al historial de comentarios???','rskhobs');""  ","","2%","","","",wkprocOPT1,"si")
              CeldaHO = ofv.GenCelda("","ttr title='Historial de comentarios.' style='cursor:pointer;'  onclick=""disp_confirm('Desea acceder al historial de comentarios???','rskhobs');""  ","","2%","","","",wkprocOPT1,"si")

        '''     sImage  = "<FONT FACE=corflaticon COLOR='#336699' SIZE=3 style='border-style:OUTSET;border-width:2;'  >2</FONT>"
              sImage  = "<FONT FACE=corflaticon COLOR='#336699' SIZE=3   >2</FONT>"

              SACCIONF = "CoHistWkf.asp?DSD=RSK"

              sformobj = "rskhwkf"
              sAccion     = SACCIONF
              formobj1    = ofv.genform(sformobj,"",sAccion,"_blank","post"," style='display:none;' ",sInput,"si")

              wkprocOPT1 = "<B>" & sImage & "</B>" & formobj1

          '''    Celda = Celda & ofv.GenCelda("","ut title='Historial de circuitos de autorizacion y observaciones.' style='border-style:OUTSET;border-width:1;color:#AA0000;' onclick=""disp_confirm('Desea acceder al historial de circuitos de autorizacion y observaciones???','rskhwkf');""  ","","2%","","","",wkprocOPT1,"si")
              CeldaHC = ofv.GenCelda("","ttr title='Historial de circuitos de autorizacion y comentarios.' style='cursor:pointer;'  onclick=""disp_confirm('Desea acceder al historial de circuitos de autorizacion y comentarios???','rskhwkf');""  ","","2%","","","",wkprocOPT1,"si")


              FilaN = ofv.GenRow("","","","","25","","",Celda,"si")
              FilaW1  = FilaW1 & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' ","","100%","","0","","1","1","","",FilaN,"si")


              FilaN = ""

              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","ESTADO","si")
              SINPUT = GTESTADODSC(RSX(8))
              Celda = Celda & ofv.GenCelda("","uti  ","","","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")
              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","RESPONSABLE","si")

              xxxx = getInfoUser(RSX(4) & "")
              SINPUT = session("SGU_NOMBRE")
              Celda = Celda & ofv.GenCelda("","uti  ","","","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")
              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","CREACION","si")
              SINPUT = GTFECHADOR(RSX(5))
              Celda = Celda & ofv.GenCelda("","uti  ","","","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")
              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","VIGENCIA","si")
              SINPUT = GTFECHADOR(RSX(6))
              Celda = Celda & ofv.GenCelda("","uti  ","","","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")
              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","","","","","BAJA","si")
              SINPUT = GTFECHADOR(RSX(7))
              Celda = Celda & ofv.GenCelda("","uti  ","","","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")
              FilaW2  = ofv.GenTabla("","","","100%","","0","","0","1","","",FilaN,"si") & "<BR>"


          end if



'''XFSENPERIODO = "N"

          if  XFSENPERIODO = "N"   AND XFSENESTADOD <> 4 and XFORPERIODOST = "OK" then


              Celda = ""
              Celda = Celda & ofv.GenCelda("","uti COLSPAN=2 style='background-color:#336699;color:#ffffff;' ","","","","","","INICIACION DEL CIRCUITO","si")
              FilaN = ofv.GenRow("","","","","25","","",Celda,"si")
              Celda = ""
              Celda = Celda & ofv.GenCelda("","ut style='background-color:#990000;color:#ffffff;' ","CENTER","4%","","","","*","si")


                   sID = "RGREEMPLAZASP.asp?"
               IF WKAPLIC = "OR" THEN
                   sID = sID & "OPCION=TNP"
               ELSEIF WKAPLIC = "PL" THEN
                   sID = sID & "OPCION=TNPPL"
               ELSE
                   sID = sID & "OPCION=TNP"
               END IF

                   sID = sID & "&CodESQS=" & RS(0)
                   sID = sID & "&TipESQS=" & RS(2)
                   sID = sID & "&volver=" & REPLACE(Session("ASPSTR") & "","&","_")



                   sinput2 = "<a class=NT style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "none" & chr(34) & ";' "
                   sinput2 = sinput2  & "  href='" & sID & "' onclick=""document.all.DX6.style.display='none';document.all.DX4.style.display='block';"" "
                   sinput2 = sinput2  & "  TITLE='INICIAR CIRCUITO NUEVO PARA ESTE PERIODO.'  target='fesquemas' >INICIAR CIRCUITO NUEVO PARA ESTE PERIODO.&nbsp;</a>"


              SINPUT = sinput2
              Celda = Celda & ofv.GenCelda("","uti  ","","96%","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")

           if  XFSENPERANT = "S" then

             DO WHILE NOT RSX2.EOF

              Celda = ""
              Celda = Celda & ofv.GenCelda("","ut style='background-color:#009900;color:#ffffff;' ","CENTER","4%","","","","*","si")
              SINPUT = "TOMAR CIRCUITO DEL PERIODO " & RSX2(2) & "."


                   sID = "RGREEMPLAZASP.asp?"
                   sID = sID & "OPCION=TNPDA"
                   sID = sID & "&CodESQS=" & RS(0)
                   sID = sID & "&TipESQS=" & RS(2)
                   sID = sID & "&periodo=" & RSX2(1)
                   sID = sID & "&volver=" & REPLACE(Session("ASPSTR") & "","&","_")



                   sinput2 = "<a class=NT style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "none" & chr(34) & ";' "
                   sinput2 = sinput2  & "  href='" & sID & "' onclick=""document.all.DX6.style.display='none';document.all.DX4.style.display='block';"" "
                   sinput2 = sinput2  & "  TITLE='TOMAR CIRCUITO DEL PERIODO " & RSX2(2) & ".'  target='fesquemas' >TOMAR CIRCUITO DEL PERIODO " & RSX2(2) & ".&nbsp;</a>"




              SINPUT = sinput2
              Celda = Celda & ofv.GenCelda("","uti  ","","96%","","","",SINPUT,"si")




              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")

             RSX2.MOVENEXT
             LOOP


           END IF


           if  XFSENDOCUANT = "S" then

               '''         XFSENDOCUANTP = rsX3(0)
               '''         XFSENDOCUANTCO = rsX3(1)
               '''         XFSENDOCUANTTO = rsX3(2)

              ''     if  not rsX32.eof then
              ''          XFSENDOCUANT     = "S"
              ''          XFSENDOCUANTP    = rsX32(0)
              ''          XFSENDOCUANTPDSC = rsX32(4)
              ''          XFSENDOCUANTCO   = rsX32(1)
              ''          XFSENDOCUANTTO   = rsX32(2)
              ''      end if

             DO WHILE NOT rsX32.EOF

                        XFSENDOCUANTP    = rsX32(0)
                        XFSENDOCUANTPDSC = rsX32(4)
                        XFSENDOCUANTCO   = rsX32(1)
                        XFSENDOCUANTTO   = rsX32(2)


              Celda = ""
              Celda = Celda & ofv.GenCelda("","ut style='background-color:#000099;color:#ffffff;' ","CENTER","4%","","","","*","si")
              SINPUT = "TOMAR CIRCUITO DEL PROCESO ORIGINAL " & rsX32(3) & " / PERIODO: " & XFSENDOCUANTPDSC & "."


                   sID = "RGREEMPLAZASP.asp?"
                   sID = sID & "OPCION=TNPDAV"
                   sID = sID & "&CodESQS="  & RS(0)
                   sID = sID & "&TipESQS="  & RS(2)
                   sID = sID & "&CodESQSo=" & XFSENDOCUANTCO
                   sID = sID & "&TipESQSo=" & XFSENDOCUANTTO
                   sID = sID & "&periodo="  & XFSENDOCUANTP
                   sID = sID & "&volver="   & REPLACE(Session("ASPSTR") & "","&","_")



                   sinput2 = "<a class=NT style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "none" & chr(34) & ";' "
                   sinput2 = sinput2  & "  href='" & sID & "' onclick=""document.all.DX6.style.display='none';document.all.DX4.style.display='block';"" "
                   sinput2 = sinput2  & "  TITLE='TOMAR CIRCUITO DEL PROCESO ORIGINAL .'  target='fesquemas' >TOMAR CIRCUITO DEL PROCESO ORIGINAL:<BR>-----<B>" & rsX32(3) & " / PERIODO: " & XFSENDOCUANTPDSC & "</B>.&nbsp;</a>"

'''''''''''''''''''target='fesquemas'


              SINPUT = sinput2
              Celda = Celda & ofv.GenCelda("","uti  ","","96%","","","",SINPUT,"si")
              FilaN = FilaN & ofv.GenRow("","","","","","","",Celda,"si")

            rsX32.MOVENEXT
            LOOP




           END IF



              FilaW3  = ofv.GenTabla("","","","100%","","0","","0","1","","",FilaN,"si")


          end if





          FilaW  = "<CENTER>" & FilaW1  & FilaW2   & FilaW3 & "</CENTER>"



          Celda = "" & Chr(13)

          Celda = Celda & "<div id='DX6' style='z-index:1;margin-top:1;margin-left:1;position:absolute;left:400px;top:140px;display:" & XFDX6DSP & ";"
          Celda = Celda & "width:600;height:350;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
          Celda = Celda & "background-color:#F9F9F9;"
          Celda = Celda & "' >" & Chr(13)



          Celda = Celda & FilaW & Chr(13)


          Celda = Celda & "</div>" & Chr(13)


          sHTML = sHTML & Celda & Chr(13)

          Celda = ""
          Filan  = ""
          FilaW1  = ""
          FilaW2  = ""
          FilaW3  = ""
          FilaW  = ""

          wkproccod3 = "<a class=nt style='cursor:pointer;COLOR:#990000;' "
          wkproccod3 = wkproccod3 & " title='Ver estado.' onclick=""document.all.DX6.style.display='BLOCK';"" href='javascript:void(0);' "
          wkproccod3 = wkproccod3 & " ><B>"

          IF XFSENPERIODO = "S" THEN
             wkproccod3 = wkproccod3 & "&nbsp;VER&nbsp;ESTADO&nbsp;"
          ELSE
             if XFORPERIODOST = "OK" then
                wkproccod3 = wkproccod3 & "&nbsp;INICIAR&nbsp;"
             else
                wkproccod3 = wkproccod3 & "&nbsp;NO&nbsp;INICIADO&nbsp;"
             end if
          END IF

          wkproccod3 = wkproccod3 & "</B></a>"


       wkprocSPW = wkproccod3


          IF WKSPCIRCCT <> "S" THEN

             wkprocSPW = "&nbsp;RESTRINGIDO&nbsp;"

          END IF






    END IF

    Strsql = "Select codobjmaster,tipobjmaster From relaobj  where CODojasoc = " & scodmaster
    Strsql = Strsql & " and tipobjasoc = '" & stipmaster & "'  "

 '' response.write strsql & "<br>"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then

       sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0233;</font>"

       wkproccod = "<a class=nt onclick=""document.all.DX4.style.display='none';"" href='emdetaproc.asp?Nombre=" & request.querystring("Nombre")
       wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & rs(0) & "&tipmaster=" & rs(1)
       wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO
       wkproccod = wkproccod & "&CODNIVEL=" & (sCODNIVEL - 1) & "' title='SUBIR DE NIVEL.' >"
       wkproccod = wkproccod & sImage
       wkproccod = wkproccod & "</a>"

    end if


    WKDIVSECUEN = GENPROCSECUENEX()






 end if

''WKSPPERMISOS



           select case wspTIPO


                  case "SP"

                        sInput = "PROCESO"
                        wspTIPOCLC = "#993333"

                        IF sCODNIVEL > 1 THEN SINPUT = "SUB" & SINPUT

                  case "SPI"

                        sInput = "PROCESO INFORMATICO"
                        wspTIPOCLC = "#8899BB"

                        IF sCODNIVEL > 1 THEN SINPUT = "SUB" & SINPUT

                  case "VN"

                        sInput = "EVENTO VINCULADO EMANAGER"
                        wspTIPOCLC = "#008800"

                  case "CR"

                        sInput = "EVENTO MANUAL"
                        wspTIPOCLC = "#99CCDD"

                  case "AI"

                        sInput = "ACTIVO INFORMATICO"
                        wspTIPOCLC = "#0000CC"

                  case else

                        sInput = "SIN DETERMINAR"
                        wspTIPOCLC = "#000000"

           end select


    sImagePRC = "<font face='corflaticon' size=2 color='" & wspTIPOCLC & "'  >&#060;</font>"

    WKPROCDESC = WKDIVSECUEN & sInput & ":&nbsp;" & WKPROCDESC


    Celda = ""

   Celda = Celda & ofv.GenCelda("","Ut style='background-color:#f0f0f0;color:#990000;' ","CENTER","1%","","","",sImagePRC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","",WKPROCDESC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='border-style:none;border-width:1;border-color:#aaaaaa;' ","CENTER","1%","","","",wkproccodL2,"si")
   Celda = Celda & ofv.GenCelda("","ut style='border-style:none;border-width:1;border-color:#aaaaaa;' ","CENTER","1%","","","",wkproccodL,"si")
   Celda = Celda & ofv.GenCelda("","ut style='border-style:none;border-width:1;border-color:#aaaaaa;' ","CENTER","1%","","","",wkproccodN,"si")
  ' Celda = Celda & ofv.GenCelda("","ut style='border-style:none;border-width:1;border-color:#aaaaaa;' ","CENTER","1%","","","",wkproccod2,"si")
   Celda = Celda & ofv.GenCelda("","ut style='border-style:none;border-width:1;border-color:#aaaaaa;' ","CENTER","1%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = ofv.GenTabla("","aaa style='background-color:#f0eedd;border-style:solid;border-width:1;border-color:#000000;' valign='middle' align='center' ","","100% ","100%","0","","-1","1","","",Fila,"si")

''''''''''''''''''''''''''''''''''''''''''''
   TablaVH = ""

   if WKSPTIPOPR = "SI" then


       Celdah = ofv.GenCelda("","VT style='background-color:#EEEEEE;color:#990000;' ","center","100%","","","",WKSPTABCIR,"si")
       CeldaH = CeldaH & ofv.GenCelda("","TTR  ","CENTER","10%","","","",WKSPPERMISOS,"si")
       CeldaH = CeldaH & ofv.GenCelda("","TTR  ","CENTER","10%","","","",wkprocSPW,"si") & celdaHO & celdaHC
       Filah  = ofv.GenRow("","","","","25","","",Celdah,"si")
       TablaVH = ofv.GenTabla("","aaa style='background-color:#888888;border-style:SOLID;border-width:2;Border-color:#888888;'  ","","100% ","","0","","1","-1","","",Filah,"si")



   end if

''''''''''''''''''''''''''''''''''''''''''''



   Fila = ""

          wspcolor1  = "#990000"
          wspcolor2  = "#009900"
          sImageO1   = "&#0253;"
          sImageO2   = "&#0254;"


   IF wspTIPO = "SP" OR wspTIPO = "SPI" THEN




      Strsql = "Select A.*, C.SEGCT, C.SEGCO, C.SEGMO, C.SEGCE, c.SEGOWN, D.PERMITE, E.PERMITEX, CRT.OBJID AS CRTY, CRTO.CODOBJETO AS CRTYO From OBJETOS A"
      Strsql = Strsql & " INNER JOIN relaobj B ON ( B.CODOJASOC = A.CODOBJETO )"
      Strsql = Strsql & " INNER JOIN OBJSDK  C ON ( C.CODOBJETO = A.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITE FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE  ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    D ON ( A.CODOBJETO = D.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITEX FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 1 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    E ON ( A.CODOBJETO = E.CODOBJETO )"

      Strsql = Strsql & " LEFT JOIN ORRSKCIRCUITO CRT ON ( A.CODOBJETO = CRT.CODOBJETO AND PERIODO = " & SESSION("ORPERIODO") & " )"

      Strsql = Strsql & " LEFT JOIN (SELECT DISTINCT CODOBJETO FROM ORRSKCIRCUITO) CRTO "
      Strsql = Strsql & "    ON ( A.CODOBJETO = CRTO.CODOBJETO )"

      Strsql = Strsql & " where  "
      Strsql = Strsql & "  B.codobjmaster = '" & scodmaster & "' and B.tipobjmaster = '" & stipmaster & "' AND "
      Strsql = Strsql & "  ( (C.SEGCO = 0) OR  "
      Strsql = Strsql & "    ( (C.SEGCO = 1) AND "
      Strsql = Strsql & "      ( (SEGOWN = '" & Session("Usuario") & "' ) OR ( D.CODOBJETO <> '') "
      Strsql = Strsql & "      )  "
      Strsql = Strsql & "    )  "
      Strsql = Strsql & "  )  "
      Strsql = Strsql & " ORDER BY a.prioridad , A.INDOBJ "


  '' response.write strsql & "<br>" & chr(13)


   ELSE

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' MODIFICA LA CRITICIDAD DE UNA ACTIVIDAD
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    IF SACTIVIDAD <> "" THEN


      IF SACTIVIDAD = "SEC" THEN

         WKNEWSEC = "0" & REQUEST.FORM("NEWSEC")
         WKCROBJ  = "0" & REQUEST.QUERYSTRING("CROBJ")

         strsql = "UPDATE OBJSECTOR SET CODSECTOR = '" & WKNEWSEC & "' "
         strsql = strsql & " WHERE CODOBJETO = '" & WKCROBJ & "' "
      '''   RESPONSE.WRITE STRSQL & "<BR>"
         call ofv.crearconsultaEx(StrSql,cn,1,volver)


      ELSEIF SACTIVIDAD = "FNC" THEN

         WKNEWFNC = "" & REQUEST.FORM("NEWFNC")
         WKCROBJ  = "" & REQUEST.QUERYSTRING("CROBJ")

         strsql = "UPDATE OBJPROCED SET CODPERFIL = '" & WKNEWFNC & "' "
         strsql = strsql & " WHERE CODOBJETO = '" & WKCROBJ & "' "
       '''  RESPONSE.WRITE STRSQL & "<BR>"
         call ofv.crearconsultaEx(StrSql,cn,1,volver)


      ELSEIF SACTIVIDAD <> "ADD" THEN

        if wkaplic = "OR" then
           WKCRITIK = REQUEST.FORM("CRITICO") & ""
           WKCRITIKDAT = "CRITICO"
        elseif wkaplic = "PL" then
           WKCRITIK = REQUEST.FORM("PRIORIDAD") & ""
           WKCRITIKDAT = "PRIORIDAD"
        else
           WKCRITIK = REQUEST.FORM("CRITICO") & ""
           WKCRITIKDAT = "CRITICO"
        end if

         WKNEWDSCMN = "" & REQUEST.FORM("NEWDSCMN")
         WKNEWDSC   = "0" & REQUEST.FORM("NEWDSC")
         WKNEWSEC   = "0" & REQUEST.FORM("NEWSEC")

         WKNEWDSCUPD = ""

       IF wspTIPO = "CR" THEN

         IF ISNUMERIC(WKNEWDSC) THEN
            IF WKNEWDSC > "0" THEN
               WKNEWDSCUPD = " TIPOACCION = 0" & WKNEWDSC & ","
            END IF

         END IF

           WKNEWDSCUPD = WKNEWDSCUPD & " DESCRIACC = '" & WKNEWDSCMN & "',"

       ELSE
               WKNEWDSCUPD = " DESCRIACC = '" & WKNEWDSCMN & "',"
       END IF
         WKNEWDSCUPD=ucase(WKNEWDSCUPD)
         strsql = "UPDATE OBJACCION SET " & WKCRITIKDAT & " = '" & WKCRITIK & "'," & WKNEWDSCUPD  & "SECUEN = " & WKNEWSEC & " "
         strsql = strsql & " WHERE CODACCION = '" & SACTIVIDAD & "' "
       ''''  RESPONSE.WRITE STRSQL & "<BR>"
         call ofv.crearconsultaEx(StrSql,cn,1,volver)

      end if


    END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' FIN MODIFICA LA CRITICIDAD DE UNA ACTIVIDAD
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   IF wspTIPO = "VN" THEN


       Strsql = "SELECT * From objlnk where objid = " & WKPROCVINC
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rs.eof then
           wtipoL   = clng(rs(6))
           wtipOBJ  = rs(7)
           wUrl     = rs(8)
           wUrlDesc = rs(9)
           wtipo    = rs(10)

           wNV1    = rs(11) & ""
           wNV2    = rs(12) & ""
           wNV3    = rs(13) & ""

           IF  WNV1 = "" THEN WNV1 = "0"
           IF  WNV2 = "" THEN WNV2 = "0"
           IF  WNV3 = "" THEN WNV3 = "0"


      END IF


          strsql = "SELECT C.DESCRI, D.DESCRI, B.CODSECTOR, A.CODPERFIL FROM OBJPROCED A "
          strsql = strsql & "     LEFT JOIN OBJSECTOR B ON (A.CODOBJETO = B.CODOBJETO AND A.CODSECTOR = B.OBJID)"
          strsql = strsql & "     LEFT JOIN ECO C ON (B.CODSECTOR = C.ECO)"
          strsql = strsql & "     LEFT JOIN TIPOPERFIL D ON (A.CODPERFIL = D.TIPOPERFIL) "
          strsql = strsql & " WHERE A.CODPROCED = '" & WNV2 & "' "


          set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsN.eof then

             WKSECRSP = RSN(0) & ""
             WKFUNRSP = RSN(1) & ""
             WKSECCOD = RSN(2) & ""
             WKFUNCOD = RSN(3) & ""

          END IF


   ELSEIF wspTIPO = "AI" THEN


       Strsql = "SELECT * From objlnk where objid = " & WKPROCVINC
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rs.eof then
           wtipoL   = clng(rs(6))
           wtipOBJ  = rs(7)
           wUrl     = rs(8)
           wUrlDesc = rs(9)
           wtipo    = rs(10)

           wNV1    = rs(11) & ""
           wNV2    = rs(12) & ""
           wNV3    = rs(13) & ""

           IF  WNV1 = "" THEN WNV1 = "0"
           IF  WNV2 = "" THEN WNV2 = "0"
           IF  WNV3 = "" THEN WNV3 = "0"


      END IF


          strsql = "SELECT C.DESCRI, D.DESCRI, B.CODSECTOR, A.CODPERFIL FROM OBJPROCED A "
          strsql = strsql & "     LEFT JOIN OBJSECTOR B ON (A.CODOBJETO = B.CODOBJETO AND A.CODSECTOR = B.OBJID)"
          strsql = strsql & "     LEFT JOIN ECO C ON (B.CODSECTOR = C.ECO)"
          strsql = strsql & "     LEFT JOIN TIPOPERFIL D ON (A.CODPERFIL = D.TIPOPERFIL) "
          strsql = strsql & " WHERE A.CODPROCED = '" & WNV2 & "' "


          set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsN.eof then

             WKSECRSP = RSN(0) & ""
             WKFUNRSP = RSN(1) & ""
             WKSECCOD = RSN(2) & ""
             WKFUNCOD = RSN(3) & ""

          END IF

             IF SACTIVIDAD = "ADD" THEN

                GETFECH()

                WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
                WKHORA  = XLHORDAT

                WKFORDEN = 10


                strsql = "SELECT MAX(SECUEN) FROM OBJACCION "
                strsql = strsql & " WHERE CODOBJETO = " & wtipoL & " AND TIPOBJETO = '" & wtipOBJ & "' "
                strsql = strsql & " AND CODPROCED = '" & wNV2 & "' "
                set rsN2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsN2.eof then

                   WKFORDEN = rsN2(0) + 10

                end if


                sevento = 0
                sukey = GTkey(wtipoL)

                strsql = "insert into objaccion  ("
                strsql = strsql & "codobjeto,tipobjeto,codproced,codaccion,"
                strsql = strsql & "secuen,tipproced,descriacc,MODI,"
                strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA,"
                strsql = strsql & "CRITICO"
                strsql = strsql & ") values ("
                strsql = strsql & wtipoL & ",'" & wtipOBJ & "','" & wNV2 & "','" & sukey & "',"
                strsql = strsql & "'" & WKFORDEN & "','" & wkaplic & "','NUEVO ACTIVO','1',"
                strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "',"
                strsql = strsql & "'SI'"
                strsql = strsql & ")"
                xxx=SqlExLog("INSERT",strsql,CN,"objaccion","","","")
                'call ofv.crearconsultaEx(StrSql,cn,1,volver)


             end if

   ELSEIF wspTIPO = "CR" THEN


       Strsql = "SELECT * From objlnk where objid = " & WKPROCVINC
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rs.eof then
           wtipoL   = clng(rs(6))
           wtipOBJ  = rs(7)
           wUrl     = rs(8)
           wUrlDesc = rs(9)
           wtipo    = rs(10)

           wNV1    = rs(11) & ""
           wNV2    = rs(12) & ""
           wNV3    = rs(13) & ""

           IF  WNV1 = "" THEN WNV1 = "0"
           IF  WNV2 = "" THEN WNV2 = "0"
           IF  WNV3 = "" THEN WNV3 = "0"


      END IF


          strsql = "SELECT C.DESCRI, D.DESCRI, B.CODSECTOR, A.CODPERFIL FROM OBJPROCED A "
          strsql = strsql & "     LEFT JOIN OBJSECTOR B ON (A.CODOBJETO = B.CODOBJETO AND A.CODSECTOR = B.OBJID)"
          strsql = strsql & "     LEFT JOIN ECO C ON (B.CODSECTOR = C.ECO)"
          strsql = strsql & "     LEFT JOIN TIPOPERFIL D ON (A.CODPERFIL = D.TIPOPERFIL) "
          strsql = strsql & " WHERE A.CODPROCED = '" & WNV2 & "' "

        '''  RESPONSE.WRITE STRSQL & "<BR>"


          set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsN.eof then

             WKSECRSP = RSN(0) & ""
             WKFUNRSP = RSN(1) & ""
             WKSECCOD = RSN(2) & ""
             WKFUNCOD = RSN(3) & ""


          END IF

             IF SACTIVIDAD = "ADD" THEN

                GETFECH()

                WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
                WKHORA  = XLHORDAT

                WKFORDEN = 10


                strsql = "SELECT MAX(SECUEN) FROM OBJACCION "
                strsql = strsql & " WHERE CODOBJETO = " & wtipoL & " AND TIPOBJETO = '" & wtipOBJ & "' "
                strsql = strsql & " AND CODPROCED = '" & wNV2 & "' "
                set rsN2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsN2.eof then

                   WKFORDEN = rsN2(0) + 10

                end if


                sevento = 0
                sukey = GTkey(wtipoL)

                strsql = "insert into objaccion  ("
                strsql = strsql & "codobjeto,tipobjeto,codproced,codaccion,"
                strsql = strsql & "secuen,tipproced,descriacc,MODI,"
                strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA,"
                strsql = strsql & "CRITICO"
                strsql = strsql & ") values ("
                strsql = strsql & wtipoL & ",'" & wtipOBJ & "','" & wNV2 & "','" & sukey & "',"
                strsql = strsql & "'" & WKFORDEN & "','" & wkaplic & "','NUEVA ACTIVIDAD','1',"
                strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "',"
                strsql = strsql & "'SI'"
                strsql = strsql & ")"
                'call ofv.crearconsultaEx(StrSql,cn,1,volver)
                xxx=SqlExLog("INSERT",strsql,CN,"objaccion","","","")


             end if


   ELSE


          strsql = "SELECT C.DESCRI, D.DESCRI, A.CODPROCED, B.CODSECTOR, A.CODPERFIL FROM OBJPROCED A "
          strsql = strsql & "     LEFT JOIN OBJSECTOR B ON (A.CODOBJETO = B.CODOBJETO AND A.CODSECTOR = B.OBJID)"
          strsql = strsql & "     LEFT JOIN ECO C ON (B.CODSECTOR = C.ECO)"
          strsql = strsql & "     LEFT JOIN TIPOPERFIL D ON (A.CODPERFIL = D.TIPOPERFIL) "
          strsql = strsql & " WHERE A.CODOBJETO = " & scodmaster & " "


          set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsN.eof then

             WKSECRSP = RSN(0) & ""
             WKFUNRSP = RSN(1) & ""
             WKSECCOD = RSN(3) & ""
             WKFUNCOD = RSN(4) & ""
             wNV2     = RSN(2) & ""

             IF SACTIVIDAD = "ADD" THEN

                WKFORDEN = 10


                strsql = "SELECT MAX(SECUEN) FROM OBJACCION "
                strsql = strsql & " WHERE CODOBJETO = " & scodmaster & " AND TIPOBJETO = '" & sTIPmaster & "' "
                strsql = strsql & " AND CODPROCED = '" & wNV2 & "' "
                set rsN2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsN2.eof then

                   WKFORDEN = rsN2(0) + 10

                end if


                sevento = 0
                sukey = GTkey(scodmaster)

                strsql = "insert into objaccion  ("
                strsql = strsql & "codobjeto,tipobjeto,codproced,codaccion,"
                strsql = strsql & "secuen,tipproced,descriacc,MODI,"
                strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA,"
                strsql = strsql & "CRITICO"
                strsql = strsql & ") values ("
                strsql = strsql & scodmaster & ",'" & sTIPmaster & "','" & wNV2 & "','" & sukey & "',"
                strsql = strsql & "'" & WKFORDEN & "','" & wkaplic & "','NUEVA ACTIVIDAD','1',"
                strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "',"
                strsql = strsql & "'SI'"
                strsql = strsql & ")"
               ' call ofv.crearconsultaEx(StrSql,cn,1,volver)
                xxx=SqlExLog("INSERT",strsql,CN,"objaccion","","","")


             end if



          ELSE

             GETFECH()

             WKFECHA = XLFECDAT - clng(cdate("01-01-1900"))
             WKHORA  = XLHORDAT
             WKSECTOR = 0

             strsql = "insert into OBJSECTOR  ("
             strsql = strsql & "codobjeto,tipobjeto,codSECTOR,MODI,"
             strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA"
             strsql = strsql & ") values ("
             strsql = strsql & scodmaster & ",'" & sTIPmaster & "','0','1',"
             strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "'"
             strsql = strsql & ")"
             'call ofv.crearconsultaEx(StrSql,cn,1,volver)
             xxx=SqlExLog("INSERT",strsql,CN,"OBJSECTOR","","","")

             strsql = "SELECT OBJID FROM OBJSECTOR "
             strsql = strsql & " WHERE CODOBJETO = " & scodmaster & " "
             set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsN.eof then
                WKSECTOR = rsN(0)
             END IF


             sevento = 0
             sukey = GTkey(scodmaster)

             strsql = "insert into objproced  ("
             strsql = strsql & "codobjeto,tipobjeto,codproced,codeven,"
             strsql = strsql & "frecuencia,tipproced,CODSECTOR,MODI,"
             strsql = strsql & "TIPMODI,MODAUTOR,MODFECHA,MODHORA"
             strsql = strsql & ") values ("
             strsql = strsql & scodmaster & ",'" & sTIPmaster & "','" & sukey & "','" & sevento & "',"
             strsql = strsql & "'DIARIA','" & wkaplic & "','" & WKSECTOR & "','1',"
             strsql = strsql & "'A','" & SESSION("USUARIO") & "','" & WKFECHA & "','" & WKHORA & "'"
             strsql = strsql & ")"
             'call ofv.crearconsultaEx(StrSql,cn,1,volver)
             xxx=SqlExLog("INSERT",strsql,CN,"objproced","","","")


             strsql = "SELECT C.DESCRI, D.DESCRI, A.CODPROCED, B.CODSECTOR, A.CODPERFIL FROM OBJPROCED A "
             strsql = strsql & "     LEFT JOIN OBJSECTOR B ON (A.CODOBJETO = B.CODOBJETO AND A.CODSECTOR = B.OBJID)"
             strsql = strsql & "     LEFT JOIN ECO C ON (B.CODSECTOR = C.ECO)"
             strsql = strsql & "     LEFT JOIN TIPOPERFIL D ON (A.CODPERFIL = D.TIPOPERFIL) "
             strsql = strsql & " WHERE A.CODOBJETO = " & scodmaster & " "
             set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsN.eof then

                WKSECRSP = RSN(0) & ""
                WKFUNRSP = RSN(1) & ""
                WKSECCOD = RSN(3) & ""
                WKFUNCOD = RSN(4) & ""
                wNV2     = RSN(2) & ""
             END IF


          END IF




   END IF

'''''''''''''''''''''''''''''''''



'''''''''''''''''''''''''''''''''

           if wkaplic = "OR" then

              IF SMETODO = "AUTOEVAL" THEN

                 strsql = "select a.objid,b.actividad,a.*, C.codinx2, COR.ACTIVIDAD AS CORY, CPL.CODACCION AS CPLY from OBJACCION a "

                 strsql = strsql & " left join (select distinct actividad "
                 strsql = strsql & " from oractivriesgo where periodo = 0" & session("ORPERIODO") & ") b "
                 strsql = strsql & " on a.codaccion = b.actividad "


                 strsql = strsql & " left join (select distinct actividad "
                 strsql = strsql & " from oractivriesgo ) COR "
                 strsql = strsql & " on a.codaccion = COR.actividad "

                 strsql = strsql & " left join ("
                 strsql = strsql & " SELECT DISTINCT W.CODACCION "
                 strsql = strsql & "   FROM OBJACCION W "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCALT) Y ON (W.CODACCION = Y.CODACCION) "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCORR) Z ON (W.CODACCION = Z.CODACCION) "
                 strsql = strsql & "  WHERE (Y.CODACCION <> '' ) "
                 strsql = strsql & "     OR (Z.CODACCION <> '' ) "
                 strsql = strsql & " ) CPL "
                 strsql = strsql & " on a.codaccion = CPL.codaccion "


'                strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null ) C on (A.objid = C.codinx2) "
                ' strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT  ) C on (A.objid = C.codinx2) "
                strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT  WHERE TIPO='OBS' ) C on (cast (A.objid as nvarchar) = C.codinx2) "

                 strsql = strsql & " WHERE a.codproced = '" &  WNV2 & "' "
                 strsql = strsql & " and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
                 strsql = strsql & " Order By a.secuen"

              else

                 strsql = "select a.objid,a.objid,a.*, COR.ACTIVIDAD AS CORY, CPL.CODACCION AS CPLY from OBJACCION a "

                 strsql = strsql & " left join (select distinct actividad "
                 strsql = strsql & " from oractivriesgo ) COR "
                 strsql = strsql & " on a.codaccion = COR.actividad "

                 strsql = strsql & " left join ("
                 strsql = strsql & " SELECT DISTINCT W.CODACCION "
                 strsql = strsql & "   FROM OBJACCION W "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCALT) Y ON (W.CODACCION = Y.CODACCION) "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCORR) Z ON (W.CODACCION = Z.CODACCION) "
                 strsql = strsql & "  WHERE (Y.CODACCION <> '' ) "
                 strsql = strsql & "     OR (Z.CODACCION <> '' ) "
                 strsql = strsql & " ) CPL "
                 strsql = strsql & " on a.codaccion = CPL.codaccion "


                 strsql = strsql & " WHERE a.codproced = '" &  WNV2 & "' "
                 strsql = strsql & " and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
                 strsql = strsql & " Order By a.secuen"

              end if

           ELSEif wkaplic = "PL" then

              IF SMETODO = "AUTOEVAL" THEN

                 strsql = "select a.objid,b.codaccion,a.*, C.codinx2, COR.ACTIVIDAD AS CORY, CPL.CODACCION AS CPLY from OBJACCION a "

                 strsql = strsql & " left join ("
                 strsql = strsql & " SELECT DISTINCT W.CODACCION "
                 strsql = strsql & "   FROM OBJACCION W "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION, PERIODO FROM PLCIRCALT) Y ON (W.CODACCION = Y.CODACCION) "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION, PERIODO FROM PLCIRCORR) Z ON (W.CODACCION = Z.CODACCION) "
                 strsql = strsql & "  WHERE (Y.CODACCION <> '' AND Y.PERIODO = 0" & session("ORPERIODO") & " ) "
                 strsql = strsql & "     OR (Z.CODACCION <> '' AND Z.PERIODO = 0" & session("ORPERIODO") & " ) "
                 strsql = strsql & " ) b "
                 strsql = strsql & " on a.codaccion = b.codaccion "


                 strsql = strsql & " left join (select distinct actividad "
                 strsql = strsql & " from oractivriesgo ) COR "
                 strsql = strsql & " on a.codaccion = COR.actividad "

                 strsql = strsql & " left join ("
                 strsql = strsql & " SELECT DISTINCT W.CODACCION "
                 strsql = strsql & "   FROM OBJACCION W "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCALT) Y ON (W.CODACCION = Y.CODACCION) "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCORR) Z ON (W.CODACCION = Z.CODACCION) "
                 strsql = strsql & "  WHERE (Y.CODACCION <> '' ) "
                 strsql = strsql & "     OR (Z.CODACCION <> '' ) "
                 strsql = strsql & " ) CPL "
                 strsql = strsql & " on a.codaccion = CPL.codaccion "



     '            strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null ) C on (A.objid = C.codinx2) "
             '     strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT ) C on (A.objid = C.codinx2) "
                 strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT WHERE TIPO='OBS') C on ( cast (A.objid as nvarchar) = C.codinx2) "
                 strsql = strsql & " WHERE a.codproced = '" &  WNV2 & "' "
                 strsql = strsql & " and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
                 strsql = strsql & " Order By a.secuen"

                '' response.write strsql & "<br>" & chr(13)

              else

                 strsql = "select a.objid,a.objid,a.*, COR.ACTIVIDAD AS CORY, CPL.CODACCION AS CPLY from OBJACCION a "

                 strsql = strsql & " left join (select distinct actividad "
                 strsql = strsql & " from oractivriesgo ) COR "
                 strsql = strsql & " on a.codaccion = COR.actividad "

                 strsql = strsql & " left join ("
                 strsql = strsql & " SELECT DISTINCT W.CODACCION "
                 strsql = strsql & "   FROM OBJACCION W "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCALT) Y ON (W.CODACCION = Y.CODACCION) "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCORR) Z ON (W.CODACCION = Z.CODACCION) "
                 strsql = strsql & "  WHERE (Y.CODACCION <> '' ) "
                 strsql = strsql & "     OR (Z.CODACCION <> '' ) "
                 strsql = strsql & " ) CPL "
                 strsql = strsql & " on a.codaccion = CPL.codaccion "


                 strsql = strsql & " WHERE a.codproced = '" &  WNV2 & "' "
                 strsql = strsql & " and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
                 strsql = strsql & " Order By a.secuen"

              end if


           else


              strsql = "select a.objid,a.objid,a.*, C.codinx2, COR.ACTIVIDAD AS CORY, CPL.CODACCION AS CPLY from OBJACCION a "

                 strsql = strsql & " left join (select distinct actividad "
                 strsql = strsql & " from oractivriesgo ) COR "
                 strsql = strsql & " on a.codaccion = COR.actividad "

                 strsql = strsql & " left join ("
                 strsql = strsql & " SELECT DISTINCT W.CODACCION "
                 strsql = strsql & "   FROM OBJACCION W "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCALT) Y ON (W.CODACCION = Y.CODACCION) "
                 strsql = strsql & "   LEFT JOIN (SELECT DISTINCT CODACCION FROM PLCIRCORR) Z ON (W.CODACCION = Z.CODACCION) "
                 strsql = strsql & "  WHERE (Y.CODACCION <> '' ) "
                 strsql = strsql & "     OR (Z.CODACCION <> '' ) "
                 strsql = strsql & " ) CPL "
                 strsql = strsql & " on a.codaccion = CPL.codaccion "


'              strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null ) C on (A.objid = C.codinx2) "
            '  strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT  ) C on (A.objid = C.codinx2) "
              strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT WHERE TIPO='OBS' ) C on ( cast(A.objid as nvarchar)= C.codinx2) "
              strsql = strsql & " WHERE a.codproced = '" &  WNV2 & "' "
              strsql = strsql & " and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
              strsql = strsql & " Order By a.secuen"

           end if



	'RESPONSE.Write(CHR(13) & STRSQL & CHR(13))




   END IF

     Tabla = ""
     fila = ""

   IF wspTIPO <> "SP" AND wspTIPO <> "SPI" and SMODAL = "LISTA" THEN


      SESSION("WKSECCOD")  =  WKSECCOD
      SESSION("WKFUNCOD")  =  WKFUNCOD




     Celda = ""
     Celda = Celda & ofv.GenCelda("","UTI style='background-color:#888888;color:#ffffff;' ","","15%","","","","AREA/SECTOR: ","si")

      IF wspTIPO = "CR" or wspTIPO = "AI" THEN


       IF WKSPCIRCMO = "S" THEN



          sformobj = "FRMSEC"
          Strsql1 = "Select * From Eco Order By descri"
          sInput = ofv.GenerarCombo(sformobj,"NEWSEC",WKSECCOD,Strsql1,cn,0,1,"")

          sAccion     = Session("ASPSTR") & "&ACTIVIDAD=SEC&CROBJ=" & wtipoL

          sInput = ofv.GenCelda("","UT ","","20%","","","",sInput,"si")

          formobj1    = ofv.genform(sformobj,"",sAccion,"","post","",sInput,"si")

          Celda = Celda & formobj1


          Celda = Celda & ofv.GenCelda("","UT ","","2%","","","","&nbsp;-&nbsp;","si")

          sformobj = "FRMFNC"

          Strsql1 = "SELECT TIPOPERFIL.tipoperfil,TIPOPERFIL.descri "
          Strsql1 = Strsql1 & "FROM ECO, ECOPERFIL, TIPOPERFIL "
          Strsql1 = Strsql1 & " where  "
          Strsql1 = Strsql1 & " ECO.eco = ECOPERFIL.codsector and "
          Strsql1 = Strsql1 & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
          Strsql1 = Strsql1 & " and Eco.eco = 0" & WKSECCOD

          sInput = ofv.GenerarCombo(sformobj,"NEWFNC",WKFUNCOD,Strsql1,cn,0,1,"")

          sAccion     = Session("ASPSTR") & "&ACTIVIDAD=FNC&CROBJ=" & wtipoL

          sInput = ofv.GenCelda("","UT ","","20%","","","",sInput,"si")

          formobj2    = ofv.genform(sformobj,"",sAccion,"","post","",sInput,"si")

          Celda = Celda & formobj2

          Celda = Celda & ofv.GenCelda("","UT ","","60%","","","","&nbsp;","si")

       ELSE

          Celda = Celda & ofv.GenCelda("","UT ","","85%","","","",WKSECRSP & "&nbsp;-&nbsp;" & WKFUNRSP,"si")

       end if

      ELSE

          Celda = Celda & ofv.GenCelda("","UT ","","85%","","","",WKSECRSP & "&nbsp;-&nbsp;" & WKFUNRSP,"si")

      END IF





     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#000000;' BGCOLOR='#DDDDFF' ","","100% ","","0","","1","2","","",Fila,"si")

     fila = ""

      IF wspTIPO = "CR" THEN


      ELSE


      END IF


  END IF


 ' response.write strsql & "<br>" & chr(13)
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


     IF wspTIPO <> "SP" AND wspTIPO <> "SPI" and SMODAL = "LISTA" THEN






        scomK = "SI" & chr(9) & "NO"
        scomD = "SI" & chr(9) & "NO"

       CeldaK  = ""
       Celda  = ""

      IF wspTIPO = "AI" THEN

       Celda  = Celda & ofv.GenCelda("","uti colspan=6 ","","","","","","ACTIVOS","si")

      else

       IF wspTIPO = "CR" THEN

          Celda  = Celda & ofv.GenCelda("","uti colspan=7 ","","","","","","ACTIVIDADES","si")

       else
          Celda  = Celda & ofv.GenCelda("","uti colspan=6 ","","","","","","ACTIVIDADES","si")
          Celda  = Celda & ofv.GenCelda("","Ut ","CENTER","","","","","CRITICO","si")

       end if

      end if




       IF SMETODO = "AUTOEVAL" THEN


          IF WKAPLIC = "OR" THEN


             Celdak  = Celdak & ofv.GenCelda("","Ut ","CENTER","","","","","&nbsp;RIESGOS&nbsp;&nbsp;","si")

          ELSEIF WKAPLIC = "PL" THEN


             Celdak  = Celdak & ofv.GenCelda("","Ut ","CENTER","","","","","&nbsp;CONTINGENCIAS&nbsp;&nbsp;","si")

          ELSE


             Celdak  = Celdak & ofv.GenCelda("","Ut ","CENTER","","","","","&nbsp;----&nbsp;&nbsp;","si")


          END IF


       END IF

       Celda = Celda & Celdak
       FilaN = filaN & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","25","","",Celda,"si")

     end if

       wkcoddelete = "N"
       wktitdelete = " "
       sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
       sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"


    do until rs.eof

       wkcoddelete = "N"
       wktitdelete = " "

     IF wspTIPO <> "SP" AND wspTIPO <> "SPI" THEN

       Form2N     = "ORCRITICIDAD" & rs(0)




       wkproccod = "javascript:void(0);"

       sHref     = wkproccod
       sImage    = "<font face='corflaticon' size=1 color='#FFFFFF'>&#060;</font>"
       sLink     = ofv.GenLink("","nt",sHref,"","","",sImage,"si")

     IF wspTIPO = "VN" THEN

        wspcolor  = "#336699"
        sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

       'Acciones
        if rs(6) <> "" then
           Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs(6)
           set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsAcciones.eof then
               sAccionM = rsAcciones(0)
               sModo    = rsAcciones(1)
           end if
           call ofv.cerrarconsulta(rsAcciones)
         end if

          if SMODAL = "LISTA"  THEN  sAccionM = rs(12) & "." & sAccionM

          wkproccod = "" & sAccionM & "" & ofv.GenerarInput("NEWSEC",rs(12),"HIDDEN",Form2N,"1","dt","")
          wktitdelete = "No puede eliminar una actividad de un proceso vinculado."

     ELSE

         IF wspTIPO = "AI" THEN

            wspcolor  = "#aa00aa"
            sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

         else

            wspcolor  = "#F7AA22"
            sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

         end if


      IF  WKSPCIRCMO = "S" THEN


          if ((RS("CORY") & "") <> "" OR (RS("CPLY") & "") <> "") then




            IF wspTIPO = "CR" THEN
               sAccionM = ""
               if rs(6) <> "" then
                  Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs(6)
                  set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if  not rsAcciones.eof then
                   '   sAccionM = rsAcciones(0) & "<br>"
                      sModo    = rsAcciones(1)
                  end if
                  call ofv.cerrarconsulta(rsAcciones)
               end if

      '         wktitdelete = "No puede eliminar esta actividad, tiene circuitos definidos."
      '     ELSEIF wspTIPO = "CR" THEN
               sAccionM = sAccionM & rs(30)
               wktitdelete = "No puede eliminar esta actividad, tiene circuitos definidos."
            ELSE
               sAccionM = rs(30)
               wktitdelete = "No puede eliminar este activo, tiene circuitos definidos."
            END IF


               sAccionM = rs(12) & "." & sAccionM

               wkproccod = "" & sAccionM & ""
               wkproccod2 = ""
               wkcoddelete = "N"


          ELSE

            IF wspTIPO = "CR" THEN

               sInput = ofv.GenerarInput("NEWDSC","0","hidden","","1","dt","")

               Strsql = "SELECT tipoaccionalt,descri From TIPOACCALT Order By descri"
            '   sInput = ofv.GenerarCombo(Form2N,"NEWDSC",rs(6).value,Strsql,cn,0,1,"") & "<br>"

          '     wktitdelete = "Confirma la eliminacion de la actividad???."

         '   ELSEIF wspTIPO = "CR" THEN

          nCol = 80

           wdescri = " "

          if  not  isnull(rs(30)) then
              wdescri = rs(30) & ""
          end if

          t3b = "<font class=at><textarea class=dt STYLE='FONT-SIZE:7PT;BACKGROUND-COLOR:#EEEEFF;' "
          t3b = t3b & "name=" & "NEWDSCMN" & " cols=" & nCol & " rows=2 "
        '  if wkmodif = "S" then
              t3b = t3b & " onblur=" & Form2N & ".submit() "
              t3b = t3b & " onkeypress=" & chr(34)
              t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
              t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       '   else
        '      t3b = t3b & " readonly >"
       '   end if
          xdatos =  wdescri
          t3b    = t3b & xdatos & "</textarea></font>"
          SINPUT = SINPUT & t3b

             ''  sInput = ofv.GenerarInput("NEWDSCMN",rs(30),"text",Form2N,"70","dt","")
               wktitdelete = "Confirma la eliminacion de la actividad???."

            ELSE

               sInput = ofv.GenerarInput("NEWDSCMN",rs(30),"text",Form2N,"70","dt","")
               wktitdelete = "Confirma la eliminacion del activo???."
            END IF

             wkproccod2 = ofv.GenerarInput("NEWSEC",rs(12),"text",Form2N,"3","dt","") & "."
             wkproccod = sInput
             wkcoddelete = "S"

          END IF

      ELSE



            IF wspTIPO = "CR" THEN
               if rs(6) <> "" then
                  Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs(6)
                  set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if  not rsAcciones.eof then
                      sAccionM = rsAcciones(0) & "<br>"
                      sModo    = rsAcciones(1)
                  end if
                  call ofv.cerrarconsulta(rsAcciones)
               end if
           '''    wktitdelete = "No tiene autorizada la eliminacion de la actividad."
           ''' ELSEIF wspTIPO = "CR" THEN
               sAccionM = sAccionM & rs(30)
               wktitdelete = "No tiene autorizada la eliminacion de la actividad."
            ELSE
               sAccionM = rs(30)
               wktitdelete = "No tiene autorizada la eliminacion del activo."
            END IF


          sAccionM = rs(12) & "." & sAccionM

          wkproccod = "" & sAccionM & ""
          wkproccod2 = ""
          wkcoddelete = "N"


      END IF




     END IF




     IF SMODAL = "GRAF" THEN

       Filan  = ""
       Filan1 = ""
       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","aaa style='background-color:" & wspcolor & ";color:#ffffff;font-family:tahoma;font-size:7pt;' ","","99%","","","","ACTIVIDAD","si")
       Celda  = Celda & ofv.GenCelda("","aaa style='background-color:" & wspcolor & ";color:#ffffff;font-size:7pt;' ","center","1","","","",sLink,"si")
     '  FilaN  = ofv.GenRow("","aaa bgcolor='#fefefe'","","","","","",Celda,"si")

       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","aaa style='background-color:" & wspcolor & ";color:#ffffff;font-size:7pt;' valign='top' ","center","1","","","",sLink,"si")
       Celda  = Celda & ofv.GenCelda("","aaa colspan=2 STYLE='overflow:none;color:#888888;font-family:tahoma;font-size:7pt;'  ","center","","","","",wkproccod,"si")
       FilaN  = filaN & ofv.GenRow("","","","","40","","",Celda,"si")


     else

       ANCHOK = "99%"
       Celdak = ""
       Celda  = ""

       if wkcoddelete = "S" then

             sImageEL   = sImageELS
             sformobj   = "FRMOBJ" & rs(0)
             sAccion    = "eliminarproceso.asp?codobjeto=" & rs(0) & "&tipobjeto=" & rs(20) & "&OPCION=ACC"
             formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
             sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "



       else
             sImageEL   = sImageELN
             formobj    = ""
             sfunobj    = ""
       end if




       Celda     = ""


      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")



      if SMETODO = "AUTOEVAL" AND ((XFSENESTADOP = 5  or XFSENESTADOP = 6) and (rs("codinx2") & "") <> "")  then

               xclas     = "ut"
      	       sAccion   = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & WKSPCIRCCOD
      	       sAccion   = sAccion & "&Tipobjeto=RSK&ID=" & rs(0)  & "&ID2=" & WKSPCIRCCOD

      	       xsffun    = " onclick=""document.getElementById('DXA').style.display='block';"" "
      	       sHref     = sAccion
      	       sInput    = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver comentarios para la actividad.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"fesquemas","","",Celdaobs,"si")
      end if


      Celda = Celda & Celdaobs




       Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageEL,"si")



     IF wspTIPO = "VN" THEN


       Celda  = Celda & ofv.GenCelda("","ut","CENTER","2","","","",sImageP,"si")
       Celda  = Celda & ofv.GenCelda("","UT bgcolor=" & wspcolor & " COLSPAN=2 ","center","2","","","",sLink,"si")

     ELSE

       Celda  = Celda & ofv.GenCelda("","ut","CENTER","2","","","",sImageP,"si")
       Celda  = Celda & ofv.GenCelda("","UT bgcolor=" & wspcolor & " ","center","2","","","",sLink,"si")
       Celda  = Celda & ofv.GenCelda("","UT ","","2","","","",wkproccod2,"si")
     END IF


     IF WKAPLIC = "OR" THEN
       WKCRITIK = rs(28) & ""
       WKCRITIKDAT = rs(28).NAME
     ELSEIF WKAPLIC = "PL" THEN
       WKCRITIK = rs(29) & ""
       WKCRITIKDAT = rs(29).NAME
     ELSE
       WKCRITIK = rs(28) & ""
       WKCRITIKDAT = rs(28).NAME
     END IF

       IF WKCRITIK <> "" THEN

       else
          WKCRITIK = "NO"
       end if

       Celda = Celda & ofv.GenCelda("","UtI","",ANCHOK,"","","",wkproccod,"si")




       IF SMETODO = "AUTOEVAL" THEN

     ''''''    if XFSENPERIODO = "S" and WKSPCIRCMO = "S" then
         if WKSPCIRCMO = "S" then


          if WKCRITIK = "NO"  or XFSENPERIODO <> "S" then

              Form1      = ""

             sImageO0   = "4"
             wspcolor0  = "#DDDDDD"

             wsstyloF   = " style='border-style:OUTSET;border-width:1;"
             wsstyloF   = wsstyloF & "background-color:" & wspcolor0 & ";' "

             sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"
             Celdak  = Celdak & ofv.GenCelda("","Ut title='NO DISPONIBLE PARA ACTIVIDADES NO CRITICAS o CIRCUITOS NO INICIADOS.' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")


          else







              sImageO0   = "4"
              wspcolor0  = "#000099"

              if (rs(1) & "") = "" then

                 sImageO0   = sImageO1
                 wspcolor0  = wspcolor1

              else

                 sImageO0   = sImageO2
                 wspcolor0  = wspcolor2
                 strsql="select * from oractivriesgo where actividad='"&rs("actividad")&"' and evento='"&wnv2&"' and periodo=" & session("orperiodo")
                  set rsAccionesPx = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                  if not rsAccionesPx.eof then

                   if rsAccionesPx("itemriesgo")=0 and rsAccionesPx("impacto")=0 and rsAccionesPx("riesgorepalc") &"" ="" and rsAccionesPx("probabilidad")=0 then
                    sImageO0   = sImageO1
                    wspcolor0  = wspcolor1
                   end if
                  end if
                  call ofv.cerrarconsulta(rsAccionesPx)
                 'response.write strsql

              end if


           IF WKAPLIC = "OR" THEN


              Form1N     = "ORCIRCUITO" & rs(0)
              sAccion    = "ORCIRCUITO.ASP?ACTIVIDAD=" & rs(20) & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
              Form1      = ofv.genform(Form1N,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")


              wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
              wsstyloF   = wsstyloF & "background-color:" & wspcolor0 & ";' "

              sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"
              Celdak  = Celdak & ofv.GenCelda("","Ut title='INDICAR RIESGOS ASOCIADOS A LA ACTIVIDAD.' onclick='document." & Form1N & ".submit();' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")


           ELSEIF WKAPLIC = "PL" THEN


              Form1N     = "PLCIRCUITO" & rs(0)
              sAccion    = "PLCIRCUITO.ASP?ACTIVIDAD=" & rs(20) & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
              Form1      = ofv.genform(Form1N,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")


              wsstyloF   = " style='border-style:OUTSET;border-width:1;cursor:pointer;"
              wsstyloF   = wsstyloF & "background-color:" & wspcolor0 & ";' "

              sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"
              Celdak  = Celdak & ofv.GenCelda("","Ut title='CONTINGENCIAS ASOCIADAS A LA ACTIVIDAD.' onclick='document." & Form1N & ".submit();' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")
           ELSE

              Celdak  = Celdak & ofv.GenCelda("","Ut  ","CENTER","2","","","","---","si")

           END IF



          end if



'''' CRITICIDAD

          if WKCRITIK = "NO" then

             wspcolor0  = "#DDDDDD"

          else

             wspcolor0  = wspcolor2

          end if

          wsstyloF   = " style='background-color:" & wspcolor0 & ";COLOR:#FFFFFF;' "


''XFORPERIODOST
''XFSENESTADOP

       IF wspTIPO = "CR" THEN

          IF ( XFSENPERIODO = "S" and XFORPERIODOST <> "OK") OR XFSENESTADOP = 4 OR XFSENESTADOD = 4 THEN
            sInput = "&nbsp;"
          ELSE
            sInput = "&nbsp;<input type='text' name='" & WKCRITIKDAT & "' style='display:none;' value='" & WKCRITIK & "' > "
          END IF


          Celda  = Celda & ofv.GenCelda("","Ut ","","","","","",sInput,"si")

       else


          IF ( XFSENPERIODO = "S" and XFORPERIODOST <> "OK") OR XFSENESTADOP = 4 OR XFSENESTADOD = 4 THEN
            sInput = WKCRITIK
          ELSE
            sInput = ofv.gencomboFX(Form2N,WKCRITIKDAT,WKCRITIK,scomK,scomD," ")
          END IF


         Celda  = Celda & ofv.GenCelda("","Ut " & wsstyloF,"","","","","",sInput,"si")

       end if


'''''''''




       else


              Form1      = ""


             sImageO0   = "4"
             wspcolor0  = "#f7eedd"

             wsstyloF   = " style='border-style:OUTSET;border-width:1;"
             wsstyloF   = wsstyloF & "background-color:" & wspcolor0 & ";' "

             sImageOF   = WKCRITIK
             Celdak  = Celdak & ofv.GenCelda("","Ut title='Debe habilitar el periodo actual para el circuito.' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")

             sImageO0   = "4"
             wspcolor0  = "#f7eedd"

             wsstyloF   = " style='border-style:OUTSET;border-width:1;"
             wsstyloF   = wsstyloF & "background-color:" & wspcolor0 & ";' "

             sImageOF   = "<font face='corflaticon' size=2 color='#FFFFFF'  >" & sImageO0 & "</font>"
             Celdak  = Celdak & ofv.GenCelda("","Ut title='Debe habilitar el periodo actual para el circuito.' " & wsstyloF,"CENTER","2","","","",sImageOF,"si")


        end if



       END IF




       Celda = Celda & Celdak
       FilaX = ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


       sAccion    = Session("ASPSTR") & "&ACTIVIDAD=" & rs(20)
       FilaN      = filaN & ofv.genform(Form2N,"AAA STYLE='display:none;' ",sAccion,"","post","",FilaX,"si") & Form1 & formobj


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF



     end if




     ELSE
''''''''''''''''''''''''''''''

     XFNOVECRT = ""
     XFNOVEPROC = ""

     IF RS("CRTY") <> "" THEN
        XFNOVECRT = sImageCRT

     END IF

     WKSPESTADO = 3

      Strsql = "SELECT estado From objprty where codobjeto = 0" & rs(0)
      set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if  not rsX.eof then

          if rsx(0) = 4 then

             WKSPESTADO = 4



          end if


      end if









       if (rs(17) & "") = "SP" then
          wspcolor = "#993333"
          sImageP    = "&#0167;"

       ELSEif (rs(17) & "") = "SPI" then
          wspcolor = "#8899BB"
          sImageP    = "&#0167;"

       elseif (rs(17) & "") = "VN" then
          wspcolor = "#008800"
          sImageP    = "&#0167;"

       elseif (rs(17) & "") = "CR" then
          wspcolor = "#99CCDD"
          sImageP    = "&#0167;"

       elseif (rs(17) & "") = "AI" then
          wspcolor = "#0000CC"
          sImageP    = "&#0167;"

       else
          wspcolor = "#000000"
          sImageP    = "&#0167;"

       end if

       wspcolor2 = wspcolor

       if wkspestado = 4 then

          wspcolor = "#dddddd"
          sImageP    = "&#0172;"

       end if


'''C.SEGCT, C.SEGCO, C.SEGMO, C.SEGCE, c.SEGOWN, D.PERMITE, E.PERMITEX

     wkcoddeleteMST = "S"

     IF RS("SEGCT") THEN

        IF RS("SEGOWN") <> SESSION("USUARIO") THEN

           IF (RS("PERMITEX") & "") <> "OK" THEN

              wkcoddeleteMST = "N"

           END IF


        END IF

     ELSE

      wkcoddeleteMST = WKSPMASTCT

     END IF


     IF RS("CRTYO") <> "" THEN

        wkcoddeleteMST = "N"

     END IF


       IF ( RS("SEGCT")  OR RS("SEGCO")  OR RS("SEGCE")  OR RS("SEGMO") ) THEN
          sImagePER = "<font face='corflaticon' size=3 color='#F7AA22'>&#0207;</font>"
       ELSEIF ( WKSPMASTCT = "N" OR WKSPMASTMO = "N" OR WKSPMASTCO = "N" OR WKSPMASTCE = "N" ) THEN
          sImagePER = "<font face='corflaticon' size=3 color='#aa3399'>&#0207;</font>"
       ELSE
        '''  sImagePER = "" & WKSPMASTCT &  WKSPMASTMO &  WKSPMASTCO &  WKSPMASTCE
          sImagePER = ""
       END IF


          sImageP    =  sImagePER & "<font face='corflaticon' size=2 color='" & wspcolor2 & "'>" & sImageP & "</font>"







       IF  XFORPERIODOST = "OK" AND (rs(17) & "") = "VN" THEN

           WKPROCVINCsp = CLNG("0" & RS(15))

           Strsql = "SELECT * From objlnk where objid = 0" & WKPROCVINCsp

           set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if  not rsX.eof then
               wtipoL   = clng(rsX(6))
               wtipOBJ  = rsX(7)






          strsql = "SELECT B.FECBAJA, B.ESTADO FROM OBJETOS A INNER JOIN OBJPRTY B   "
          strsql = strsql & " ON A.CODOBJETO = B.CODOBJETO "
          strsql = strsql & " WHERE A.CODOBJETO = 0" & wtipoL
       '''   response.write strsql & "<br>"
          set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsN.eof then



             IF RSN(1) = 4 THEN

                IF (CSTR(RSN(0) & "") > CSTR(XFORPERIODOFD & "") AND CSTR(RSN(0) & "") < CSTR(XFORPERIODOFH & "")) THEN

                   XFNOVEPROC = "&nbsp;<FONT COLOR='#FF0000' SIZE=2 style='cursor:pointer;border-style:OUTSET;border-width:1;' "
                   XFNOVEPROC = XFNOVEPROC & " title='El subproceso fue reemplazado.' ><b>&#032;!&#032;</b></FONT>"

                END IF

             END IF

          END IF


         END IF


       END IF




       wkproccod = "emdetaproc.asp?Nombre=" & request.querystring("Nombre")
       wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster
       wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL & "&PROPIEDIN=S"
       wkproccod = wkproccod & "&CODPROCESOIN=" & rs(0) & "&TIPPROCESOIN=" & rs(2)
       wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

       sHref     = wkproccod
       sImage    = "<font face='corflaticon' size=2 color='#FFFFFF'>4</font>"
       sLink     = ofv.GenLink("","nt",sHref,"","","",sImage,"si")

       wkproccod = "<a class=nt onclick=""document.all.DX4.style.display='none';"" "

       if (rs(17) & "") = "AI" then
       ''   wkproccod = wkproccod & " href='ORCIRCUITOINFO.ASP?ACTIVIDAD=" & rs(0) & "&codmaster=" & rs(0) & "&TIPmaster=" & rs(2)  & "' "

          IF (RS(15) & "") > "0" THEN
             wkproccod = wkproccod & " href='emdetaproc.asp?Nombre=" & request.querystring("Nombre")
             wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & rs(0) & "&tipmaster=" & rs(2)
             wkproccod = wkproccod & "&CODNIVEL=" & (sCODNIVEL + 1) & "&MODAL=" & SMODAL & "&METODO=" & SMETODO & "' "
          ELSE
             wkproccod = wkproccod & " href='javascript:void(0);' TITLE='El subproceso no esta habilitado para acceder a los activos.' "
             XFNOVEPROC = XFNOVEPROC & "&nbsp;<FONT COLOR='#f7aa22' SIZE=2 face='corflaticon' style='cursor:pointer;border-style:OUTSET;border-width:1;' "
             XFNOVEPROC = XFNOVEPROC & " title='El subproceso no esta habilitado para acceder a los activos.' ><b>&#0120;</b></FONT>"

          END IF

       ELSEif (rs(17) & "") = "VN" OR (rs(17) & "") = "CR" then
          IF (RS(15) & "") > "0" THEN
             wkproccod = wkproccod & " href='emdetaproc.asp?Nombre=" & request.querystring("Nombre")
             wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & rs(0) & "&tipmaster=" & rs(2)
             wkproccod = wkproccod & "&CODNIVEL=" & (sCODNIVEL + 1) & "&MODAL=" & SMODAL & "&METODO=" & SMETODO & "' "
          ELSE
             wkproccod = wkproccod & " href='javascript:void(0);' TITLE='El subproceso no esta habilitado para acceder a las actividades.' "
             XFNOVEPROC = XFNOVEPROC & "&nbsp;<FONT COLOR='#f7aa22' SIZE=2 face='corflaticon' style='cursor:pointer;border-style:OUTSET;border-width:1;' "
             XFNOVEPROC = XFNOVEPROC & " title='El subproceso no esta habilitado para acceder a las actividades.' ><b>&#0120;</b></FONT>"

          END IF
       ELSE
          wkproccod = wkproccod & " href='emdetaproc.asp?Nombre=" & request.querystring("Nombre")
          wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & rs(0) & "&tipmaster=" & rs(2)
          wkproccod = wkproccod & "&CODNIVEL=" & (sCODNIVEL + 1) & "&MODAL=" & SMODAL & "&METODO=" & SMETODO & "' "
       END IF

       wkproccod = wkproccod & " >"
       wkproccod = wkproccod & rs(3)
       wkproccod = wkproccod & "</a>" & "&nbsp;" & XFNOVEPROC & "&nbsp;" & XFNOVECRT

     IF SMODAL = "GRAF" THEN


       Filan     = ""
       Filan1    = ""
       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut style='background-color:" & wspcolor & ";color:#ffffff;' ","","99%","","","","PROCESO","si")
       Celda = Celda & ofv.GenCelda("","aaa style='background-color:" & wspcolor & ";color:#ffffff;' ","center","1","","","",sLink,"si")
     '  FilaN = ofv.GenRow("","aaa bgcolor='#fefefe'","","","15","","",Celda,"si")



       Celda = ""
       Celda = Celda & ofv.GenCelda("","aaa style='background-color:" & wspcolor & ";color:#ffffff;font-family:tahoma;font-size:7pt;' valign='top' ","center","1","","","",sLink,"si")
       Celda = Celda & ofv.GenCelda("","aaa colspan=2 style='overflow:none;color:#888888;font-family:tahoma;font-size:7pt;' ","center","","","","",wkproccod,"si")
       FilaN = filaN & ofv.GenRow("","","","","40","","",Celda,"si")


     ELSE





     if wkcoddeleteMST = "S" then



       if (rs(17) & "") = "SP" OR (rs(17) & "") = "SPI"  then
          wktitdelete = "Debe eliminar todos los subprocesos dependientes previamente."
          Strsql = "Select * From relaobj  where  "
          Strsql = Strsql & " codobjmaster = '" & rs(0) & "' and tipobjmaster = '" & rs(2) & "'  "
          set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if rsx2.eof then
             wktitdelete = "Confirma la eliminacion del subproceso???."
             wkcoddelete = "S"
          end if


       elseif (rs(17) & "") = "VN" then
          wktitdelete = "Debe desvincular el subproceso previamente."


       elseif (rs(17) & "") = "CR" OR (rs(17) & "") = "AI" then
          wkcoddelete = "S"
          wktitdelete = "Confirma la eliminacion del subproceso y todas las actividades asociadas???."
       else
          wktitdelete = "No puede eliminar este tipo de objeto."
       end if


       if wkcoddelete = "S" then
          sImageEL = sImageELS
          sformobj = "FRMOBJ" & rs(0)
          sAccion    = "eliminarproceso.asp?codobjeto=" & rs(0) & "&tipobjeto=" & rs(2) & ""
        ''  sAccion    = "respproceso.asp?codobjeto=" & rs(0) & "&tipobjeto=" & rs(2) & ""
          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "


       else
          sImageEL = sImageELN
          formobj    = ""
          sfunobj    = ""
       end if

    else
          sImageEL = sImageELN
          formobj    = ""
          sfunobj    = ""
    end if


       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageEL,"si") & formobj
       Celda = Celda & ofv.GenCelda("","ut","CENTER","2","","","",sImageP,"si")
       Celda = Celda & ofv.GenCelda("","UT bgcolor=" & wspcolor & " ","center","2","","","",sLink,"si")
       Celda = Celda & ofv.GenCelda("","UtI","","99%","","","",wkproccod,"si")

       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF


     END IF




     END IF


     IF SMODAL = "GRAF" THEN

       Filan = Filan1 & ofv.GenTabla("","aaa style='background-color:#f9f9f9;border-style:solid;border-width:1;border-COLOR:" & wspcolor & ";' valign='middle' ","","120","","0","","0","1","","",Filan,"si")

       Celda = ""

       Celda = Celda & "<div id='DX" & rs(0) & "' style='position:absolute;left:" & wkleft & "px;top:" & wktop & "px;"
       Celda = Celda & "width:130;height:50;overflow:none;display:block;"
       Celda = Celda & "filter:shadow(color=#dddddd,direction=125);"
      ' Celda = Celda & "background-color:#f0f0f0;"
       Celda = Celda & "' >"
       Celda = Celda & Filan
       Celda = Celda & "</center></div>"

       Tabla = Tabla & Celda

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

     IF SMODAL = "LISTA" THEN

       Tabla = "<BR>" & Tabla & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")

     END IF

  end if

    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "function send_confirm(send_msg,send_href,send_aid)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(send_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('ok ');"
    sHTML = sHTML & "    var fr= document.getElementById(send_aid);"
    sHTML = sHTML & "    fr.href = send_href;"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion! ');"
  '  sHTML = sHTML & "    return void(0);"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


  sHTML = sHTML & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34) & ">" & Chr(13)

  sHTML = sHTML & " function Abre() " & Chr(13)
  sHTML = sHTML & " if window.event.button=2  then  " & Chr(13)
  sHTML = sHTML & "       Editar(window.event.srcElement.ID) " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " end function " & Chr(13) & Chr(13)

  sHTML = sHTML & " function AbreNew(abproceso) " & Chr(13)
  sHTML = sHTML & "       Editar(abproceso) " & Chr(13)
  sHTML = sHTML & " end function " & Chr(13) & Chr(13)

  sHTML = sHTML & " function Editar(Elemento) " & Chr(13)
  sHTML = sHTML & " if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
  sHTML = sHTML & "    exit function " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " dato = " & Chr(34) & "edicionProc.asp?"
  sHTML = sHTML & "PRCSUP=" & Chr(34) & " & ELEMENTO & " & Chr(34)
  sHTML = sHTML & "&ID=" & scodmaster & "-" & STIPMASTER & "X-X" & Chr(34) & Chr(13)
  sHTML = sHTML & " window.showModalDialog "
  sHTML = sHTML & "dato,window," & Chr(34) & "dialogWidth:400px;dialogHeight:360px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
  sHTML = sHTML & " end function " & Chr(13)


  sHTML = sHTML & " function EditarCR(Elemento,CRCOBJ,CRTOBJ) " & Chr(13)
  sHTML = sHTML & " if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
  sHTML = sHTML & "    exit function " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " dato = " & Chr(34) & "edicionProc.asp?SPTIPO=CR&"
  sHTML = sHTML & "ID=" & Chr(34) & " & CRCOBJ & " & Chr(34) & "-" & Chr(34) & " & CRTOBJ & " & Chr(34) & "X-X" & Chr(34) & Chr(13)
  sHTML = sHTML & " window.showModalDialog "
  sHTML = sHTML & "dato,window," & Chr(34) & "dialogWidth:400px;dialogHeight:360px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
  sHTML = sHTML & " end function " & Chr(13)


  sHTML = sHTML & " function EditarAI(Elemento,CRCOBJ,CRTOBJ) " & Chr(13)
  sHTML = sHTML & " if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
  sHTML = sHTML & "    exit function " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " dato = " & Chr(34) & "edicionProc.asp?SPTIPO=AI&"
  sHTML = sHTML & "ID=" & Chr(34) & " & CRCOBJ & " & Chr(34) & "-" & Chr(34) & " & CRTOBJ & " & Chr(34) & "X-X" & Chr(34) & Chr(13)
  sHTML = sHTML & " window.showModalDialog "
  sHTML = sHTML & "dato,window," & Chr(34) & "dialogWidth:400px;dialogHeight:360px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
  sHTML = sHTML & " end function " & Chr(13)

  sHTML = sHTML & "</SCRIPT> " & Chr(13)

'''  sHTML = sHTML & FormuP



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = ""

  IF SMODAL = "LISTA" THEN

      Celda = Celda & "<DIV id='DX01' style='margin-top:1;margin-left:1;display:block;"
      Celda = Celda & "width:95%;height:30;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
      Celda = Celda & "background-color:#f9f9f9;"
      Celda = Celda & "' >"
      Celda = Celda & wktabperiodo
      Celda = Celda & "</div>"

      Celda = Celda & "<DIV id='DX0' style='margin-top:1;margin-left:1;display:block;"
      Celda = Celda & "width:95%;height:30;overflow:auto;"
      Celda = Celda & "background-color:#f9f9f9;"
      Celda = Celda & "' >"
      Celda = Celda & TablaH
      Celda = Celda & "</div>"

      Celda = Celda & "<DIV id='DX2' style='margin-top:1;margin-left:1;display:block;"
      Celda = Celda & "width:95%;height:70%;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
      Celda = Celda & "background-color:" & session("APLBGD") & ";"
      Celda = Celda & "' >"
      Celda = Celda & TablaVH
      Celda = Celda & TABLA
      Celda = Celda & "</div>"


    '  Celda = Celda & "<div id='DX01' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:110px;display:block;"
    '  Celda = Celda & "width:95%;height:30;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
    '  Celda = Celda & "background-color:#f9f9f9;"
    '  Celda = Celda & "' >"
    '  Celda = Celda & wktabperiodo
    '  Celda = Celda & "</div>"

    '  Celda = Celda & "<div id='DX0' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:140px;display:block;"
    '  Celda = Celda & "width:95%;height:30;overflow:auto;"
    '  Celda = Celda & "background-color:#f9f9f9;"
    '  Celda = Celda & "' >"
    '  Celda = Celda & TablaH
    '  Celda = Celda & "</div>"

    '  Celda = Celda & "<div id='DX2' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:170px;display:block;"
    '  Celda = Celda & "width:95%;height:70%;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
    '  Celda = Celda & "background-color:" & session("APLBGD") & ";"
    '  Celda = Celda & "' >"
    '  Celda = Celda & TablaVH
    '  Celda = Celda & TABLA
    '  Celda = Celda & "</div>"


  else



      Celda = Celda & "<div id='DX01' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:80px;display:block;"
      Celda = Celda & "width:800;height:30;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
      Celda = Celda & "background-color:#f9f9f9;"
      Celda = Celda & "' >"
      Celda = Celda & wktabperiodo
      Celda = Celda & "</div>"

      Celda = Celda & "<div id='DX0' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:110px;display:block;"
      Celda = Celda & "width:800;height:30;overflow:auto;"
      Celda = Celda & "background-color:#f9f9f9;"
      Celda = Celda & "' >"
      Celda = Celda & TablaH
      Celda = Celda & "</div>"

      Celda = Celda & "<div id='DX2' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:140px;display:block;"
      Celda = Celda & "width:800;height:420;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
      Celda = Celda & "background-color:" & session("APLBGD") & ";"
      Celda = Celda & "' >"
      Celda = Celda & TABLA
      Celda = Celda & "</div>"

 end if

  sHTML = sHTML & Celda

'''''''''''''''''''''''''''''''''''''''''''''
'' PROPIEDADES ''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''

 sPROPIEDIN     = request.querystring("PROPIEDIN")
 sCODPROCESOIN   = request.querystring("CODPROCESOIN")
 sTIPPROCESOIN   = request.querystring("TIPPROCESOIN")


 IF sPROPIEDIN = "S" THEN


  set cn = ofv.conectar(ofv.strconn2)

  sForm = "EDPROCF"

  wkproccod = "../emdetaproc.asp?Nombre=" & request.querystring("Nombre")
  wkproccod = wkproccod & "&menu=" & sMenu & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster
  wkproccod = wkproccod & "&CODNIVEL=" & sCODNIVEL & "&PROPIEDIN=S"
  wkproccod = wkproccod & "&CODPROCESOIN=" & sCODPROCESOIN & "&TIPPROCESOIN=" & sTIPPROCESOIN
  wkproccod = wkproccod & "&MODAL=" & SMODAL & "&METODO=" & SMETODO

  VUELTA = REPLACE(wkproccod,"&","_")

  Session("ASPSTRD") = REPLACE(wkproccod,"../","")

  sImagenI3 = "<B><FONT FACE=corflaticon COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"


      WKSPPERMCT = WKSPMASTCT
      WKSPPERMMO = WKSPMASTMO
      WKSPPERMCO = WKSPMASTCO
      WKSPPERMCE = WKSPMASTCE



      Strsql = "Select A.*, C.SEGCT, C.SEGCO, C.SEGMO, C.SEGCE, c.SEGOWN, D.PERMITECT, E.PERMITEMO, F.PERMITECE From OBJETOS A"
      Strsql = Strsql & " INNER JOIN OBJSDK  C ON ( C.CODOBJETO = A.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITECT FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 1 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    D ON ( A.CODOBJETO = D.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITEMO FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 2 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    E ON ( A.CODOBJETO = E.CODOBJETO )"
      Strsql = Strsql & " LEFT  JOIN "
      Strsql = Strsql & "       (SELECT DISTINCT CODOBJETO, 'OK' AS PERMITECE FROM OBJSDKDET "
      Strsql = Strsql & "        WHERE TIPOSEG = 4 AND ( GRPUSR = 'USR-" & Session("Usuario") & "' OR GRPUSR IN (" & WKSPSDKGRP & ") ) "
      Strsql = Strsql & "        ) "
      Strsql = Strsql & "                    F ON ( A.CODOBJETO = E.CODOBJETO )"
      Strsql = Strsql & " where  "
      Strsql = Strsql & "  A.CODOBJETO = '" & sCODPROCESOIN & "' and A.TIPOBJETO = '" & sTIPPROCESOIN & "'  "


 ' Strsql = "Select * From OBJETOS  where CODOBJETO = " & sCODPROCESOIN
 ' Strsql = Strsql & " and tipobjeto = '" & sTIPPROCESOIN & "'  "

 ' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     WKPROCDESC = RS(0)


      IF RS("SEGCE") THEN

         IF (RS("PERMITECE") & "") <> "OK" THEN

            WKSPPERMCE = "N"

         ELSE

            WKSPPERMCE = "S"

         END IF


      END IF

      IF RS("SEGMO") THEN

         IF (RS("PERMITEMO") & "") <> "OK" THEN

            WKSPPERMMO = "N"

         ELSE

            WKSPPERMMO = "S"

         END IF


      END IF


      IF RS("SEGCT") THEN

         IF (RS("PERMITECT") & "") <> "OK" THEN

            WKSPPERMCT = "N"

         ELSE

            WKSPPERMCT = "S"
            WKSPPERMCE = "S"
            WKSPPERMMO = "S"

         END IF


      END IF



      IF (RS("SEGOWN") & "") = SESSION("USUARIO") THEN

         WKSPPERMCT = "S"
         WKSPPERMMO = "S"
         WKSPPERMCO = "S"
         WKSPPERMCE = "S"

      END IF




      sHref = "javascript:void(0);"

      sImage = "<font face='corflaticon' size=3 color='#FFFFFF'>&#0120;</font>"


       IF ( RS("SEGCT")  OR RS("SEGCO")  OR RS("SEGCE")  OR RS("SEGMO") ) THEN
          sImagePER = "<font face='corflaticon' size=3 color='#FFFFFF'>&#0207;</font>"
       ELSE
          sImagePER = "<font face='corflaticon' size=3 color='#FFFFFF'>&#0208;</font>"
       END IF




    IF  WKSPPERMCT = "S" THEN

      WKSPPERMISOS = "<a href='edobjsdk.asp?Tipobjeto=" & sTIPPROCESOIN & "&Codobjeto="
      WKSPPERMISOS = WKSPPERMISOS & sCODPROCESOIN & "&Menu=TEMAS&Sector=0&GLBSOLN=PRGSDK&MODALIDAD=PRG' >"
      WKSPPERMISOS = WKSPPERMISOS & sImagePER & "&nbsp;PERMISOS</a>"

      WKSPPERMOPC  = " style='background-color:#993333;color:#ffffff;cursor:pointer;' title='Permisos del proceso/subproceso.' "

    ELSE

      WKSPPERMISOS = ""
      WKSPPERMISOS = WKSPPERMISOS & sImagePER & "&nbsp;RESTRINGIDO"

      WKSPPERMOPC  = " style='background-color:#DDDDDD;color:#AAAAAA;'  "

    END IF



       Filan = ""
       Filan1 = ""
       Celda = ""
       Celda = Celda & ofv.GenCelda("","ut COLSPAN=2 style='background-color:#993333;color:#ffffff;' ","","","","","","PROPIEDADES","si")
       Celda = Celda & ofv.GenCelda("","aaa " & WKSPPERMOPC & "  ","center","","","","",WKSPPERMISOS,"si")
       Celda = Celda & ofv.GenCelda("","aaa style='background-color:#993333;color:#ffffff;cursor:pointer;' onclick=""document.all.DX4.style.display='none';return false;"" ","center","","","","",sImage,"si")
       FilaN = ofv.GenRow("","aaa bgcolor='#fefefe'","","","15","","",Celda,"si")


          t2 = "<a href='edobjsdk.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal>"
          t2 = t2 & sImagen4 & "</a>"



       Celda = ""
       sInput = "&nbsp;"
       Celda = Celda & ofv.GenCelda("","UT ","","10%","","","",sInput,"si")
       Celda = Celda & ofv.GenCelda("","UT ","","89%","","","",sInput,"si")
       Celda = Celda & ofv.GenCelda("","UT ","","2","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","15","","",Celda,"si")

       Strsql = "Select * From tipobjeto  where tipobjeto = '" & rs(2) & "' "
 '' response.write strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rs1.eof then
          WKTIPODESC = RS1(1)
       ELSE
          WKTIPODESC = "Sin Definir"
       END IF




       Celda = ""

       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","ID&nbsp;","si")
       sInput = sCODPROCESOIN
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       CELDA=""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","TIPO&nbsp;","si")
       sInput = WKTIPODESC
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","NOMBRE&nbsp;","si")
       sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"50","dt","")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","CODIGO&nbsp;","si")
       sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"50","dt","")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","SECUENCIA&nbsp;","si")
       sInput = ofv.GenerarInput(rs(13).name,rs(13),"text",sForm,"5","dt","")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

          WKESTADODESC = 4


       Strsql = "Select * From OBJPRTY  where CODOBJETO = " & sCODPROCESOIN
       Strsql = Strsql & " and tipobjeto = '" & sTIPPROCESOIN & "'  "
 '' response.write strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rs1.eof then
          WKFECCREDESC = RS1(6)
          WKFECBAJDESC = RS1(8)
          WKESTADODESC = RS1(9)



       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","ESTADO&nbsp;","si")
       sInput = GTESTADODSC(WKESTADODESC)
       Celda = Celda & ofv.GenCelda("","UT colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","CREACION&nbsp;","si")
       sInput = GTFECHADOR(WKFECCREDESC)
       Celda = Celda & ofv.GenCelda("","UT colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")


       IF RS1(9) = 4 THEN

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","BAJA&nbsp;","si")
       sInput = GTFECHADOR(WKFECBAJDESC)
       Celda = Celda & ofv.GenCelda("","UT colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       END IF



       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","VERSION&nbsp;","si")

       IF RS1(14) = "V" THEN
          sInput = rs1(15)
       ELSE
          sInput = "ORIGINAL"
       END IF
       Celda = Celda & ofv.GenCelda("","UT colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")


       ELSE

       END IF


       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","PRIORIDAD&nbsp;","si")
       sInput = ofv.GenerarInput(rs(19).name,rs(19),"text",sForm,"5","dtN","")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","LINEA NEGOCIOS&nbsp;","si")
       Sql1 = "Select OBJID,DESCRI From ORLINNEGOCIOS Order By DESCRI"
       sInput = ofv.GenerarCombo(sForm,rs(18).name,rs(18).value,Sql1,cn,0,1,"")
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

    if sCODNIVEL = 1 then

       scomK = "1" & chr(9) & "2"
       scomD = "MESES" & chr(9) & "AÑOS"

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","FRECUENCIA EVALUACION&nbsp;","si")
       sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),scomK,scomD," ") & ofv.GenerarInput(rs(10).name,rs(10),"text",sForm,"3","dtN","")
       Celda = Celda & ofv.GenCelda("","vt COLSPAN=2 ","","","","","",sInput,"si")

       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

    end if





       Celda = ""


             wkcodMODIF = "S"

          Strsql = "Select * From relaobj  where  "
          Strsql = Strsql & " codobjmaster = '" & rs(0) & "' and tipobjmaster = '" & rs(2) & "'  "
          set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsx2.eof then
             wkcodMODIF = "N"
          end if


'''wspTIPO

       scomK = ""
       scomD = ""

       select case wspTIPO


                  case "SP"


                       scomK = "SP" & chr(9) & "SPI" & chr(9) & "VN" & chr(9) & "CR"
                       scomD = "PROCESO" & chr(9) & "PROCESO INFORMATICO" & chr(9) & "EVENTO VINCULADO EMANAGER" & chr(9) & "EVENTO MANUAL"

                  case "SPI"

                       scomK = "SPI" & chr(9) & "AI"
                       scomD = "PROCESO INFORMATICO" & chr(9) & "ACTIVO INFORMATICO"

                  case "VN"

                       scomK = "VN"
                       scomD = "EVENTO VINCULADO EMANAGER"

                  case "CR"

                       scomK = "CR"
                       scomD = "EVENTO MANUAL"

                  case "AI"

                       scomK = "AI"
                       scomD = "ACTIVO INFORMATICO"

                  case else




       end select

     '  scomK = "SP" & chr(9) & "SPI" & chr(9) & "VN" & chr(9) & "CR" & chr(9) & "AI"
     '  scomD = "PROCESO" & chr(9) & "PROCESO INFORMATICO" & chr(9) & "EVENTO VINCULADO EMANAGER" & chr(9) & "EVENTO MANUAL" & chr(9) & "ACTIVO INFORMATICO"

       Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","CLASE&nbsp;","si")

       if  rs(15) > 0 OR wkcodMODIF = "N" then

           select case (RS(17) & "")


                  case "SP"

                        sInput = "PROCESO"

                  case "SPI"

                        sInput = "PROCESO INFORMATICO"

                  case "VN"

                        sInput = "EVENTO VINCULADO EMANAGER"

                  case "CR"

                        sInput = "EVENTO MANUAL"

                  case "AI"

                        sInput = "ACTIVO INFORMATICO"

                  case else

                        sInput = "SIN DETERMINAR"

           end select




       else

           sInput = ofv.gencomboFX(sForm,rs(17).name,rs(17),scomK,scomD," ")

       end if



       Celda = Celda & ofv.GenCelda("","vt colspan=2   ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

''''''''''''''''

       XFVONOBJ = "N"

    IF (RS(17) & "") = "VN" or (RS(17) & "") = "CR"  or (RS(17) & "") = "AI" THEN

      IF (RS(17) & "") = "VN" THEN

         sInstrumento = "<big>Sin documento asociado ........</big>"

         if  rs(15) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & rs(15)
           set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs1.eof then
              wtipoL   = clng(rs1(6))
              wtipOBJ  = rs1(7)
              wUrl     = rs1(8)
              wUrlDesc = rs1(9)
              wtipo    = rs1(10)


              if trim(wtipo) = "LRGVN" then

                 XFVONOBJ = "S"

                 Strsql = "Select xdescri From csGenObjetos where codobjeto = " & wtipoL
     ''      response.write strsql & "<br>"
                 set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
                 call ofv.cerrarconsulta(rsInstrumento)

              end if


          end if
          call ofv.cerrarconsulta(rs1)




         end if
         sInput = sInstrumento


         Celda = ""
         Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","EVENTO&nbsp;","si")

        IF WKSPPERMMO = "S" THEN

           sID = "EdGENRG.asp?"
           sID = sID & "ID=" & rs(0) & "&TABLA=OBJETOS&Codobjeto=" & sCODPROCESOIN
           sID = sID & "&Tipobjeto=" & sTIPPROCESOIN  & "&volver=" & VUELTA

          sinput2 = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput2 = sinput2  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
          sinput2 = sinput2  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
          sinput2 = sinput2  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
          sinput2 = sinput2  & "none" & chr(34) & ";' "
          sinput2 = sinput2  & "  href='" & sID & "'  "
          sinput2 = sinput2  & "  TITLE='Asociar/cambiar documento operativo.'  >" & sInput & "</a>"


           IF WKESTADODESC = 4 THEN
              sinput2 = sInput

           END IF

        ELSE
            sinput2 = sInput

        END IF

         Celda = Celda & ofv.GenCelda("","ut colspan=2 ","","","","","",sinput2,"si")
         FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")


  elseIF (RS(17) & "") = "CR" OR (RS(17) & "") = "AI" THEN

           'sInstrumento = "<big>Subproceso no habilitado para acceder a las actividades.qqq</big>"
           sInstrumento= "<big>Habilitar Subproceso</big>"
           IF (RS(17) & "") = "CR"  THEN
              'sInstrumento = "<big>Subproceso no habilitado para acceder a las actividades.gggg</big>"
               sInstrumento = "<big>Habilitar Subproceso</big>"
           else
              'sInstrumento = "<big>Subproceso no habilitado para acceder a los activos informaticos.</big>"
               sInstrumento = "<big>Habilitar Subproceso</big>"
           end if


         if  rs(15) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & rs(15)
           set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs1.eof then
              wtipoL   = clng(rs1(6))
              wtipOBJ  = rs1(7)
              wUrl     = rs1(8)
              wUrlDesc = rs1(9)
              wtipo    = rs1(10)


              if trim(wtipo) = "LRGCR" then

                 XFVONOBJ = "S"



                 IF (RS(17) & "") = "CR"  THEN
                    sInstrumento = "<big>El Subproceso esta habilitado para acceder a las actividades.</big>"
                 else
                    sInstrumento = "<big>El Subproceso esta habilitado para acceder a los activos informaticos.</big>"
                 end if


              end if


          end if
          call ofv.cerrarconsulta(rs1)




         end if
         sInput = sInstrumento

       Celda = ""
     '  Celda = Celda & ofv.GenCelda("","UTI style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","EVENTO&nbsp;","si")


         if XFVONOBJ = "S" then

            sinput2 = sInput

         else

          IF WKSPPERMMO = "S" THEN


            sinput2 = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
            sinput2 = sinput2  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
            sinput2 = sinput2  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "none" & chr(34) & ";' "
            sinput2 = sinput2  & "  href='javascript:void(0);' onclick="
           IF (RS(17) & "") = "CR"  THEN
            sinput2 = sinput2  & chr(34) & "EditarCR('NEW','" & sCODPROCESOIN & "','" & sTIPPROCESOIN & "');" & chr(34)
           else
            sinput2 = sinput2  & chr(34) & "EditarAI('NEW','" & sCODPROCESOIN & "','" & sTIPPROCESOIN & "');" & chr(34)
           end if
            sinput2 = sinput2  & "  TITLE='Habilitar el subproceso.'  >" & sInput & "</a>"

''''<b>&nbsp;" & sImagenI3 & "...&nbsp;</b>

         ELSE
              sinput2 = sInput

         END IF

       end if



       IF WKESTADODESC = 4 THEN
          sinput2 = sInput

       END IF



       Celda = Celda & ofv.GenCelda("","ut colspan=3 style='background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#993333;' ","center","","","","",sinput2,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

  end if








'''''''''''''''''''''''''''''''''''''''''''''''
'' analiza cambio de version
'''''''''''''''''''''''''''''''''''''''''''''''


       XFNOVEPROC = ""



       IF  XFORPERIODOST = "OK"  AND WKESTADODESC <> 4 THEN


           if  XFVONOBJ = "S" then


            IF (RS(17) & "") = "VN" THEN



               strsql = "SELECT B.FECBAJA, B.ESTADO FROM OBJETOS A INNER JOIN OBJPRTY B   "
               strsql = strsql & " ON A.CODOBJETO = B.CODOBJETO "
               strsql = strsql & " WHERE A.CODOBJETO = 0" & wtipoL
       '''     response.write strsql & "<br>"
               set rsN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
               if NOT rsN.eof then

                IF RSN(1) = 4 THEN

                  IF (CSTR(RSN(0) & "") > CSTR(XFORPERIODOFD & "") AND CSTR(RSN(0) & "") < CSTR(XFORPERIODOFH & "")) THEN

                   XFNOVEPROC = "<FONT COLOR='#FF0000' SIZE=2 >**</FONT>" & "&nbsp;" &  XFORPERIODOFD & "&nbsp;" &  XFORPERIODOFH & "&nbsp;" &  XFORPERIODOst & "&nbsp;" &  cstr(RSN(0) & "")


                  IF WKSPPERMCE = "S" THEN

                   sID = "RGREEMPLAZASP.asp?"
                   sID = sID & "Codigo=" & sCODPROCESOIN
                   sID = sID & "&Tipo=" & sTIPPROCESOIN
                   sID = sID & "&CodESQS=" & scodmaster
                   sID = sID & "&TipESQS=" & sTIPMASTER
                   sID = sID & "&volver=" & VUELTA



                   sinput2 = "<a class=NT style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "none" & chr(34) & ";' "
                   sinput2 = sinput2  & "  href='" & sID & "'  "
                   sinput2 = sinput2  & "  TITLE='GENERAR NUEVA VERSION ACTUALIZADA.'  target='fesquemas' >EL DOCUMENTO VINCULADO A SIDO REEMPLAZADO&nbsp;</a>"

                  ELSE
                   sinput2 = "EL DOCUMENTO VINCULADO A SIDO REEMPLAZADO&nbsp;"

                  END IF


                   Celda = ""
                   Celda = Celda & ofv.GenCelda("","UTI COLSPAN=3 style='background-color:#990000;color:#ffffff;' ","center","","","","",sinput2,"si")
                   FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")



                 END IF

               END IF

             END IF

         elseIF (RS(17) & "") = "CR" or (RS(17) & "") = "AI" THEN



                  IF WKSPPERMCE = "S" THEN


                   sID = "RGREEMPLAZASP.asp?"
                   sID = sID & "opcion=RGCRNV"
                   sID = sID & "&Codigo="  & sCODPROCESOIN
                   sID = sID & "&Tipo="    & sTIPPROCESOIN
                   sID = sID & "&CodESQS=" & scodmaster
                   sID = sID & "&TipESQS=" & sTIPMASTER
                   sID = sID & "&volver="  & VUELTA



                   sinput2 = "<a class=NT style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
                   sinput2 = sinput2  & "none" & chr(34) & ";' "
                   sinput2 = sinput2  & "  href='" & sID & "'  "
                   sinput2 = sinput2  & "  TITLE='GENERAR NUEVA VERSION ACTUALIZADA.'  target='fesquemas' >GENERAR NUEVA VERSION DEL SUBPROCESO&nbsp;</a>"




                   Celda = ""
                   Celda = Celda & ofv.GenCelda("","UTI COLSPAN=3 style='background-color:#990000;color:#ffffff;' ","center","","","","",sinput2,"si")
                   FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

                  END IF


         else



         end if

           END IF


       END IF






'''''''''''''''''''''''''''''''''''''''''''''''
'' fin analiza cambio de version
'''''''''''''''''''''''''''''''''''''''''''''''


  END IF


'''''''''''''''''

     IF WKESTADODESC = 4 THEN
       sAccion = ""

     ELSE


      IF WKSPPERMMO = "S" THEN

       sAccion = "asp/actualizar.asp?fname=" & rs(0)
       sAccion = sAccion & "&tabla=OBJETOS&volver=" & vuelta

     ELSE
       sAccion = ""
     END IF

     END IF
       Formulario = ofv.genform(sForm,"",sAccion,"","post","",Filan,"si")


       Filan = Filan1 & ofv.GenTabla("","aaa ","","98%","","0","","0","1","","",Formulario,"si")

       Filan = "<CENTER>" & Filan & "</CENTER>"



  ELSE

       Filan = ""

  end if


      call ofv.cerrarconsulta(rs)
      call ofv.cerrarconn(cn)

      Fscript = "<SCRIPT language='javascript' >document.all.DX4.style.display='block';</script>" & Chr(13)


 ELSE

      Filan   = "&nbsp;"
      Fscript = ""

 END IF

'''''''''''''''''''''''''''''''''''''''''''''

  Celda = "" & Chr(13)


  Celda = Celda & "<div id='DX4' style='margin-top:1;margin-left:1;position:absolute;left:500px;top:140px;display:none;"
  Celda = Celda & "width:500;height:400;overflow:auto;border-style:OUTSET;border-width:2;"
  Celda = Celda & "background-color:#F9F9F9;"
  Celda = Celda & "' >" & Chr(13)



  Celda = Celda & Filan & Chr(13)


  Celda = Celda & "</div>" & Chr(13)

  sHTML = sHTML & Celda & Chr(13)

'''  sHTML = sHTML & WKDIVSECUEN & Chr(13)

  sHTML = sHTML & Fscript & Chr(13)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no","","no","")

  sHTML = sHTML & GTCOBSRSK()

  sHTML = sHTML & "</body></html>" '& SESSION("STRSQLX")
 END IF

  response.write  sHTML


function verdata()
dim sst

sst="<table border><tr>"
sst=sst&"<td>AEstado=" & AEstado &"</td></tr>"
sst=sst&"<td>AConsultar=" & AConsultar &"</td></tr>"
sst=sst&"<td>AEliminar=" & AEliminar &"</td></tr>"
sst=sst&"<td>AAgregar=" & AAgregar &"</td></tr>"
sst=sst&"<td>grupos=" &replace(session("SegGrpLst") & Session("SegGrp")&"",chr(9),"<br>") &"</td></tr>"
sst=sst& "</table>"
verdata=sst

end function

 FUNCTION GENPROCSECUEN()

    GENPROCSECUEN = ""

    Strsql = "Select * From CSRELPATHDESCRI  where COBJ = " & scodmaster
 '' response.write strsql & "<br>"
    set Wrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT Wrs.eof then

       WKITOPE = 8

       FOR WKI=0 TO 7

           WKIPROC = WRS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKITOPE = WKI
              EXIT FOR
           END IF


       NEXT

       WKICCOL = 7 - WKITOPE

       WKINCOL = 0

       FilaW = ""

       FOR WKI=WKITOPE TO 7

           Celda = ""
           Celda2 = ""

               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")

               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")

           FOR WKIC=1 TO WKINCOL

               WKISTYLO = ""

               IF WKIC = WKINCOL THEN

                  WKISTYLO = "border-LEFT-style:SOLID;border-LEFT-width:1;Border-LEFT-color:#F7AA22;"
                  WKISTYLO = WKISTYLO & "border-BOTTOM-style:SOLID;border-BOTTOM-width:1;Border-BOTTOM-color:#F7AA22;"

               END IF

               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;" & WKISTYLO & "' ","","","","","","&nbsp;","si")

               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")


           NEXT

           Celda = Celda & ofv.GenCelda("","aaa colspan=4 ROWSPAN=2 bgcolor='#ffffff' style='font-family:tahoma;font-size:7pt;color:#888888;border-style:solid;border-width:1;border-color:#993333;'  ","","","","","","&nbsp;" & WRS(WKI),"si")


           FOR WKIC=WKINCOL TO WKICCOL

               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")

               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
               Celda2 = Celda2 & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")


           NEXT

           FilaW = FilaW & ofv.GenRow("","","","","5","","",Celda,"si")
           FilaW = FilaW & ofv.GenRow("","","","","5","","",Celda2,"si")

           WKINCOL = WKINCOL + 1

           WKISTYLO = ""

           IF WKINCOL < (WKICCOL + 1) THEN

              WKISTYLO = "border-LEFT-style:SOLID;border-LEFT-width:1;Border-LEFT-color:#F7AA22;"

           END IF


           Celda = ""
           Celda = Celda & ofv.GenCelda("","aaa colspan=" & ((WKINCOL * 2) ) & " style='font-family:tahoma;font-size:4pt;' ","","","","","","&nbsp;","si")
           Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;' ","","","","","","&nbsp;","si")
           Celda = Celda & ofv.GenCelda("","aaa style='font-family:tahoma;font-size:6pt;" & WKISTYLO & "' ","","","","","","&nbsp;","si")
           Celda = Celda & ofv.GenCelda("","aaa colspan=" & (((WKICCOL - WKINCOL) * 2) + 6) & " style='font-family:tahoma;font-size:4pt;' ","","","","","","&nbsp;","si")
           FilaW = FilaW & ofv.GenRow("","","","","5","","",Celda,"si")




       NEXT


       IF WKITOPE < 8 THEN

           Celda = ""
           Celda = Celda & ofv.GenCelda("","aaa colspan=" & ((WKICCOL * 2) + 8) & " style='font-family:tahoma;font-size:4pt;' ","","","","","","&nbsp;","si")
           FilaW = ofv.GenRow("","","","","5","","",Celda,"si") & FilaW

          FilaW  = ofv.GenTabla("","aaa ","","100%","","0","","-1","-1","","",FilaW,"si")

          sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0120;</font>"

          wkproccod3 = "<a class=nt onclick=""document.all.DX5.style.display='none';"" href='javascript:void(0);' "
          wkproccod3 = wkproccod3 & " ><b>"
          wkproccod3 = wkproccod3 & sImage
          wkproccod3 = wkproccod3 & "</b></a>"


          Celda = ""
          Celda = Celda & ofv.GenCelda("","uti ","","98%","","","","ENCADENAMIENTO DEL PROCESO","si")
          Celda = Celda & ofv.GenCelda("","ut ","","2%","","","",wkproccod3,"si")
          FilaW2 = ofv.GenRow("","","","","","","",Celda,"si")
          FilaW  = ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' ","","100%","","0","","0","1","","",FilaW2,"si") & FilaW

          FilaW  = "<CENTER>" & FilaW & "</CENTER>"



          Celda = "" & Chr(13)


          Celda = Celda & "<div id='DX5' style='z-index:1;margin-top:1;margin-left:1;position:absolute;left:400px;top:140px;display:none;"
          Celda = Celda & "width:400;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
          Celda = Celda & "background-color:#F9F9F9;"
          Celda = Celda & "' >" & Chr(13)



          Celda = Celda & FilaW & Chr(13)


          Celda = Celda & "</div>" & Chr(13)


          sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:1;' >4</font>"

          wkproccod2 = "<a class=nt onclick=""document.all.DX5.style.display='block';"" href='javascript:void(0);' "
          wkproccod2 = wkproccod2 & " title='ENCADENAMIENTO DEL PROCESO.' >"
          wkproccod2 = wkproccod2 & sImage
          wkproccod2 = wkproccod2 & "</a>"


          GENPROCSECUEN = Celda

       END IF



    END IF


 END FUNCTION

 FUNCTION GENPROCSECUENEX()

    GENPROCSECUENEX = ""

    Strsql = "Select * From CSRELPATHDESCRI  where COBJ = " & scodmaster
 '' response.write strsql & "<br>"
    set Wrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT Wrs.eof then

       WKITOPE = 8

       FOR WKI=0 TO 7

           WKIPROC = WRS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKITOPE = WKI
              EXIT FOR
           END IF


       NEXT

       WKICCOL = 7 - WKITOPE

       WKINCOL = 0

       FilaW = ""

           Celda = ""
           Celda2 = ""

       FOR WKI=WKITOPE TO 7




           Celda = Celda & Celda2 & WRS(WKI)

           Celda2 = "<b><font color='#00aa00' size=2 >&#0124;</font></b>"



       NEXT


       IF WKITOPE < 8 THEN

          GENPROCSECUENEX = Celda & Celda2

       END IF



    END IF


 END FUNCTION

function ObtenerAtributos(sForm,  sNodo,  sCampo,  sValor)

       ' On Error GoTo ObtenerAtributosErr

     Dim strsql,  atrok,   sSQL,  xcount , XMFI , XMFSEP
     Dim XMFTGRP



     Dim XTabPer(4)

     xcount = 0


     AEstado = True
     AConsultar = True
     AEliminar = True
     AModificar = True
     AAgregar = True

     atrok = "N"
     strsql = "Select Alias,usrEstado,usragregar,usrmodificar,usreliminar,usrconsultar "
     strsql = strsql & " From atrmXusuario Where "
    If sNodo <> "" Then
        strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
    Else
        strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
    End If
    strsql = strsql & " And usuario = '" & Session("segperfil") & "' And "
    strsql = strsql & "  Aplicacion = '" & Session("CorApli") & "'"
 ''   response.write strsql & "<br>"
    Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    If rs.EOF Then
        atrok = "O"
  '  response.write "aaa<br>"

    Else
        atrok = "S"
        AEstado = rs(1)
        AConsultar = rs(5)
        AEliminar = rs(4)
        AModificar = rs(3)
        AAgregar = rs(2)

  '  response.write "bbb<br>"


    End If
    Call ofv.CerrarConsulta(rs)

  '  response.write "atrok " & atrok & "<br>"

    If atrok = "O" Then
        strsql = "Select Alias,grpEstado,grpagregar,grpmodificar,grpeliminar,grpconsultar "
        strsql = strsql & " From atrmXgrupo Where  "
        If sNodo <> "" Then
            strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
        Else
            strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
        End If
        strsql = strsql & " And Aplicacion = '" & Session("CorApli") & "'"

     If Session("DELLOG") = "S" Then

          XMFTGRP = Split(Session("SegGrpLst"), Chr(9))
          xcount = UBound(XMFTGRP) + 1
          XMFSEP = ""
          For XMFI = 0 To 4
              XTabPer(XMFI) = xcount
          Next

          strsql = strsql & " and grupo in ("

          For XMFI = 0 To UBound(XMFTGRP)
              strsql = strsql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
              XMFSEP = ","
          Next
          strsql = strsql & ") "
 '  response.write strsql & "<br>"
          Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          If rs.EOF Then
             atrok = "O"
   ' response.write "ccc<br>"

          Else
   ' response.write "ddd<br>"

             atrok = "S"

             If Not rs(5) Then XTabPer(1) = XTabPer(1) - 1 ''' consultar
             If Not rs(4) Then XTabPer(2) = XTabPer(2) - 1 ''' eliminar
             If Not rs(3) Then XTabPer(3) = XTabPer(3) - 1 ''' modificar
             If Not rs(2) Then XTabPer(4) = XTabPer(4) - 1 ''' agregar

          End If

          If atrok = "S" Then

               do while not rs.eof

                 If XTabPer(1) = 0 Then AConsultar = False
                 If XTabPer(2) = 0 Then AEliminar = False
                 If XTabPer(3) = 0 Then AModificar = False
                 If XTabPer(4) = 0 Then AAgregar = False

                 rs.movenext
                loop

          End If


      Else
        strsql = strsql & " And grupo = '" & Session("seggrp") & "'  "
  ''  response.write strsql & "<br>"
        Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        If rs.EOF Then
            atrok = "O"

   ' response.write "eee<br>"
        Else

   ' response.write "fff<br>"
           atrok = "S"
           '  If Not rs(1) Then AEstado = False ''' consultar
           '  If Not rs(5) Then AConsultar = False ''' consultar
           '  If Not rs(4) Then AEliminar  = False ''' eliminar
           '  If Not rs(3) Then AModificar = False ''' modificar
           '  If Not rs(2) Then AAgregar   = False ''' agregar

           '  If Not rs(1) Then AEstado = 0 ''' consultar
           '  If Not rs(5) Then AConsultar = 0 ''' consultar
           '  If Not rs(4) Then AEliminar  = 0 ''' eliminar
           '  If Not rs(3) Then AModificar = 0 ''' modificar
           '  If Not rs(2) Then AAgregar   = 0 ''' agregar

           '  If rs(1) = 0 Then AEstado = 0 ''' consultar
           '  If rs(5) = 0 Then AConsultar = 0 ''' consultar
           '  If rs(4) = 0 Then AEliminar  = 0 ''' eliminar
           '  If rs(3) = 0 Then AModificar = 0 ''' modificar
           '  If rs(2) = 0 Then AAgregar   = 0 ''' agregar


  '  response.write AConsultar & "<br>"
  '  response.write AEliminar & "<br>"
  '  response.write AModificar & "<br>"
   ' response.write AAgregar & "<br>"


             AEstado = rs(1)
            AConsultar = rs(5)
            AEliminar = rs(4)
            AModificar = rs(3)
            AAgregar = rs(2)
        End If

       End If

        Call ofv.CerrarConsulta(rs)
    End If


    ObtenerAtributos = "ok"


End function



%>

