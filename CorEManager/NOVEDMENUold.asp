<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
   uv = ofv.uv
   if uv = "NO" then ok = ofv.ValidarUsuario
   sHTML = ok
else
    response.expires=0

    session("ModMain") = "CL"
    wmnnoved = "NOVED" & WKAPLIC

    session(wmnnoved) = "SI"
    aplcheck()

    wspviewopc = 2

    stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
    stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"


    ReDim Preserve bConsultar(10)

''''response.write "<br><br><br><br>" & Session("Form") & "<br>"

 Novedform = "CLIMN16317206150930885546227"

 if (session("UN") & "") = "" then


     XXXX = NEWEMPRE("NO")


     SEMPRE = ATREMP(0,0)
     session("UN") = SEMPRE & ""

     if (session("UN") & "") = "" then session("UN") = 1


 end if


    set cn = ofv.conectar(ofv.strconn2)
    set cn0 = ofv.conectar(ofv.strconn0)

wmnfuncion = ""

atmensajes = 0
dim atitsecciones(50,2)
ttitsecc = 0

dim atit(50)
dim awmnasp(50)
wmntope = 10
wmntopeX = 10
WMNINITX = 2

 FOR TT=1 TO wmntope
     atit(tt) = "NO"
 NEXT

 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN
 '   wmntopeX = 5
 '   WMNINITX = 4
 END IF


titc = 0
Strsql = "SELECT * FROM novedtit "
set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rsx.eof then

   for tt = WMNINITX to wmntopeX
       atit(tt) = rsx(tt)
       if rsx(tt) <> "NO" then
          OFV.ObtenerAtributosNOdo Novedform, "", "TIPO", CSTR(tt)
          bConsultar(TT) = OFV.AConsultar

          IF TT = 10 THEN

             IF bConsultar(TT) THEN
                session("SMPLMENU") = "N"
                atit(tt) = "NO"

             else
                if session("SMPLMENU") <> "S" then
                   session("SMPLMENU") = "S"
                   wmnfuncion = "<script language='javascript' > hdemenu(); </script> "
                end if
                titc = titc + 1



             end if


          ELSE

             IF bConsultar(TT) THEN
                 titc = titc + 1

             else

                 atit(tt) = "NO"

             end if

          END IF

       END IF
   next

   if  atit(10) <> "NO" then

       ii = 10
       xxxx = wmnbuscamenu("CLIS150172171122331051096140")
       wmntope = ii

   end if

end if
call ofv.cerrarconsulta(rsx)

menuSel = request.querystring("menu")
menu    = menuSel

if menu = "" then

   if   atit(3) <> "NO" then
        'menu = 3
        if WKAPLIC = "OR" then
           menu = 4
           menuSel = 4
        else
           menu=3
           menuSel = 3
        end if


   else
        menu = 0
        menuSel = 0

        for iix = 1 to 10
            if   atit(iix) <> "NO" then
                 menu = iix
                 menuSel = iix
                 exit for
            end if
        next
   end if

end if

menu    = cint(menu)

  if session("FormN") = "" then session("FormN") = "NOVEDADES"

  if session("SMPLMENU") = "S" then session("FormN") = "MENU PRINCIPAL"

  sHTML = ofv.FormHeader(session("FormN"))

 '' sHTML = ofv.menuheader("","novedades ")

    xs = ""

    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

      s =  "<script language=JavaScript>" & Chr(13)
      s = s & "function AbrirVentana(Datosx) {" & Chr(13)
      s = s & " myWin= open(Datosx"
      s = s & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
      s = s & "menubar=no,toolbar=no,height=570,width=790,status=yes,resizable=yes" & Chr(34)
      s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
      s = s & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs & s & wmnfuncion



  TABLA1 = ""
  TABLAB = ""
  TABLAN = ""
  TABLAB1 = ""




nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

   dFechai  = clng(date()) - nFecini
   XTrColor = "#FFFFFF"

   sglmar   = "<font class=btiW >*</font>"
   sglman   = "<font class=btiW >1</font>"
   sglprf   = "<font class=btic >&#0128;</font>"
   sgllnk   = "<font class=btic >&#0126;</font>"
   sgladj   = "<font class=btic >&#0204;</font>"


   dim ANTabla(500,25)


if session("WEBDESK") = "S" then


'''''''''''''' buscador por texto     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  TABLAB = TablaB & "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:NONE;width:98%;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"

'''  sInputT  = ofv.GenerarInput("textot","","agregar","","80","dt style='border-style:solid;border-width:1;border-color:#666666;' ","")

  if session("APLMDCLI") ="BGAL" THEN
      sInputT  = "<input type='text' class='dt' name='BUSCAR' id='textot' size='40' style='border-style:solid;border-width:1;border-color:#666666;' >"
  ELSE
    sInputT  = "<input type='text' class='dt' name='textot' id='textot' size='40' style='border-style:solid;border-width:1;border-color:#666666;' >"
  END IF

    celda  = ""
    Celda      = Celda & ofv.GenCelda("","UT  ","right","30%","","","",sInputT,"si")

    sInput   = ofv.GenerarInput("","Buscar","submit","","5","at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;' ","")
    Celda    = Celda & ofv.GenCelda("","vt","","70%","","","",sInput,"si")

    formx    = ofv.GenRow("","","","","","","",Celda,"si")

    sTARGET  = ""


    'sAccion  = "consBuscadorSet.asp"
      'sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")

    if session("APLMDCLI") ="BGAL" THEN
     sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOR"
     sTARGET="_new"
    ELSE
     sAccion  = "consBuscadorSet.asp"
    END IF

    FILA     = ofv.GenForm("Subconsulta","",sAccion,STARGET,"post","",formx,"si")

  TABLAB = TablaB & ofv.GenTabla("","aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ","","98%","","0","","-2","0","","",Fila,"si")

  TABLAB = TablaB & "</DIV>" & chr(13)



'''''''''''''' indice buscador normativo '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  TABLA1 = Tabla1 & "<div id='DTTEMA' style='float: left;width:98%; " & chr(13)
  TABLA1 = Tabla1 & "border-style:solid;border-width:2;border-color:#ffffff;' >"

  TABLA1 = Tabla1 & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
  TABLA1 = Tabla1 & "<tr HEIGHT=25  >"
  xsclase=" CLASS=iett "
  TABLA1 = Tabla1 & "<td "& xsclase &"  >" & Capitalize("INDICE NORMATIVO") & "</TD>"
  TABLA1 = Tabla1 & "</TR>"
  TABLA1 = Tabla1 & "</table>" & chr(13)

  TABLA1 = Tabla1 & "<div id='DTTEMAINT' style='overflow:auto;height:400px;" & WStyl & ";"
  TABLA1 = Tabla1 & "border-style:solid;border-width:1;border-color:#dddddd;' >" & chr(13)

  WSPFRMN = "DTTEMAIT"
  'WSPFRMA = "consdocumindex.asp"
  WSPFRMA = "CONSTEMA.asp?FORMATO=ABR"

  WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"

  WSPDTTSCP = ""
  WSPDTTSCP = WSPDTTSCP &  "<script FOR=window EVENT=onload  >"
 '' WSPDTTSCP = WSPDTTSCP &  " alert('ENTRO: ');"
  '  WSPDTTSCP = WSPDTTSCP &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
  WSPDTTSCP = WSPDTTSCP &  "document." & WSPFRMN & ".submit();"
'''  WSPDTTSCP = WSPDTTSCP &  "document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";"
  WSPDTTSCP = WSPDTTSCP &  "</script>"


'TABLA1 = Tabla1 & "<script src='../coremanager/consdocumindex.asp?sid=1&target=_self' ></script>"


  TABLA1 = Tabla1 & "</DIV>"
  TABLA1 = Tabla1 & "</DIV>" & WSPFORME & chr(13)


  XTABLAN = ""




   XTrColor = "#FFFFFF"

   XCANT = 0



             for ii=2 to wmntope
                 if  atit(ii) <> "NO" then
                     XCANT = XCANT + 1

                     TNTabla = 0
                     INTabla = 0
                     xnohay1 = "S"
                     xnohay2 = "S"
                     SaleSel = "N"

                    if  ii = 10 then
                         stilo  = stilo2
                    else

                        xxxx    = noveselec(II)

                        TNTabla = INTabla

                        XSOPTABN    = noveformat(II,menu)

                        TABLAN = TablaN &  XSOPTABN

                    end if

                 end if
            next


         if ttitsecc > 0 then


            XTABLAN = ""
           ' xsclase = " CLASS=VTI STYLE='background-color:#6666ff;color:#ffffff;font-size:9pt;' "

            xsclase=" class=iett "
            XTABLAN = XTablaN & "<div id='DTNOVSEC' style='float: left; "
            XTABLAN = XTablaN & "border-style:solid;border-width:2;border-color:#ffffff;' >"


            XTABLAN = XTablaN &  "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
            XTABLAN = XTablaN &  "<tr HEIGHT=25  >"
            XTABLAN = XTablaN &  "<td " & xsclase & "   >" & Capitalize("secciones") & "</TD>"
            XTABLAN = XTablaN &  "</TR>"
            XTABLAN = XTablaN &  "</table>"

            XTABLAN = XTablaN & "<div id='DTNOVINTSEC' style='"
            XTABLAN = XTablaN & "border-style:solid;border-width:1;border-color:#dddddd;' >"

            XTABLAN = XTablaN &  "<table width=99% cellspacing=0 cellpadding=0  >"

            xxfondo = ""


             for ii=1 to ttitsecc



                         XTABLAN = XTablaN &  "<tr HEIGHT='25'  " & xxfondo & " >"
                         XTABLAN = XTablaN &  "<td class=ut width=1% >&nbsp;<B><font color='#f7aa22' size=2 >&#0149;</font></B></TD>"
                         XTABLAN = XTablaN &  "<td class=ut STYLE='color:#000000;font-size:8pt;font-weight:500;' >&nbsp;"
                         XTABLAN = XTablaN &  ""
                         XTABLAN = XTablaN &  atitsecciones(II,0) & "&nbsp;(" & atitsecciones(II,1) & ")" & ""
                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                         if xxfondo = "" then
                             xxfondo = " bgcolor='#ddccff' "
                         else
                             xxfondo = ""
                         end if



            next



            XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
            XTABLAN = XTablaN &  "<td class=UT colspan=2  >&nbsp;</TD>"
            XTABLAN = XTablaN &  "</TR>"

            IF atmensajes > 0 THEN

                         XTABLAN = XTablaN &  "<tr HEIGHT='20' bgcolor='#f0eedd' >"
                         XTABLAN = XTablaN &  "<td class=vTi colspan=2 STYLE='color:#000000;font-size:8pt;"
                         XTABLAN = XTablaN &  "border-style:solid;border-width:1;border-color:#666666;' >"
                         XTABLAN = XTablaN &  Capitalize("mensajes") & "</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                         XTABLAN = XTablaN &  "<tr HEIGHT='25' >"
                         XTABLAN = XTablaN &  "<td class=ut width=1% >&nbsp;<B><font color='#990000' size=3 >!</font></B></TD>"
                         XTABLAN = XTablaN &  "<td class=ut STYLE='color:#000000;' >&nbsp;"
                         XTABLAN = XTablaN &  "<font color=#666666 style='font-size:8pt;font-weight:700;' >"
                         XTABLAN = XTablaN &  "Ud. tiene&nbsp;" & atmensajes & "&nbsp;mensajes&nbsp;pendientes." & "</font>"
                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"

            END IF


            XTABLAN = XTablaN &  "</table>"
            XTABLAN = XTablaN & "</DIV>"
            XTABLAN = XTablaN & "</DIV>"

            TABLAB1 = XTABLAN

         end if





else

   TNTabla = 0
   INTabla = 0
   xnohay1 = "S"
   xnohay2 = "S"
   SaleSel = "N"


   xxxx = noveselec(menu)

   TNTabla = INTabla


   TABLAN = ""
   TABLAN = TablaN & "<DIV id='DTNOVE' STYLE='display:BLOCK;z-index:1;margin-top:1;margin-left:1;overflow:auto;'>"


   TABLAN = TablaN &  "<table  width=100% cellspacing=0 cellpadding=2 class='HX1CL'  style='border-style:solid;border-width:1;border-color:#ffffff;"
   TABLAN = TablaN & "background-color:#cccccc;"
   TABLAN = TablaN & "' >"

   xnewmenu = ""
   xoldmenu = ""

   TABLAN = TablaN &  "<tr >"
   TABLAN = TablaN &  "<td class='UT'  >"

     TABLAN = TablaN &  "<table  cellspacing=1 cellpadding=3  >"

             for ii=2 to wmntope
                 if  atit(ii) <> "NO" then

                    if  ii = 10 then
                         stilo  = stilo2
                    else

                        if  ii = menu  then
                            xsclase = "ttr STYLE='border-style:inset;border-color:" & Session("CMFPNTITFT") & ";border-width:1;padding-left:5px;padding-right:5px;" & XspEx("#f7aa22","#dddddd",0) & ";' "
                        else
                            xsclase = "ttr STYLE='border-right-style:solid;border-right-color:" & Session("CMFPNTITFT") & ";border-right-width:1;padding-left:5px;padding-right:5px;' "
                        end if

                        xmDato = REPLACE(atit(ii)," ","&nbsp;")

                        if ii < 11 then
                           xmhref    = "<a class='xts' href='NOVEDMENU.asp?menu=" & ii & "' >" & xmDato & "</a>"
                           xoldmenu  = xoldmenu &  "<td class=" & xsclase & " align=center  >"
                           xoldmenu  = xoldmenu &  xmhref & "</TD>"
                        else
                           xmhref    = "<a class='xts' href='" & awmnasp(ii) & "' >" & xmDato & "</a>"
                           xnewmenu  = xnewmenu &  "<td class=" & xsclase & " align=center  >"
                           xnewmenu  = xnewmenu &  xmhref & "</TD>"

                        end if



                    end if

                 end if
            next

     TABLAN = TablaN &  xnewmenu & xoldmenu
     TABLAN = TablaN &  "</table>"

   TABLAN = TablaN &  "</TD>"
   TABLAN = TablaN &  "</TR>"
   TABLAN = TablaN &  "</table>"

'''''''''''''''''''''''''''''
   TABLAN = TablaN &  "<center>"

   TABLAN = TablaN & "<div id='WMN0' style='overflow:auto;height:50%;" & WStyl & "' >"

   if  xnohay1 = "S" and xnohay2 = "S" then

       TABLAN = TablaN &  "<table  margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"
       TABLAN = TablaN &  "<tr  height=200 >"
       TABLAN = TablaN &  "<td class=ttt width=100% align=center valign=middle ><big>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(menu) & "&nbsp;</TD>"
       TABLAN = TablaN &  "</TR>"

   else
       TABLAN = TablaN &  "<table bgcolor=#eeeeee margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"

       for INTabla=1 to TNTabla

          xtmcod   = ANTabla(INTabla,0)
          xtipom   = ANTabla(INTabla,1)
          xtpri    = ANTabla(INTabla,2)
          xtfvig   = ANTabla(INTabla,3)
          xthvig   = ANTabla(INTabla,4)
          xtsadj   = ANTabla(INTabla,5)
          xtslnk   = ANTabla(INTabla,6)
          xfromusr = ANTabla(INTabla,7)
          xtmsg    = ANTabla(INTabla,8)
          xtmsgimg    = ANTabla(INTabla,9)
          xtmsglnk    = ANTabla(INTabla,10)

          xtmsglnk    = replace(xtmsglnk,"XXMSGXX",xtmsg)



          select case xtipom
                 case 1
                      xtipomes = "GENERAL"
                 case 2
                      xtipomes = "PARTICULAR"
                 case 3
                      xtipomes = "PUBLIC"
                 case 4
                      xtipomes = "AUTORIZ"
                 case 5
                      xtipomes = "WORKFLOW"


          end select

          xfclas    = "dt"
          sImage    = "&nbsp;"

          select case xtpri
                 case 0

                      xfclas = "btbn"
                      sglmod = 5
                      sgltit = "Prioridad Alta."

                case 1

                     xfclas = "btba"
                     sglmod = 4
                     sgltit = "Prioridad Normal."

                case 2

                     xfclas = "btbs"
                     sglmod = 6
                     sgltit = "Prioridad baja."


          end select

          sglmod     = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
          sImage     = ofv.GenLink("",xfclas & " title='" & sgltit & "' ","","","","",sglmod,"si")

          if  xtfvig = dFechai then
              XMARQ  = sglmar
          else
              XMARQ  = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
          end if

          IF  xtipom = 5 THEN
              xtfechavig = xthvig
          ELSE
              xtfechavig = fechaArg(cdate(xtfvig + clng(cdate("01-01-1900")))) & "&nbsp;&nbsp;" & xthvig
          END IF

          Xnlnk      = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
          Xnadj      = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

          if xtsadj  = "S" then  Xnadj = sgladj
          if xtslnk  = "S" then  Xnlnk = sgllnk

                         Celdak = ""
                         Celdak = Celdak & ofv.GenCelda("","btbs","","","","","",sglprf,"si")
                         Celdak = Celdak & ofv.GenCelda("","ut","","","","","","&nbsp;<font color=#999999 ><u>" & xfromusr & "</u></font>&nbsp;&nbsp;","si")

                         Filak = ofv.GenRow("","","","","","","",Celdak,"si")
                         SArma = ofv.GenTabla("","aaa  ","","","","0","0","-1","-2","","",Filak,"si")

                     if xtmcod <> "0" then
                         TABLAN = TablaN &  "<tr >"
                         TABLAN = TablaN &  "<td class=AAA colspan=8 style='line-height:2pt;' >&nbsp;</TD>"
                         TABLAN = TablaN &  "</TR>"



                         TABLAN = TablaN &  "<tr bgcolor=" & xTrColor & " >"
                         TABLAN = TablaN &  "<td class=" & xfclas & " width=1% bgcolor=#FFFFFF >" & sImage & "</TD>"
                         TABLAN = TablaN &  "<td class=vt width=55%  STYLE='text-align:justify;background-color:#DDDDDD;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & Session("CMFPNTIT") & ", endColorStr=" & xTrColor & ", gradientType=1);' >&nbsp;&nbsp;<a class=tt >" & xtfechavig & "</a></TD>"
                         TABLAN = TablaN &  "<td class=ut width=13% >&nbsp;</TD>"


                       if session("NOVFULL") = "S" THEN

                         TABLAN = TablaN &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                         TABLAN = TablaN &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                         TABLAN = TablaN &  "<td class=btba width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnadj & "</TD>"
                         TABLAN = TablaN &  "<td class=btbn width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnlnk & "</TD>"
                         TABLAN = TablaN &  "<td class=qt width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & XMARQ & "</TD>"
                       ELSE
                         TABLAN = TablaN &  "<td class=UT COLSPAN=3 style='border-style:solid;border-width:1;border-color:#eeeeee;' >&nbsp;</TD>"
                         TABLAN = TablaN &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                         TABLAN = TablaN &  "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                       END IF


                         TABLAN = TablaN &  "</TR>"

                    end if

                         TABLAN = TablaN &  "<tr bgcolor=" & xTrColor & " >"
                         TABLAN = TablaN &  "<td class=uti colspan=8 STYLE='text-align:justify;line-height:30px;' ><font color=#999999 >" & xtmsgimg & "&nbsp;" & xtmsglnk & "</font></TD>"
                         TABLAN = TablaN &  "</TR>"



                     if session("NOVFULL") = "S" THEN

                       if xtmcod <> "0" then


                         TABLAN = TablaN &  "<tr bgcolor=" & xTrColor & " height=20 >"
                         TABLAN = TablaN &  "<td class=ut colspan=8 valign=top align=right style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' >" & sArma & "</TD>"
                         TABLAN = TablaN &  "</TR>"

                       end if

                     ELSE

                       if xtmcod <> "0" then


                         TABLAN = TablaN &  "<tr bgcolor=" & xTrColor & " height=20 >"
                         TABLAN = TablaN &  "<td class=ut colspan=8 valign=top align=right  >&nbsp;</TD>"
                         TABLAN = TablaN &  "</TR>"

                       end if

                     END IF


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF

       next


    end if



    TABLAN = TablaN &  "</table>"
    TABLAN = TablaN & "</DIV>"
    TABLAN = TablaN & "</div>"



end if



     call ofv.cerrarconn(cn)
     call ofv.cerrarconn(cn0)




''''''''''''''''''''''''''''

''''style='border-style:solid;border-width:1;border-color:#eeeeee;'
'  sHTML = sHTML &  "<table  margin=3 width=100% height=75% cellspacing=0 cellpadding=0    >"
'  sHTML = sHTML &  "<tr height=25 >"
'  sHTML = sHTML &  "<td class=ut valign=top width=20% rowspan=2 >" & tabla1 & "</TD>"
'  sHTML = sHTML &  "<td class=ut valign=top width=80%  >" & tablab & "</TD>"
'  sHTML = sHTML &  "</TR>"
'  sHTML = sHTML &  "<tr >"
'  sHTML = sHTML &  "<td class=ut valign=top width=80%  >" & tablan & "</TD>"
'  sHTML = sHTML &  "</TR>"
'  sHTML = sHTML &  "</table>"


''''''''''''''''''''''''''''

  sHTML = sHTML  &  "<DIV id='DTGENERAL' STYLE='display: table;' >"

  sHTML = sHTML  &  "<DIV id='DTheader' STYLE='float: left; z-index:1;width:22%;' >"

  sHTML = sHTML  &  tablab &  tablab1  &  tabla1

 ' sHTML = sHTML  &  "<DIV STYLE='float: left;width:95%;' ><hr></DIV >"

 ' sHTML = sHTML  &  "<DIV STYLE='float: left;width:95%;' >"
 ' sHTML = sHTML  &  "<div style='width: 260px; height: 150px; background-image: "
 ' sHTML = sHTML  &  "url( http://vortex.accuweather.com/adcbin/netweather_v2/backgrounds/gold_300x250_bg.jpg ); "
 ' sHTML = sHTML  &  "background-repeat: no-repeat; ' ><div style='height: 138px;' >"
 ' sHTML = sHTML  &  "<script src='http://netweather.accuweather.com/adcbin/netweather_v2/netweatherV2.asp?"
 ' sHTML = sHTML  &  "partner=netweather&tStyle=whteYell&logo=0&zipcode=SAM|AR|AR007|BUENOS AIRES|&lang=esp&size=8"
 ' sHTML = sHTML  &  "&theme=gold&metric=1&target=_self'></script>"
 ' sHTML = sHTML  &  "</div>"
 ' sHTML = sHTML  &  "<div style='text-align: center; font-family: arial, helvetica, verdana, sans-serif; font-size: 10px;"
 ' sHTML = sHTML  &  " line-height: 12px; color: #FFFFFF;background-color: #C39E6E;' ><a style='color: #FFFFFF' "
 ' sHTML = sHTML  &  "href='http://www.accuweather.com/world-index-forecast.asp?partner=netweather&locCode=SAM|AR|AR007|BUENOS AIRES|&metric=1' "
 ' sHTML = sHTML  &  " target='_blank' >Pronostico</a> | <a style='color: #FFFFFF' href='http://www.accuweather.com/maps-satellite.asp' target='_blank' >"
 ' sHTML = sHTML  &  "Mapas</a></div></div>"
 ' sHTML = sHTML  &  "</DIV >"
'
 ''' sHTML = sHTML  &  "<DIV STYLE='float: left;width:95%;' ><hr></DIV >"
 ''' sHTML = sHTML  &  "<iframe STYLE='float: left;' src='http://widgets.clarin.com/galerias/galeria.html'    width='260' height='240' frameborder='0' scrolling='no'></iframe>"
 ' sHTML = sHTML  &  "<DIV STYLE='float: left;width:95%;' ><hr></DIV >"
 ' sHTML = sHTML  &  "<iframe STYLE='float: left;' src='http://widgets.clarin.com/deportes/resultados.html' width='260' height='240' frameborder='0' scrolling='no'></iframe>"
 ' sHTML = sHTML  &  "<DIV STYLE='float: left;width:95%;' ><hr></DIV >"
 ' sHTML = sHTML  &  "<iframe STYLE='float: left;' src='http://widgets.clarin.com/horoscopo/horoscopo.html' width='260' height='240' frameborder='0' scrolling='no'></iframe>"
'
  sHTML = sHTML  &  "</DIV >"

 sHTML = sHTML  &  tablan

  sHTML = sHTML  &  "</DIV >"

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & WSPDTTSCP



  sHTML = sHTML &  "</body></html>"



  response.write sHTML
end if

function EsqnomP(xtipo,xcodigo,ofv)

         set xcn = ofv.conectar(ofv.strconn2)


 IF xtipo = "RSK" THEN

    strsql = "SELECT * FROM  ORRSKCIRCUITO "
    strsql = strsql & " Where OBJID = 0" & xcodigo
    set xrs = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
    if not xrs.eof then
       IF XRS("APLIC") = "OR" THEN
          xdescri = "CIRCUITO AUTOEVALUACION"
       ELSE
          xdescri = "PLAN DE CONTINGENCIAS"
       END IF

       Strsql = "Select * From objetos where codobjeto = 0" & XRS("CODOBJETO")
       set xrs = ofv.crearconsultaEx(StrSql,xcn,1,parametros)
       if not xrs.eof then
          xiden = xrs(1)
          xobjdes = xrs(3)
       end if


    END IF
    call ofv.cerrarconsulta(xrs)

    sTipobjetoX  = xdescri & ": "  & xobjdes

 ELSE


         sTipobjetoX  = xcodigo

         Strsql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
         set rsM = ofv.crearconsultaEx(StrSql2,xcn,1,parametros)
         IF NOT RSM.EOF THEN

           sTipobjeto1 = RSM(1) & ":&nbsp;"
           sTipobjeto2 = RSM(3)
           sTipobjeto3 = " - " & RSM(2)
           sTipobjeto4 = " ("  & RSM(5) & " "
           sTipobjeto5 = RSM(6) & ") "

           if session("FULLDOCN") = "1" then
              sTipobjetoX  = sTipobjeto1 & sTipobjeto2
           elseif session("FULLDOCN") = "2" then
              sTipobjetoX  = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
           else
              sTipobjetoX  = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
           end if



         END IF

 END IF





         call ofv.cerrarconn(xcn)


         esqnomP = sTipobjetoX


end function


function fechaArg(f)
 dia=day(f)
 mes=month(f)
 ano=year(f)
 fechaarg=dia &"/" & mes & "/" & ano
end function

function wmnbuscamenu(wmnSubmenu)

          Set cx0 = OFV.conectar(OFV.strconn0)
          Set cx1 = OFV.conectar(OFV.strconn1)

            xsel = 1
          StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & Session("CorApli") & "' "

          Set rs = OFV.crearconsultaEx(StrSql, cx1, 1, "")
            If Not rs.EOF Then
               If rs.fields.Count > 6 Then
                  xsel = 6
                  Call OFV.cerrarconsulta(rs)
                  StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & Session("CorApli") & "' order by Morden "

                  Set rs = OFV.crearconsultaEx(StrSql, cx1, 1, "")
               Else
                  xsel = 1
               End If
            End If
          Do Until rs.EOF



              atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'M' "
              StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & Session("CorApli") & "'"

              Set rs1 = OFV.crearconsultaEx(StrSql, cx0, 1, "")
              If rs1.EOF Then
                 atrok = "O"
              ElseIf rs1(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs1)

              xcount  = 0

              If atrok = "O" Then

                 StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
                 StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

                 if  Session("DELLOG") = "S" then

                     XMFTGRP = split(session("SegGrpLst"),chr(9))
                     xcount  = ubound(XMFTGRP) + 1
                     XMFSEP  = ""
                     StrSql = StrSql & " grupo in ("

                     for XMFI = 0 to ubound(XMFTGRP)
                         StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                         XMFSEP  = ","
                     next

                     StrSql = StrSql & ") "

                 else
                     StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
                 end if


                 StrSql = StrSql & " And Aplicacion = '" & Session("CorApli") & "'"
                 Set rs2 = OFV.crearconsultaEx(StrSql, cx0, 1, "")
                 If  rs2.EOF Then
                        atrok = "O"
                 Else
                     if  Session("DELLOG") = "S" then
                         XMFSUM = 0
                         XMFSUMN = 0
                         do until rs2.eof
                            XMFSUM = XMFSUM + 1
                            If rs2(1) Then
                               atrok = "S"
                               exit do
                            Else
                               atrok   = "N"
                               XMFSUMN = XMFSUMN + 1
                            End If
                         rs2.movenext
                         loop

                         if  atrok = "N" then
                             if  xcount >  XMFSUMN then atrok = "S"
                         end if

                     else
                         If rs2(1) Then
                            atrok = "S"
                         Else
                            atrok = "N"
                         End If
                     end if
                 End If
                 Call OFV.cerrarconsulta(rs2)

              End If


              If atrok = "O" Or atrok = "S" Then


                          sInput = Trim(rs(xsel))
                          sHref = Trim(rs(2))
                              sTarget = "info"
                              sPath = Session("CorPath")
                              If rs(3) <> "" Then atrok = "N"
                              spt = InStr(rs(2), ".asp")
                              xesm = "N"
                              If spt > 1 Then
                                  If Left(rs(2), spt - 1) = "Submenus" Then atrok = "N"
                              End If
                              sseg = InStr(sHref, "?")
                              If sseg = 0 Then
                                  sHref = sHref & "?"
                              Else
                                  sHref = sHref & "&"
                              End If
                              sHref = sHref & "seg=" & OFV.convertircar(sInput, " ", "_") & "&MNN=" & OFV.convertircar(UCase(rs(1)), " ", "_")

                              If atrok <> "N" then

                                 ii = ii + 1

                                 atit(ii)    = UCase(rs(1))
                                 awmnasp(ii) = sPath & sHref

                              end if





                  End If




              rs.movenext
          Loop
          Call OFV.cerrarconsulta(rs)
          Call OFV.cerrarconn(cx1)
          Call OFV.cerrarconn(cx0)




end function

function noveselec(xmenu)

'' xscolor = "#999999 size=2 "

 xscolor = " class=istt "'"#0000 style='font-size:8pt;font-weight:normal;' "

Do until SaleSel = "S"

   select case xmenu

       case 1


       case 2

                     '''    session("ModMain") = "CL"

                         Strsql = "SELECT * FROM csnovedades "
                         Strsql = Strsql & " Where fecexp > " & dFechai
                         Strsql = Strsql & " and fecinicio < " & dFechai + 1
                         Strsql = Strsql & " And ("
                         Strsql = Strsql & " (tipo = 2 And usugrp = '" & session("usuario") & "') or "
                         Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 999 )    or "
                         Strsql = Strsql & " (tipo = 1 And usugrp = 'TODOS' and sector = 0" & session("UN") & ") OR "
                         Strsql = Strsql & " (tipo = 51 And usugrp = 'TODOS' and sector = 999 )    or "
                         Strsql = Strsql & " (tipo = 51 And usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
                         Strsql = Strsql & ") "
                         Strsql = Strsql & " order by fecinicio desc, tipo desc, importancia, objid desc"
                       'RESPONSE.WRITE STRSQL & chr(13) & "<BR>"
                         set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if   rsx.eof then
                              xnohay1 = "S"
                         else
                              xnohay1 = "N"
                           do until rsx.eof
                              INTabla = INTabla + 1

                              xfromusr = rsx(11)

                              xxxx = getInfoUser(rsx(11))

                              xfromusr = session("SGU_nombre")

                              ANTabla(INTabla,0) = rsx(1)
                              ANTabla(INTabla,1) = clng(rsx(2))
                              ANTabla(INTabla,2) = clng(rsx(6))
                              ANTabla(INTabla,3) = clng(rsx(3))
                              ANTabla(INTabla,4) = rsx(12)
                              ANTabla(INTabla,5) = "N"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = rsx(5)
                              ANTabla(INTabla,9) = "&nbsp;"
                              ANTabla(INTabla,10) = "XXMSGXX"
                              ANTabla(INTabla,11) = rsx(2)
                              ANTabla(INTabla,12) = rsx(13)
                              ANTabla(INTabla,13) = rsx(14)
                              ANTabla(INTabla,14) = rsx(15)
                              ANTabla(INTabla,15) = rsx(16)
                              ANTabla(INTabla,16) = rsx(18)

                              ANTabla(INTabla,17) = rsx(19) & ""
                              ANTabla(INTabla,18) = rsx(20) & ""
                              ANTabla(INTabla,19) = rsx(21) & ""
                              ANTabla(INTabla,20) = rsx(22) & ""

                              ANTabla(INTabla,21) = "0"



                              rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)

       case 3,9

                    '''    session("ModMain") = "CL"

                        Strsql = "SELECT distinct * FROM csnovedades "
                        Strsql = Strsql & " Where fecexp > " & dFechai
                        Strsql = Strsql & " and fecinicio <  " & dFechai + 1
                        Strsql = Strsql & " And (tipo = 3 or tipo = 9)"
                        Strsql = Strsql & " And ("
                        Strsql = Strsql & "    (usugrp = '" & session("usuario") & "') "
                        Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 999)"
                        Strsql = Strsql & " or ( usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
                        Strsql = Strsql & "     ) "
                        Strsql = Strsql & " order by fecinicio desc, importancia, objid desc"
                '   RESPONSE.WRITE STRSQL & chr(13) & "<BR>"
                        set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if   rsx.eof then
                             xnohay1 = "S"

                        else

                             xvinoant = "no"
                             do until rsx.eof


                              if xvinoant <> rsx(1) then
                                xvinoant = rsx(1)


                               xvale = "N"

                               Strsql  = "SELECT * FROM objetos "
                               Strsql  = Strsql & " Where codobjeto = 0" & rsx(8)
                               set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                               if  not rsxp.eof then
                                   xvale = "S"
                               end if
                               call ofv.cerrarconsulta(rsxp)

                               IF  xvale = "S" THEN

                                xnohay1 = "N"
                                INTabla = INTabla + 1


                                xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


                                sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(8) & "&Tipo=" & rsx(9)
                                sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                                sHref = "&nbsp;"
                              ' response.write sActcon
                                sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)


                                wmsgx = rsx(5) & ""

                                if len(wmsgx) < 5 then
                                   wmsgx = esqnomP(rsx(9),rsx(8),ofv) & " VIGENTE A PARTIR DEL " & fechaArg(cdate(clng(rsx(3)) + clng(cdate("01-01-1900")))) & ". "
                                else
'                                  response.write "aa" & len(wmsgx) & rsx(5)
                                   wmsgx = Ucase(wmsgx)
                                end if

                                wNomL = wmsgx & "&nbsp;"



                                sImagen = wNomL

                                if xarc = "NO" then
                                   xarc = "&nbsp;"
                                end if

                                sLink = sImagen
                             '   XTSLNKM = OFV.GenLink("", "cmb title='Consultar Documento...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                                 XTSLNKM = OFV.GenLink("", "istt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                                xfromusr = rsx(11)


                                xxxx = getInfoUser(rsx(11))

                                xfromusr = session("SGU_nombre")

                              ANTabla(INTabla,0) = rsx(1)
                              ANTabla(INTabla,1) = rsx(2)
                              ANTabla(INTabla,2) = clng(rsx(6))
                              ANTabla(INTabla,3) = clng(rsx(3))
                              ANTabla(INTabla,4) = rsx(12)
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = sLink
                              ANTabla(INTabla,9) = xarc
                              ANTabla(INTabla,10) = XTSLNKM
                              ANTabla(INTabla,11) = xmenu

                             if (rsx(13) & "") = "2" then
                                 ANTabla(INTabla,12) = rsx(13)
                                 ANTabla(INTabla,13) = rsx(14)
                                 ANTabla(INTabla,14) = rsx(15)
                                 ANTabla(INTabla,15) = rsx(16)
                                 ANTabla(INTabla,16) = rsx(18)
                             else

                                 ANTabla(INTabla,12) = 2
                                 ANTabla(INTabla,13) = Capitalize(esqnomP(rsx(9),rsx(8),ofv))
                                 ANTabla(INTabla,14) = Capitalize(sLink)
                                 ANTabla(INTabla,15) = "0"
                                 ANTabla(INTabla,16) = ""
                             end if

                              ANTabla(INTabla,17) = rsx(19) & ""
                              ANTabla(INTabla,18) = rsx(20) & ""
                              ANTabla(INTabla,19) = rsx(21) & ""
                              ANTabla(INTabla,20) = rsx(22) & ""

                              ANTabla(INTabla,21) = rsx(8)

                             end if
                            end if

                            rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)

        case 4

                   '''     session("ModMain") = "WK"
                        Session("Segce")   = "S"
                        Session("Segco")   = "S"

                        Strsql = "SELECT * FROM csnovedades "
                        Strsql = Strsql & " Where fecexp > " & dFechai
                        Strsql = Strsql & " and fecinicio < " & dFechai + 1
                        Strsql = Strsql & " And tipo = 4"
                        Strsql = Strsql & " And usugrp = '" & session("usuario") & "' order by fecinicio desc, importancia, objid desc"
                       ' RESPONSE.WRITE STRSQL
                        set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if  rsx.eof then
                            xnohay1 = "S"
                        else
                            xnohay1 = "N"
                            do  until rsx.eof


                            INTabla = INTabla + 1

                            Strsql = "SELECT archivo FROM tipobjeto "
                            Strsql = Strsql & " Where tipobjeto = '" & rsx(9) & "'"
                            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if  not rs.eof then
                                xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
                            else
                                xarc = "NO"
                            end if
                            call ofv.cerrarconsulta(rs)


                            Session("ASP")      = "novedmenu.asp?menu=4"
                            session("sEstado")  = "PARA APROBAR"
                            session("sEstadon") = "2"
                            Session("Form2")    = "apr"

                            aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&Perfil=xxxx&Funcion=E"
                            aspeditar = cstr(aspeditar)
                            sHref     = "&nbsp;"
                          ' response.write sActcon
                            sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                            wNomL = "&nbsp;" & Ucase(rsx(5)) & "&nbsp;"

                            sImagen = wNomL


                                if xarc = "NO" then
                                   xarc = "&nbsp;"
                                end if

                          ''''  sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=" & xscolor & " >" & sImagen & "</font>", "si")


                            sLink = sImagen
                            'XTSLNKM = OFV.GenLink("", "cmb ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                            XTSLNKM = OFV.GenLink("", "istt ", sHref, "", "", sFuncion, "XXMSGXX", "si")


                            xfromusr = rsx(11)


                            xxxx = getInfoUser(rsx(11))
                            xfromusr = session("SGU_nombre")

                            ANTabla(INTabla,0) = rsx(1)
                            ANTabla(INTabla,1) = 4
                            ANTabla(INTabla,2) = clng(rsx(6))
                            ANTabla(INTabla,3) = clng(rsx(3))
                            ANTabla(INTabla,4) = rsx(12)
                            ANTabla(INTabla,5) = "S"
                            ANTabla(INTabla,6) = "N"
                            ANTabla(INTabla,7) = xfromusr
                            ANTabla(INTabla,8) = sLink
                            ANTabla(INTabla,9) = xarc
                            ANTabla(INTabla,10) = XTSLNKM
                            ANTabla(INTabla,11) = xmenu
                            ANTabla(INTabla,12) = rsx(13)
                            ANTabla(INTabla,13) = rsx(14)
                            ANTabla(INTabla,14) = rsx(15)
                            ANTabla(INTabla,15) = rsx(16)
                            ANTabla(INTabla,16) = rsx(18)

                              ANTabla(INTabla,17) = rsx(19) & ""
                              ANTabla(INTabla,18) = rsx(20) & ""
                              ANTabla(INTabla,19) = rsx(21) & ""
                              ANTabla(INTabla,20) = rsx(22) & ""

                              ANTabla(INTabla,21) = "0"

                         rsx.movenext
                         loop
                        end if
                        call ofv.cerrarconsulta(rsx)

'''''''''''''


set cn1 = ofv.conectar(ofv.strconn5)
set cn4 = ofv.conectar(ofv.strconn6)


                        if wkaplic = "OR" or wkaplic = "PL" then

                           xtipos = " OBJTIPO = 'RSK' "


                        else

                           xtipos = " OBJTIPO in ("
                           separ = ""

                           Strsql = "SELECT tipobjeto FROM tipobjeto "
                           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                           if  not rs.eof then
                               do while not rs.eof
                                  xtipos = xtipos & separ & "'" & rs(0) & "'"
                                  separ = ","
                                  rs.movenext
                               loop
                           else
                               xtipos = xtipos & "' '"
                           end if
                           xtipos = xtipos & ")"

                        end if





                         SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")



                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

'''''''''''''''''''''''''''''''''''''''''''
                             if  Session("DELLOG") = "S" then

                                 XMFTGRP = split(session("SegGrpLst"),chr(9))
                                 xcount  = ubound(XMFTGRP) + 1
                                 XMFSEP  = ""
                                 SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

                                 for XMFI = 0 to ubound(XMFTGRP)
                                     SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                                     XMFSEP  = ","
                                 next

                                 SEGWHERE = SEGWHERE & ") ) )"

                            else
                                 SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
                            end if


''''''''''''''''''''''''''''''''''''''''''

                          '''   IF  RS(21) > 0 THEN


                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                                 SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

                       'RESPONSE.WRITE STRSQL & "<BR>" & chr(13)
                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"

                               if rsx(3) = "RSK" then

                                   Strsql  = "SELECT * FROM ORRSKCIRCUITO "
                                   Strsql  = Strsql & " Where OBJID = 0" & rsx(4)
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)


                               else



                                   Strsql  = "SELECT * FROM objetos "
                                   Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)

                               end if




                               IF  xvale = "S" THEN

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1


                                 xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

                                 Session("ASP")      = "novedmenu.asp?menu=4"
                                 session("sEstado")  = "PARA APROBAR"
                                 session("sEstadon") = "2"
                                 Session("Form2")    = "apr"

                                 aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RSX(8) & "&MAINAPL=" & xMAINAPL


                               ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                                 aspeditar = cstr(aspeditar)
                                 sHref     = "&nbsp;"
                        '        response.write sActcon
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                                 wNomL     = "" & RSX(7) & "&nbsp;-&nbsp;PASO&nbsp;"  & RSX(16) & "&nbsp;"

                                 sImagen   = wNomL


                                 sImagen =  sImagen & "-&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv)

                                if xarc = "NO" then
                                   xarc = "&nbsp;"
                                end if

                                 sLink = sImagen
                                 'XTSLNKM = OFV.GenLink("", "cmb title='Autorizar documento ...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                                 XTSLNKM = OFV.GenLink("", "istt title='Autorizar documento ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")

                                 ANTabla(INTabla,0) = "CKT" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = rsx(10)
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink
                                 ANTabla(INTabla,9) = xarc
                                 ANTabla(INTabla,10) = XTSLNKM
                                 ANTabla(INTabla,11) = xmenu


                                 ANTabla(INTabla,12) = 2
                                 ANTabla(INTabla,13) = Capitalize(esqnomP(rsx(3),rsx(4),ofv))
                                 ANTabla(INTabla,14) = Capitalize(sLink)
                                 ANTabla(INTabla,15) = "0"
                                 ANTabla(INTabla,16) = ""

                              ANTabla(INTabla,17) = ""
                              ANTabla(INTabla,18) = ""
                              ANTabla(INTabla,19) = ""
                              ANTabla(INTabla,20) = ""

                              ANTabla(INTabla,21) = "0"

                               end if



                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)

       case 5

                    '''    session("ModMain") = "CL"

                cuenta = 0

''''''''''''''''''''''''''''''''
                 XWHERE = " and USUARIO = '" & session("usuario") & "' "
                 XWHERE = XWHERE & " AND CODOBJETO > 0 "

                 if wkaplic = "OR" or wkaplic = "PL" then

                    XWHERE = XWHERE & " AND TIPOBJETO = 'RSK' "

                 else

                    XWHERE = XWHERE & " AND TIPOBJETO <> 'RSK' "

                 end if

                 XORDER = " group by codobjeto ORDER BY fecha desc "


                 strsql = ""
                 strsql = strsql & "SELECT CODOBJETO, max(TIPOBJETO), max(finicio) as fecha  "
                 strsql = strsql & " FROM OBJALOGDET B1"
                 strsql = strsql & " WHERE TACC = 'CONSEN'  " & XWHERE

                 strsql = strsql & XORDER

               ' RESPONSE.WRITE STRSQL & "<BR>"

                        set  rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        if   rsx.eof then
                             xnohay1 = "S"

                        else

                             xvinoant = "no"
                             do until rsx.eof


                               xvale = "N"



                               if wkaplic = "OR" or wkaplic = "PL" then


                                   Strsql  = "SELECT * FROM ORRSKCIRCUITO "
                                   Strsql  = Strsql & " Where OBJID = 0" & rsx(0)
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)

                                   WKSPTIPOOBJ = "RSK"


                               else

                                   Strsql  = "SELECT * FROM objetos "
                                   Strsql  = Strsql & " Where codobjeto = 0" & rsx(0)
                 ''    RESPONSE.WRITE STRSQL & "<BR>"
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)


                                   WKSPTIPOOBJ = "VER"

                               end if




                               IF  xvale = "S" THEN


                                cuenta = cuenta + 1

                                xnohay1 = "N"
                                INTabla = INTabla + 1


                                xarc = "<font face=corflaticon color=#993333 size=4  >&#0157;</font>"


                                sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
                                sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                                sHref = "&nbsp;"
                              ' response.write sActcon
                             ''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

                                sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


                                sTipobjeto = EsqnomP(WKSPTIPOOBJ,RSX(0),ofv)






                                wmsgx = sTipobjeto & ""
                                wNomL = wmsgx & "&nbsp;"



                                sImagen = wNomL

                                if xarc = "NO" then
                                   xarc = "&nbsp;"
                                end if

                                sLink = sImagen
                                'XTSLNKM = OFV.GenLink("", "cmb title='Consultar Documento...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                                 XTSLNKM = OFV.GenLink("", "nt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                                xfromusr = "COREMANAGER"


                                xxxx = getInfoUser(xfromusr)

                                xfromusr = session("SGU_nombre")


                              ANTabla(INTabla,0) = "0"
                              ANTabla(INTabla,1) = 3
                              ANTabla(INTabla,2) = 1
                              ANTabla(INTabla,3) = dFechai
                              ANTabla(INTabla,4) = 0
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = xfromusr
                              ANTabla(INTabla,8) = sLink
                              ANTabla(INTabla,9) = xarc
                              ANTabla(INTabla,10) = XTSLNKM
                              ANTabla(INTabla,11) = xmenu
                              ANTabla(INTabla,12) = "1"
                              ANTabla(INTabla,13) = ""
                              ANTabla(INTabla,14) = ""
                              ANTabla(INTabla,15) = ""
                              ANTabla(INTabla,16) = ""

                              ANTabla(INTabla,17) = ""
                              ANTabla(INTabla,18) = ""
                              ANTabla(INTabla,19) = ""
                              ANTabla(INTabla,20) = ""

                              ANTabla(INTabla,21) = "0"

                             end if

                            if cuenta > 9 then exit do


                            rsx.movenext
                            loop
                        end if
                        call ofv.cerrarconsulta(rsx)



        case 6

                    '''    session("ModMain") = "WK"
                        Session("Segce")   = "S"
                        Session("Segco")   = "S"

                        ' RESPONSE.WRITE "6"

                        set cn1 = ofv.conectar(ofv.strconn5)
                        set cn4 = ofv.conectar(ofv.strconn6)

                         xtipos = " OBJTIPO in ('RSK'"
                         separ = ","
                         Strsql = "SELECT tipobjeto FROM tipobjeto "
                         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         if  not rs.eof then
                             do while not rs.eof
                                xtipos = xtipos & separ & "'" & rs(0) & "'"
                                separ = ","
                                rs.movenext
                             loop
                         else
                             xtipos = xtipos & "' '"
                         end if
                         xtipos = xtipos & ")"

                         SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & Session("Usuario") & "') OR ("

                         xxxx     = getInfoUser(Session("Usuario"))
                         xusrlog  = session("SGU_nombre")




                         if  xusrlog <> "NO" THEN


                             SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & Session("Usuario") & "') ) "
                             SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

'''''''''''''''''''''''''''''''''''''''''''
                             if  Session("DELLOG") = "S" then

                                 XMFTGRP = split(session("SegGrpLst"),chr(9))
                                 xcount  = ubound(XMFTGRP) + 1
                                 XMFSEP  = ""
                                 SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

                                 for XMFI = 0 to ubound(XMFTGRP)
                                     SEGWHERE = SEGWHERE  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                                     XMFSEP  = ","
                                 next

                                 SEGWHERE = SEGWHERE & ") ) )"

                            else
                                 SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & SESSION("SEGGRP") & "') ) "
                            end if


''''''''''''''''''''''''''''''''''''''''''

                          '''   IF  RS(21) > 0 THEN


                             if  cstr(session("SGU_idcostcenter") & "") <> "0" then
                                 SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                                 SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                                 SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & SESSION("SEGNIVEL") & " ) "
                                 SEGWHERE = SEGWHERE & " ) ) "
                             END IF


                         END IF

                         SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

                         Strsql = "SELECT * FROM CKTVIEW " & SEGWHERE

                    '  RESPONSE.WRITE STRSQL & "4444<BR>"
                         set rsx = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
                         if  rsx.eof then
                             xnohay2 = "S"
                         else

                             do  until rsx.eof

                               xvale = "N"

                            '   Strsql  = "SELECT * FROM objetos "
                            '   Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
                            '   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            '   if  not rsxp.eof then
                            '       xvale = "S"
                            '   end if
                            '   call ofv.cerrarconsulta(rsxp)

                               if rsx(3) = "RSK" then

                                   Strsql  = "SELECT * FROM ORRSKCIRCUITO "
                                   Strsql  = Strsql & " Where OBJID = 0" & rsx(4)
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)


                               else



                                   Strsql  = "SELECT * FROM objetos "
                                   Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
                                   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                   if  not rsxp.eof then
                                       xvale = "S"
                                       xMAINAPL = rsxp("APLIC")
                                   end if
                                   call ofv.cerrarconsulta(rsxp)

                               end if


                               IF  xvale = "S" THEN

                                 xnohay2 = "N"

                                 INTabla = INTabla + 1

                                 Strsql  = "SELECT archivo FROM tipobjeto "
                                 Strsql  = Strsql & " Where tipobjeto = '" & rsx(3) & "'"
                                 set rs  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                 if  not rs.eof then
                            '         xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
                           '      else
                           '          xarc = "NO"
                                 end if
                                 call ofv.cerrarconsulta(rs)

                                 xarc = "<font face=corflaticon color=#005500 size=4  >&#0157;</font>"

                                 Session("ASP")      = "novedmenu.asp?menu=4"
                                 session("sEstado")  = "PARA APROBAR"
                                 session("sEstadon") = "2"
                                 Session("Form2")    = "apr"

                                 aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & RSX(8)


                               ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                                 aspeditar = cstr(aspeditar)
                                 sHref     = "&nbsp;"
                        '        response.write sActcon
                                 sFuncion  = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                                 wNomL     = "" & RSX(7) & "&nbsp;-&nbsp;PASO&nbsp;"  & RSX(16) & "&nbsp;"

                                 sImagen   = wNomL


                                 sImagen =  sImagen & "-&nbsp;" & EsqnomP(rsx(3),rsx(4),ofv)


                                if xarc = "NO" then
                                   xarc = "&nbsp;"
                                end if

                                sLink = sImagen
                                XTSLNKM = OFV.GenLink("", "istt title='Ver mensaje ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")


                                 xfromusr = rsx(6)


                                 xxxx = getInfoUser(rsx(6))
                                 xfromusr = session("SGU_nombre")


                                '' ANTabla(INTabla,0) = "0"
                                 ANTabla(INTabla,0) = "MISPEN" & INTabla
                                 ANTabla(INTabla,1) = 5
                                 ANTabla(INTabla,2) = 0
                                 ANTabla(INTabla,3) = dFechai
                                 ANTabla(INTabla,4) = rsx(10)
                                 ANTabla(INTabla,5) = "S"
                                 ANTabla(INTabla,6) = "N"
                                 ANTabla(INTabla,7) = xfromusr
                                 ANTabla(INTabla,8) = sLink
                                 ANTabla(INTabla,9) = xarc
                                 ANTabla(INTabla,10) = XTSLNKM
                                 ANTabla(INTabla,11) = xmenu
                                 ANTabla(INTabla,12) = 2
                                 ANTabla(INTabla,13) = Capitalize(esqnomP(rsx(3),rsx(4),ofv))
                                 ANTabla(INTabla,14) = Capitalize(sLink)
                                 ANTabla(INTabla,15) = "0"
                                 ANTabla(INTabla,16) = ""

                              ANTabla(INTabla,17) = ""
                              ANTabla(INTabla,18) = ""
                              ANTabla(INTabla,19) = ""
                              ANTabla(INTabla,20) = ""

                              ANTabla(INTabla,21) = "0"

                               end if



                             rsx.movenext
                             loop
                        end if
                        call ofv.cerrarconsulta(rsx)





       case 10

                     '''   session("ModMain") = "CL"

                                xnohay1 = "N"


                                sLink = "<INPUT TYPE=BUTTON VALUE='HIDE MENU' onclick='hdemenu();' ><INPUT TYPE=BUTTON VALUE='SHOW MENU' onclick='mostrarmenu();' >"


                                 INTabla = INTabla + 1


                              ANTabla(INTabla,0) = "0"
                              ANTabla(INTabla,1) = 3
                              ANTabla(INTabla,2) = 1
                              ANTabla(INTabla,3) = dFechai
                              ANTabla(INTabla,4) = 0
                              ANTabla(INTabla,5) = "S"
                              ANTabla(INTabla,6) = "N"
                              ANTabla(INTabla,7) = "COREMANAGER"
                              ANTabla(INTabla,8) = sLink
                              ANTabla(INTabla,9) = "&nbsp;"
                              ANTabla(INTabla,10) = "&nbsp;"
                              ANTabla(INTabla,11) = xmenu
                              ANTabla(INTabla,12) = "1"
                              ANTabla(INTabla,13) = ""
                              ANTabla(INTabla,14) = ""
                              ANTabla(INTabla,15) = ""
                              ANTabla(INTabla,16) = ""


                              ANTabla(INTabla,17) = ""
                              ANTabla(INTabla,18) = ""
                              ANTabla(INTabla,19) = ""
                              ANTabla(INTabla,20) = ""
                              ANTabla(INTabla,21) = "0"








       case else


     end select

     if  menuSel <> "" then
         saleSel = "S"
     else
         if  INTabla > 0 then
             saleSel = "S"
         else
             menuSel = 2
             menu    = 2
         end if
     end if



  loop

end function

function noveformat(xmenu,XMENSEL)


   XTABLAN = ""


       'xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "
       xsclase = " CLASS=ietx "
       xsclase2= "class=istt "
     XTABLAN = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "


  IF XMENU = 2 THEN

     XTABLAN = XTablaN & " width:40%;"

  ELSE

     XTABLAN = XTablaN & " width:30%;"

  END IF



   XTABLAN = XTablaN & "border-style:solid;border-width:3;border-color:#ffffff;' >"


   XTABLAN = XTablaN &  "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
   XTABLAN = XTablaN &  "<tr HEIGHT=25  >"
   XTABLAN = XTablaN &  "<td " & xsclase & "   >" & Capitalize(atit(xmenu)) & "</TD>"
   XTABLAN = XTablaN &  "</TR>"
   XTABLAN = XTablaN &  "</table>"

   XTABLAN = XTablaN & "<div id='DTNOVINT" & xmenu & "' style='"

   XTABLAN = XTablaN & "border-style:solid;border-width:1;border-color:#dddddd;' >"

   XTABLAN = XTablaN &  "<table width=99% cellspacing=0 cellpadding=0  >"


   if  TNTabla = 0 then


       XTABLAN = XTablaN &  "<tr>"
       XTABLAN = XTablaN &  "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(menu) & "&nbsp;</TD>"
       XTABLAN = XTablaN &  "</TR>"

       XTABLAN = XTablaN &  "<tr>"
       XTABLAN = XTablaN &  "<td class=ut colspan=8 >&nbsp;</TD>"
       XTABLAN = XTablaN &  "</TR>"


   else

       for INTabla=1 to TNTabla

          xtmcod   = ANTabla(INTabla,0)
          xtipom   = ANTabla(INTabla,1)
          xtpri    = ANTabla(INTabla,2)
          xtfvig   = ANTabla(INTabla,3)
          xthvig   = ANTabla(INTabla,4)
          xtsadj   = ANTabla(INTabla,5)
          xtslnk   = ANTabla(INTabla,6)
          xfromusr = ANTabla(INTabla,7)
          xtmsg    = ANTabla(INTabla,8)
          xtmsgimg    = ANTabla(INTabla,9)
          xtmsglnk    = ANTabla(INTabla,10)
          xtmsgTYPE   = ANTabla(INTabla,11)



          xfclas    = "dt"
          sImage    = "&nbsp;"

          select case xtpri
                 case 0

                      xfclas = "btbn"
                      sglmod = 5
                      sgltit = "Prioridad Alta."

                case 1

                     xfclas = "btba"
                     sglmod = 4
                     sgltit = "Prioridad Normal."

                case 2

                     xfclas = "btbs"
                     sglmod = 6
                     sgltit = "Prioridad baja."


          end select

          sglmod     = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
          sImage     = ofv.GenLink("",xfclas & " title='" & sgltit & "' ","","","","",sglmod,"si")

          if  xtfvig = dFechai then
              XMARQ  = sglmar
          else
              XMARQ  = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
          end if

          IF  xtipom = 5 THEN
              xtfechavig = xthvig
              XTCOUNT    = INSTR(xtfechavig," ")
              IF XTCOUNT > 0 THEN
                 xtfechavig = LEFT(xtfechavig,XTCOUNT)
                 xtfechavig = GTFechadorsh(clng(cdate(xtfechavig)) - clng(cdate("01-01-1900")))
              end if
          ELSE
              xtfechavig = GTFechadorsh(xtfvig)
              XTCOUNT    = INSTR(xtfechavig," ")
              IF XTCOUNT > 0 THEN xtfechavig = LEFT(xtfechavig,XTCOUNT)


          END IF

          Xnlnk      = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
          Xnadj      = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

          if xtsadj  = "S" then  Xnadj = sgladj
          if xtslnk  = "S" then  Xnlnk = sgllnk









             xtmsgVRS    = "0" & ANTabla(INTabla,12)
             xtmsgVRS    = CINT(xtmsgVRS)
             xtmsgTIT    = ANTabla(INTabla,13)
             xtmsgCOP    = ANTabla(INTabla,14)
             xtmsgIMG    = ANTabla(INTabla,15)
             xtmsgVID    = ANTabla(INTabla,16)


             xtmsgVAENC      = ANTabla(INTabla,17)
             xtmsgVAFORO     = ANTabla(INTabla,18)
             xtmsgENCCOD     = ANTabla(INTabla,19)
             xtmsgFOROCOD    = ANTabla(INTabla,20)

             xtmsgdoccod    = ANTabla(INTabla,21)


             IF xtmsgVRS = 2 THEN


                         'xtmsglnk    = replace(xtmsglnk,"XXMSGXX","<font STYLE='color:#990000;font-size:10pt;font-weight:bold;' >" & xtmsgTIT & "</font>")
                          xtmsglnk    = replace(xtmsglnk,"XXMSGXX", xtmsgTIT )

                       '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                       '  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                       '  XTABLAN = XTablaN &  "</TR>"

                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                             XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
                         for yy=1 to 6
                             XTABLAN = XTablaN &  "<td class=UT width='15%'  >&nbsp;</TD>"
                         next
                             XTABLAN = XTablaN &  "<td class=UT width='5%'  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='30' >"
                         XTABLAN = XTablaN &  "<td "& xsclase2 &" colspan=8  >" & xtmsglnk & "</TD>"
                         XTABLAN = XTablaN &  "</TR>"

''''''''''''''''''''''''''''''
                     if xtipom = "9" then


                         XTMSGSEND = ""
                       ''  XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=1 >&#0124;</font></B>"
                         XTMSGSEND = XTMSGSEND & "&nbsp;" & Capitalize(xtmsgCOP)


                         XTABLAN = XTablaN &  "<tr  >"
                         XTABLAN = XTablaN &  "<td class=ut  VALIGN='TOP' align='center' ><B>"
                         XTABLAN = XTablaN &  "<font color='#AA0000' size=2 >&#0187;</font>"
                         XTABLAN = XTablaN &  "<font color='#0000AA' size=1 >&#0124;</font>"
                         XTABLAN = XTablaN &  "</B></TD>"
                         XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                         XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
                         XTABLAN = XTablaN &  "<td class=UT  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTMSGSEND = ""
                         XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >recomendado por " & ANTabla(INTabla,7) & "</font>"

                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT align='right' colspan=8 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                     end if


'''''''''''''''''''''''''''''
                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                    if xtmsgdoccod <> "0" then



                       Strsql = "Select * From objinfo where codobjeto = 0" & xtmsgdoccod & " "
                       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                       IF NOT RS1.EOF THEN

                          xtmsgCOP    = rs1(6)
                          xtmsgIMG    = rs1(7)
                          xtmsgVID    = rs1(9)

                       end if

                       Strsql = "Select * From objinfoIMG  where codobjeto = 0" & xtmsgdoccod & " ORDER BY objSECUEN, OBJID"

                       wspimgdat = 6

                    else

                       Strsql = "Select * From NOVEDADESIMG  where NOVEDAD = '" & xtmcod & "' ORDER BY SECUEN, OBJID"
                       wspimgdat = 5
                    end if



                    IF (xtmsgIMG & "") = "1" THEN



                    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                    IF NOT RS1.EOF THEN



                         XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
                         XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

                             XTABLAN = XTablaN &  "<table cellspacing=0 cellpadding=1  width=100% >"
                             XTABLAN = XTablaN &  "<tr   >"


                         XSFSTYLE= " STYLE='display:block;' "

                       DO WHILE NOT rs1.eof

                         xCelda = ""
                         xCelda = xCelda & "<div id='DIMGN" & INTabla & "' style='margin-top:1;margin-left:1;"
                         xCelda = xCelda & "width:90%;overflow:auto;"
                         xCelda = xCelda & ""
                         xCelda = xCelda & "' >" & Chr(13)
                         xCelda = xCelda & "<IMG style='border-style:solid;border-width:1;border-color:#888888;' SRC='" & RS1(wspimgdat) & "' >"
                         xCelda = xCelda & "</div>" & Chr(13)

                             XTABLAN = XTablaN &  "<td CLASS=UT align='center' " & XSFSTYLE & " width=100% >" & xCelda & "</TD>"
                             XSFSTYLE= " STYLE='display:none;' "


                        RS1.MOVENEXT
                       LOOP

                             XTABLAN = XTablaN &  "</TR>"
                             XTABLAN = XTablaN &  "</table>"

                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                    END IF






                    ELSEIF (xtmsgIMG & "") = "2" THEN


                         XTABLAN = XTablaN &  "<tr bgcolor='#F0F0F0' >"
                         XTABLAN = XTablaN &  "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
                         XTABLAN = XTablaN &  xtmsgVID
                      '   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                    END IF


                         XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
                         XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                         XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP


                         XTABLAN = XTablaN &  "<tr  >"
                         XTABLAN = XTablaN &  "<td class=ut  VALIGN='TOP' align='center' ><B><font color='#AA0000' size=2 >&#0149;</font></B></TD>"
                         XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='text-align:justify;' >"
                         XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
                         XTABLAN = XTablaN &  "<td class=UT  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"



''''' PREGUNTA SI VA ENCUESTA '''''''''''''''''''''''

         '    xtmsgVAENC      = ANTabla(INTabla,17)
         '    xtmsgVAFORO     = ANTabla(INTabla,18)
         '    xtmsgENCCOD     = ANTabla(INTabla,19)
         '    xtmsgFOROCOD    = ANTabla(INTabla,20)



                   IF xtmsgVAENC = "S" THEN

                      Strsql = "Select * From emencuestas  where objid = 0" & xtmsgENCCOD
                      set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rs1.eof then

                         xtmsgTITENC    =  Capitalize(rs1(17) & "")
                         xtmsgENCDER    = "S"


                         if (rs1(2) & "") = "1" then
                            Strsql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & xtmsgENCCOD
                            Strsql = Strsql & " Order By BLKSECUEN,objid DESC"
                            set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if not rs1.eof then
                               sblktvoto      = rs1(4) & ""
                               sblkCODIGO     = rs1(0) & ""
                               if sblktvoto = "2" then
                                  xtmsgENCDER    = "N"
                                  xtmsgTITENC    = Capitalize(rs1(3) & "")



                         XTABLAN = XTablaN &  "<tr  >"
                         XTABLAN = XTablaN &  "<td class=ut  VALIGN='TOP' align='center' >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "<td class=ut colspan=6 STYLE='color:#000000;' align='center' ><br>"

                         XTABLAN = XTablaN &  "<table width=99% cellspacing=0 cellpadding=1  style='background-color:#eeeeee;border-style:solid;border-width:1;border-color:#dddddd;' >"

                         XTABLAN = XTablaN &  "<tr HEIGHT=20  >"
                         XTABLAN = XTablaN &  "<td class=vtI COLSPAN=3  >Encuesta</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                         XTABLAN = XTablaN &  "<tr HEIGHT=20  >"
                         XTABLAN = XTablaN &  "<td class=utI width=85% style='font-size:8pt;font-weight:normal;' >" & xtmsgTITENC & "</TD>"

                         WSPDIVN = "DENC" & xtmsgENCCOD

                         Strsql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & xtmsgENCCOD
                         Strsql2 = Strsql2 & " and BLKcodigo = 0" & sblkcodigo
                         Strsql2 = Strsql2 & " Order By ITMSECUEN,objid DESC"

                         set rs2 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
                         Do While (Not rs2.EOF)


                            WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
                            WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
                            WSPFRMR = "EMENCUESTASACT.ASP?TP=ES&OP=RES&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
                            XTABLAN = XTablaN &  "<td class=ut width=5% align='center' >"
                            WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"
                            XTABLAN = XTablaN &  "<input type='button' style='width:45px;background-color:#cc0000;color:#ffffff;font-weight:bold;border-style:solid;border-width:1;border-color:#888888;' "
                            XTABLAN = XTablaN &  "  value='&nbsp;" & rs2(4) & "&nbsp;' "
                            XTABLAN = XTablaN &  "  onclick='"
                          '  XTABLAN = XTablaN &  "alert("" voto:" & rs2(4) & " "");"
                          '  XTABLAN = XTablaN &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
                            XTABLAN = XTablaN &  "document." & WSPFRMN & ".submit();"
                            XTABLAN = XTablaN &  "document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";"
                            XTABLAN = XTablaN &  "' >"
                            XTABLAN = XTablaN &  "</TD>" & WSPFORME
                            rs2.Movenext
                         Loop

                         xCelda = ""
                         xCelda = xCelda & "<div id='" & WSPDIVN & "' style='z-index:1;margin-top:1;margin-left:1;display:none;"
                         xCelda = xCelda & "width:99%;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
                         xCelda = xCelda & "background-color:#F9F9F9;"
                         xCelda = xCelda & "' >" & Chr(13)
                         xCelda = xCelda & "&nbsp;"
                         xCelda = xCelda & "</div>" & Chr(13)


                         XTABLAN = XTablaN &  "<td class=ut width=5% >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                         XTABLAN = XTablaN &  "<tr HEIGHT=20  >"
                         XTABLAN = XTablaN &  "<td class=utI width=85%  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "<td class=vt COLSPAN=2 ALIGN='CENTER' ><A CLASS=NT HREF='" & WSPFRMR & "' TARGET='fesquemas' "
                         XTABLAN = XTablaN &  " ONCLICK='document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";' "
                         XTABLAN = XTablaN &  "><B>Resultados</B></TD>"
                         XTABLAN = XTablaN &  "<td class=ut width=5% >&nbsp;</TD>"

                         XTABLAN = XTablaN &  "</TR>"

                         XTABLAN = XTablaN &  "<tr   >"
                         XTABLAN = XTablaN &  "<td class=vtI COLSPAN=4 >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"



                         XTABLAN = XTablaN &  "</table>" & xCelda
                         XTABLAN = XTablaN &  "</TD>"
                         XTABLAN = XTablaN &  "<td class=ut   >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"





                               END IF

                            END IF


                         END IF





                      END IF

                   END IF


''''''''''''''''''''''''''''

                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"











             ELSE


                         xtmsg    = replace(ltrim(xtmsg & "")," ","&nbsp;")
                         xtmsglnk    = replace(xtmsglnk,"XXMSGXX",Capitalize(xtmsg))


                         if xtmcod = "0" then

                            XTMSGSEND = "<font color=#000000 >" & xtmsglnk & "</font>"

                         ELSE

                            XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
                            XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                            XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & xtmsglnk & "</font>"

                         END IF



                         XTABLAN = XTablaN &  "<tr bgcolor=" & xTrColor & " HEIGHT='25' >"
                         XTABLAN = XTablaN &  "<td class=" & xfclas & " align='center' >" & xtmsgimg & "</TD>"
                         XTABLAN = XTablaN &  "<td class=uti colspan=6 STYLE='text-align:justify;' >" & XTMSGSEND & "</TD>"
                         XTABLAN = XTablaN &  "<td class=UT  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"


                       if xtmcod <> "0" then


                         XTABLAN = XTablaN &  "<tr bgcolor=" & xTrColor & " height=5 >"
                         XTABLAN = XTablaN &  "<td class=ut colspan=8 valign=top align=right style='border-top-style:solid;border-top-width:1;border-top-color:#eeeeee;' >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

                       end if


             END IF


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#eeeeff"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF







       next


    end if

                         XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                         XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                         XTABLAN = XTablaN &  "</TR>"

    XTABLAN = XTablaN &  "</table>"
    XTABLAN = XTablaN & "</DIV>"
    XTABLAN = XTablaN & "</DIV>"



   if  TNTabla = 0 then
       noveformat = ""
   else

       if xmenu = 4 then  atmensajes = TNTabla

       ttitsecc = ttitsecc + 1

       atitsecciones(ttitsecc,0) = Capitalize(atit(xmenu))
       atitsecciones(ttitsecc,1) = TNTabla

       noveformat = XTablaN
   end if


end FUNCTION


%>
