<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

  stipo = REQUEST.QUERYSTRING("tipo")


  set cn = ofv.conectar(ofv.strconn5)

  CODIGO = REQUEST.QUERYSTRING("CODIGO")
  SMODAL = REQUEST.QUERYSTRING("MODAL")

  IF (SMODAL & "") <> "" THEN session("ModMain") = SMODAL

  IF CODIGO <> "" THEN  session("NCliente") = CODIGO

  IF stipo = "" THEN  stipo = "REQ"


  IF stipo = "PRY" THEN

'''''''''''''''''''''''''''''''''''''''''''''''''

     tsk = REQUEST.QUERYSTRING("tsk")

     wtsk = "NO"

     IF tsk <> "" THEN  wtsk = tsk

     VPLAN = REQUEST.QUERYSTRING("VPLAN")

     IF VPLAN <> "" THEN  session("VPLAN") = VPLAN

     XSBARR1 = "BarraEnPRY.asp?tsk=" & wtsk


'''''''''''''''''''''''''''''''''''''''''''''''''''''
  else

    IF session("NCliente") = "NUEVOREQ" THEN

        XSBARR1 = "BarraEnREQNV.asp?OPC=CREARREQ&MODAL=" & SMODAL

    ELSE 

     StrSql = "Select * From REQUERIMIENTO WHERE CODIGO = '" & SESSION("NCLIENTE") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        TITULO = rs(2)
        VERSION = rs(20)
     end if



     if session("ModMain") = "WK" then
        SEGCKTPSC  = REQUEST.QUERYSTRING("SEGCKTPSC")
        IF SEGCKTPSC <> "" THEN  session("SEGCKTPSC") = SEGCKTPSC
        strsql = "UPDATE SEGPASO SET RESPASG = '" & Session("Usuario") & "' WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)
     END IF

     IF VERSION = "2" THEN

        XSBARR1 = "BarraEnREQNV.asp?OPC=R"

     ELSE

        XSBARR1 = "BarraEnREQ.asp?OPC=R"

     END IF

    END IF
     



'''''''''''''''''''''''''''''''''''''''''''''''''''
  end if

     
  session("ModMain2") = "CP"




sHTML = "<html><meta HTTP-EQUIV='expires' CONTENT='0' >"
sHTML = sHTML & "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=windows-1252'>" 

shtml = shtml & "<PUBLIC>"
shtml = shtml & "<PROPERTY NAME='CorType' />"
shtml = shtml & "<PROPERTY NAME='CorEsq' />"
shtml = shtml & "<PROPERTY NAME='CorSol' />"
shtml = shtml & "<PROPERTY name='CorMMnu' />"
shtml = shtml & "<PROPERTY name='xfwin' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMcol1' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMcol2' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMcol3' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMtopm' />" & chr(13)
shtml = shtml & "<PROPERTY name='CorMtopf' />" & chr(13)

shtml = shtml & "<METHOD   NAME='verfrm4' />"
shtml = shtml & "<METHOD   NAME='verfrm2' />"
shtml = shtml & "<METHOD   NAME='verfrm1' />"
shtml = shtml & "<METHOD   NAME='verfrm0' />"
shtml = shtml & "<METHOD   NAME='ldfobj' />"
shtml = shtml & "<METHOD   NAME='rbarr' />"

shtml = shtml & "<method name='xopcom' />"
shtml = shtml & "<method name='xclcom' />"
shtml = shtml & "<method name='xclcomCD' />"
shtml = shtml & "<method name='xclcomVD' />"
shtml = shtml & "<method name='xclcomS' />"
shtml = shtml & "<method name='abreWin' />"
shtml = shtml & "<method name='xclrstr' />" & chr(13)

shtml = shtml & " </PUBLIC>"

sHTML = sHTML & "<script language=JavaScript for=window event=onload >" & chr(13)
sHTML = sHTML & " CorType = 'D';" & chr(13)
sHTML = sHTML & " CorMMnu = 'S';" & chr(13)
sHTML = sHTML & " var xfwin    = window;" & chr(13)
sHTML = sHTML & " var CorMcol1   = '" & Session("APLSHW") & "';" & chr(13)
sHTML = sHTML & " var CorMcol2   = '" & Session("CMFMENU") & "';" & chr(13)
sHTML = sHTML & " var CorMcol3   = '" & Session("CMFPNBOT") & "';" & chr(13)
sHTML = sHTML & " window.moveTo(0,0);" & chr(13)


if ofv.explorador = "MSIE" then
   sHTML = sHTML & "window.document.title =" & chr(34)
   sHTML = sHTML & "COR ON Time - Edicion de Proyectos"
   sHTML = sHTML & chr(34) & ";" & chr(13)
end if

sHTML = sHTML & "</script>" & chr(13)

shtml = shtml & "<SCRIPT language=javascript >"

XFSCRP = "" & chr(13)
XFSCRP = XFSCRP & " var xfcomv  =  0;" & chr(13)
XFSCRP = XFSCRP & " var xfnomc  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfnoms  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xffrm   = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xffrmot = 'N';" & chr(13)
''''XFSCRP = XFSCRP & " var xfwin   = window;" & chr(13)
XFSCRP = XFSCRP & " var xfdivc  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xftdc   = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfnfld;" & chr(13)
XFSCRP = XFSCRP & " " & chr(13)
XFSCRP = XFSCRP & "function xopcom(xnomc,xnoms,xfrm,xfrmot) {" & chr(13)
XFSCRP = XFSCRP & "    "
'''XFSCRP = XFSCRP & "    xfwin    = document.frames(1);" & chr(13)
XFSCRP = XFSCRP & "    if (xfcomv== 1) {xclcomS();}"
XFSCRP = XFSCRP & "    xfcomv   =   1;"
XFSCRP = XFSCRP & "    xfnomc   = xnomc;"
XFSCRP = XFSCRP & "    xfnoms   = xnoms;"
XFSCRP = XFSCRP & "    xffrm    = xfrm;"
XFSCRP = XFSCRP & "    xffrmot  = xfrmot;"
XFSCRP = XFSCRP & "    xfdivc   = 'DF@' + xfnoms;"
XFSCRP = XFSCRP & "    xftdc    = 'TD@' + xfnoms;"
XFSCRP = XFSCRP & " "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    "
''XFSCRP = XFSCRP & "    alert(xfnfld.all[xnomc].value);"
''XFSCRP = XFSCRP & "    alert(xfwin.document.all[xnoms].value);"
''XFSCRP = XFSCRP & "    alert(xfwin.document.all[xfdivc].style.display);"

XFSCRP = XFSCRP & "    xfwin.document.all[xnoms].value=xfnfld.all[xnomc].value;"
XFSCRP = XFSCRP & "    xfwin.document.all[xfdivc].style.display='block';"

''XFSCRP = XFSCRP & "    alert(xfnfld.all[xnomc].value);"
''XFSCRP = XFSCRP & "    alert(xfwin.document.all[xnoms].value);"
''XFSCRP = XFSCRP & "    alert(xfwin.document.all[xfdivc].style.display);"
XFSCRP = XFSCRP & ""
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcom() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xclcomCD();}"
XFSCRP = XFSCRP & "    else {if (xffrmot=='2') "
XFSCRP = XFSCRP & "             {xfnfld.all[xfnomc].value=xfwin.document.all[xfnoms].value;"
XFSCRP = XFSCRP & "              xfwin.document.all[xffrm].submit();}"
XFSCRP = XFSCRP & "          else {xclcomVD();}}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomCD() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"

XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "     xfnfld.all[xfnomc].value=xfwin.document.all[xfnoms].value;"
XFSCRP = XFSCRP & "     xfnfld.all[xftdc].innerHTML="
XFSCRP = XFSCRP & "     xfwin.document.all[xfnoms].options[xfwin.document.all[xfnoms].selectedIndex].text;"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    xclcomS();"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomVD() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xfnfld=xfwin.document;}"
XFSCRP = XFSCRP & "    else {xfnfld=xfwin.document.all[xffrm];}"
XFSCRP = XFSCRP & "    "
''''XFSCRP = XFSCRP & "    xfnfld2 = xfwin.document.all['uuu'].value;"
''''XFSCRP = XFSCRP & "    alert(xfwin.document.all[xfnoms].options.length + ' ' + xfwin.document.all[xfnoms].selectedIndex);"
XFSCRP = XFSCRP & "    if (xfwin.document.all[xfnoms].options[xfwin.document.all[xfnoms].selectedIndex].value=='nsnc') "
XFSCRP = XFSCRP & "       { "
XFSCRP = XFSCRP & "         alert('La seleccion es incorrecta ...');"
XFSCRP = XFSCRP & "         xfwin.document.all[xfnoms].value=xfnfld.all[xfnomc].value;"
XFSCRP = XFSCRP & " "
''XFSCRP = XFSCRP & "         for (i = 0; i < xfwin.document.all[xfnoms].options.length; i++) "
''XFSCRP = XFSCRP & "             {if (xfwin.document.all[xfnoms].options[i].defaultSelected== true) "
''XFSCRP = XFSCRP & "                 {xfwin.document.all[xfnoms].options[i].selected=true;} "
''XFSCRP = XFSCRP & "             } "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "       } "
XFSCRP = XFSCRP & "    else "
XFSCRP = XFSCRP & "       {xfnfld.all[xfnomc].value=xfwin.document.all[xfnoms].value;"
XFSCRP = XFSCRP & "        xfwin.document.all[xffrm].submit();}"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomS() {" & chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    xfwin.document.all[xfdivc].style.display='none';"
XFSCRP = XFSCRP & "    xfcomv     =   0;"
XFSCRP = XFSCRP & "    xfnomc     = 'N';"
XFSCRP = XFSCRP & "    xfnoms     = 'N';"
XFSCRP = XFSCRP & "    xffrm      = 'N';"
XFSCRP = XFSCRP & "    xffrmot    = 'N';"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

''XFSCRP = XFSCRP & "</script>" & chr(13) & chr(13)

shtml = shtml & XFSCRP

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''

''''''''''

shtml = shtml & "function getIFD(xelID){  "
shtml = shtml & " var xxrv = null;  "
''shtml = shtml & " // if contentDocument exists, W3C compliant (Mozilla)  "
shtml = shtml & " if (document.getElementById(xelID).contentDocument){  "
shtml = shtml & "   xxrv = document.getElementById(xelID).contentDocument;  "
shtml = shtml & " } else {  "
''shtml = shtml & "   // IE  "
shtml = shtml & "   xxrv = document.frames[xelID].document;  "
shtml = shtml & " }  "
shtml = shtml & " return xxrv;  "
shtml = shtml & "  }  " & chr(13)



'''''''''''



shtml = shtml & " function ldfobj() { "
'''shtml = shtml & "   document.frames(1).document.all.DX4.innerHTML = document.frames(5).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Principal'); "
shtml = shtml & " var XFRD2 = getIFD('WK4'); "

shtml = shtml & "   XFRD1.getElementById('DX4').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } "

shtml = shtml & " function verfrm0() { "
'''shtml = shtml & "  return  document.frames(0).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Barra'); "
shtml = shtml & "  return XFRD1.body.innerHTML; "

shtml = shtml & " } "

shtml = shtml & " function verfrm1() { "
'''shtml = shtml & "  return  document.frames(2).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('WK1'); "
shtml = shtml & "  return XFRD1.body.innerHTML; "

shtml = shtml & " } "

shtml = shtml & " function verfrm2() { "
'''shtml = shtml & "  return  document.frames(3).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('WK2'); "
shtml = shtml & "  return XFRD1.body.innerHTML; "

shtml = shtml & " } "


shtml = shtml & " function verfrm4() { "
'''shtml = shtml & "  return  document.frames(4).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('WK3'); "
shtml = shtml & "  return XFRD1.body.innerHTML; "

shtml = shtml & " } "


shtml = shtml & " function verInfrm2() { "
shtml = shtml & "  alert(document.frames(3).document.body.innerHTML); "


shtml = shtml & " } "


shtml = shtml & " function verestilo() { "
''shtml = shtml & "  return  document.frames(7).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('estilo2'); "

shtml = shtml & "  return XFRD1.getElementById('estil').innerHTML; "

shtml = shtml & " } " & chr(13)

shtml = shtml & " function verfuncion() { "
''shtml = shtml & "  return  document.frames(7).document.body.innerHTML; "

shtml = shtml & " return getIFD('estilo2'); "

shtml = shtml & " } " & chr(13)



shtml = shtml & " function rbarr() { "
'''shtml = shtml & "  document.frames(2).document.reload(); "

shtml = shtml & " var XFRD1 = getIFD('WK1'); "
shtml = shtml & "  XFRD1.reload(); "


shtml = shtml & " } "

shtml = shtml & "function abreWin(Datos,nom) {" & Chr(13)
shtml = shtml & " myWin= open(Datos"
shtml = shtml & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
shtml = shtml & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
shtml = shtml & " ); " & Chr(13) & " } " & Chr(13)

shtml = shtml & " function xclrstr() { "
'''shtml = shtml & "  xfwin    = document.frames(1);" & chr(13)

shtml = shtml & " if (document.getElementById('Principal').contentDocument){  "
shtml = shtml & "   xfwin = document.getElementById('Principal');  "
shtml = shtml & " } else {  "
shtml = shtml & "   xfwin = document.frames['Principal'];  "
shtml = shtml & " }  "

shtml = shtml & " " & chr(13)
shtml = shtml & "  xfcomv  =  0;" & chr(13)
shtml = shtml & "  xfnomc  = 'N';" & chr(13)
shtml = shtml & "  xfnoms  = 'N';" & chr(13)
shtml = shtml & "  xffrm   = 'N';" & chr(13)
shtml = shtml & "  xffrmot = 'N';" & chr(13)
shtml = shtml & "  xfdivc  = 'N';" & chr(13)
shtml = shtml & "  xftdc   = 'N';" & chr(13)
shtml = shtml & " } "

shtml = shtml & "</SCRIPT>"


sHTML = sHTML & "<script language=JavaScript for=window event=onload >" & chr(13)
sHTML = sHTML & " CorType = 'D';" & chr(13)
sHTML = sHTML & " CorMMnu = 'S';" & chr(13)

if sACC = "XP" then

   sHTML = sHTML & " window.name   = 'superm';" & chr(13)

   sHTML = sHTML & " CorMtopm = window;" & chr(13)
   sHTML = sHTML & " CorMtopf = window;" & chr(13)

else

sHTML = sHTML & " CorMtopm = window.opener;" & chr(13)

'''sHTML = sHTML & " alert(CorMtopm.name);" & chr(13)
sHTML = sHTML & " for (ii=0;ii<10;ii++) {" & chr(13)
'''''sHTML = sHTML & " alert(CorMtopm.top.name + ' ' + ii);" & chr(13)
sHTML = sHTML & " if (CorMtopm.top.name=='superm') {" & chr(13)
sHTML = sHTML & "     CorMtopf = CorMtopm.top;" & chr(13)
''sHTML = sHTML & " alert(CorMtopf.name);" & chr(13)
''sHTML = sHTML & " alert(verfuncion());" & chr(13)
'sHTML = sHTML & "     var festil = verfuncion(); " & chr(13)
''sHTML = sHTML & " alert(festil.id);" & chr(13)
'sHTML = sHTML & "     festil.getElementById('dispara').action = '" & XSBARR1 & "'; " & chr(13)
'sHTML = sHTML & "     festil.getElementById('dispara').target = 'Barra'; " & chr(13)

'sHTML = sHTML & "     festil.getElementById('dispara').submit(); " & chr(13)

sHTML = sHTML & "     break; " & chr(13)
sHTML = sHTML & " } else {" & chr(13)
sHTML = sHTML & " CorMtopm = CorMtopm.top.opener;" & chr(13)
sHTML = sHTML & " }" & chr(13)
sHTML = sHTML & " }" & chr(13)

end if


sHTML = sHTML & " var xfwin    = window;" & chr(13)
sHTML = sHTML & " var CorMcol1   = '" & Session("APLSHW") & "';" & chr(13)
sHTML = sHTML & " var CorMcol2   = '" & Session("CMFMENU") & "';" & chr(13)
sHTML = sHTML & " var CorMcol3   = '" & Session("CMFPNBOT") & "';" & chr(13)
sHTML = sHTML & " window.moveTo(0,0);" & chr(13)


if ofv.explorador = "MSIE" then
   sHTML = sHTML & "window.document.title =" & chr(34)
   sHTML = sHTML & "COR ON Time - Edicion de Proyectos"
   sHTML = sHTML & chr(34) & ";" & chr(13)
end if

sHTML = sHTML & "</script>" & chr(13)






sHTML = sHTML & "<frameset rows=0,0,*,0,0,0,0 FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"


if sACC = "XP" then

shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='estilo2'    ID='estilo2'    sRC='../coress20/estilo.asp'    SCROLLING=no STYLE='display:none;' >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Barra     ID='Barra' SRC='" & XSBARR1 & "' SCROLLING=no STYLE='display:none;' > "


else

shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='estilo2'    ID='estilo2'    sRC='corbase.htm'    SCROLLING=no STYLE='display:none;' >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Barra     ID='Barra' SRC='" & XSBARR1 & "' SCROLLING=no STYLE='display:none;' > "

end if

sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Principal SRC='CorCharge.html' SCROLLING=auto > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK1       SRC='' SCROLLING=no  STYLE='display:none;' > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK2       SRC='' SCROLLING=no  STYLE='display:none;' > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK3       SRC='' SCROLLING=no  STYLE='display:none;' > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK4       SRC='' SCROLLING=no  STYLE='display:none;' > "


sHTML = sHTML & "</frameset>"
sHTML = sHTML & "<body  ><head><title>COR OnTime - consulta</title></head></body></html>"

response.write sHTML

%>