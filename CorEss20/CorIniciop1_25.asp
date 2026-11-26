
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

opcion = REQUEST.QUERYSTRING("Opcion")
PGM = REQUEST.QUERYSTRING("PGM")
If opcion = "LC" Then
    XVAR = REQUEST.QUERYSTRING
    XVARN = InStr(XVAR, "&")
    XVAR = Right(XVAR, Len(XVAR) - XVARN)

    'XVAR=REQUEST.QUERYSTRING
    XVARN = InStr(XVAR, "&")
    XVAR = Right(XVAR, Len(XVAR) - XVARN)

End If

Session("MAINWND") = "S"

shtml = "<html><head><bgsound src=" & Chr(34) & Session("sound") & Chr(34) & " >"
shtml = shtml & "<style></style><title>COR ESolution Suite V4.8 - " & OFV.fechador() & " </title>" & Chr(13)
shtml = shtml & "<public>" & Chr(13)
shtml = shtml & "<PROPERTY NAME='CorAMNN' />" & Chr(13)
shtml = shtml & "<PROPERTY NAME='CorASLP' />" & Chr(13)
shtml = shtml & "<PROPERTY NAME='CorASLPANT' />" & Chr(13)

shtml = shtml & "<PROPERTY NAME='CorType' />" & Chr(13)
shtml = shtml & "<PROPERTY NAME='CorWini' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorUsr' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorNetUsr' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorDomain' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorLogid' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorMMnu' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorMtopm' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorMtopf' />" & Chr(13)

shtml = shtml & "<PROPERTY name='CorUsrN' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorUsrE' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorEmpL' />" & Chr(13)
shtml = shtml & "<PROPERTY name='xfwin' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorMcol1' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorMcol2' />" & Chr(13)
shtml = shtml & "<PROPERTY name='CorMcol3' />" & Chr(13)

shtml = shtml & "<method name='verfrm2' />" & Chr(13)
shtml = shtml & "<method name='verfrm1' />" & Chr(13)
shtml = shtml & "<method name='verfrm0' />" & Chr(13)
shtml = shtml & "<method name='verfrmU' />" & Chr(13)
shtml = shtml & "<method name='ldfobj' />" & Chr(13)
shtml = shtml & "<method name='LDFOBJRSK' />" & Chr(13)
shtml = shtml & "<method name='LDFOBJMENU' />" & Chr(13)
shtml = shtml & "<method name='LDCALENDAR' />" & Chr(13)
shtml = shtml & "<method name='LDVENTANA' />" & Chr(13)
shtml = shtml & "<method name='verestilo' />" & Chr(13)
shtml = shtml & "<method name='verfuncion' />" & Chr(13)

shtml = shtml & "<method name='xopcom' />" & Chr(13)
shtml = shtml & "<method name='xclcom' />" & Chr(13)
shtml = shtml & "<method name='xclcomCD' />" & Chr(13)
shtml = shtml & "<method name='xclcomVD' />" & Chr(13)
shtml = shtml & "<method name='xclcomS' />" & Chr(13)
shtml = shtml & "<method name='xclrstr' />" & Chr(13)

shtml = shtml & " </public>" & Chr(13)

shtml = shtml & "<SCRIPT language=javascript >" & Chr(13)
shtml = shtml & " window.name   = 'superm';" & Chr(13)
shtml = shtml & " var CorType   = 'M';" & Chr(13)
shtml = shtml & " var CorMMnu   = 'S';" & Chr(13)
shtml = shtml & " CorMtopm      = window;" & Chr(13)
shtml = shtml & " CorMtopf      = window;CorWini=window;" & Chr(13)

shtml = shtml & " var CorUsr    = '" & Replace(Session("Usuario") & "*", " ", "_") & "';  " & Chr(13)
shtml = shtml & " var CorNetUsr = '" & Replace(Session("UsuarNT") & "*", " ", "_") & "';  " & Chr(13)
shtml = shtml & " var CorDomain = '" & Replace(Session("DOMINIO") & "*", " ", "_") & "';  " & Chr(13)
shtml = shtml & " var CorLogid  = '" & Replace(Session("LOGIDUSRDET") & "*", " ", "_") & "';  " & Chr(13)
shtml = shtml & " var CorUsrN  = '" & Session("SegUsuario") & "';  " & Chr(13)
shtml = shtml & " var CorUsrE  = '" & Session("cliemp") & "';  " & Chr(13)
shtml = shtml & " var CorEmpL  = '" & Session("clilogo") & "';  " & Chr(13)
shtml = shtml & " var xfwin    = window;" & Chr(13)
shtml = shtml & " var CorMcol1   = '" & Session("APLSHW") & "';" & Chr(13)
shtml = shtml & " var CorMcol2   = '" & Session("CMFMENU") & "';" & Chr(13)
shtml = shtml & " var CorMcol3   = '" & Session("CMFPNBOT") & "';" & Chr(13)
shtml = shtml & " var CorAMNN    ='NO';" & Chr(13)
shtml = shtml & " var CorASLP     ='NO';" & Chr(13)
shtml = shtml & " var CorASLPANT  ='NO';" & Chr(13)

shtml = shtml & "</SCRIPT>" & Chr(13)



shtml = shtml & "<SCRIPT language=javascript >" & Chr(13)

XFSCRP = "" & Chr(13)
XFSCRP = XFSCRP & " var xfcomv  =  0;" & Chr(13)
XFSCRP = XFSCRP & " var xfnomc  = 'N';" & Chr(13)
XFSCRP = XFSCRP & " var xfnoms  = 'N';" & Chr(13)
XFSCRP = XFSCRP & " var xffrm   = 'N';" & Chr(13)
XFSCRP = XFSCRP & " var xffrmot = 'N';" & Chr(13)
XFSCRP = XFSCRP & " var xfdivc  = 'N';" & Chr(13)
XFSCRP = XFSCRP & " var xftdc   = 'N';" & Chr(13)
XFSCRP = XFSCRP & " var xfnfld;" & Chr(13)
XFSCRP = XFSCRP & " " & Chr(13)

XFSCRP = XFSCRP & "function xopcom(xnomc,xnoms,xfrm,xfrmot) {" & Chr(13)
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xfcomv== 1) {xclcomS();}"
XFSCRP = XFSCRP & "    xfcomv   =   1;"
XFSCRP = XFSCRP & "    xfnomc   = xnomc;"
XFSCRP = XFSCRP & "    xfnoms   = xnoms;"
XFSCRP = XFSCRP & "    xffrm    = xfrm;"
XFSCRP = XFSCRP & "    xffrmot  = xfrmot;"
XFSCRP = XFSCRP & "    xfdivc   = 'DF@' + xfnoms;"
XFSCRP = XFSCRP & "    xftdc    = 'TD@' + xfnoms;"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    var xfnfld2   =   '';"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1')  {                         "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD;                        "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.getElementById(xnomc);  "
XFSCRP = XFSCRP & "                       }                         "
XFSCRP = XFSCRP & "    else {                                       "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD.getElementById(xffrm);  "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.forms[xffrm].elements[xnomc];  "
XFSCRP = XFSCRP & "          } "
XFSCRP = XFSCRP & "    "

''XFSCRP = XFSCRP & " alert(xfnfld.elements.getElementById(xnomc).name); "
''XFSCRP = XFSCRP & " alert(xfwinCD.forms[xffrm].elements[xnomc].name); "
'XFSCRP = XFSCRP & "    xfnfld2.selectedIndex=0;"

'''XFSCRP = XFSCRP & "    xfwinCD.getElementById(xnoms).value=xfnfld.getElementById(xnomc).value;"
''XFSCRP = XFSCRP & "    xfwinCD.getElementById(xnoms).value = xfwinCD.getElementById(xnomc).value;"
'XFSCRP = XFSCRP  & "   xfnfld.attachEvent('onclick', function() { alert('3333'); });"
XFSCRP = XFSCRP & "    xfwinCD.getElementById(xnoms).value = xfnfld2.value; "


XFSCRP = XFSCRP & "    xfwinCD.getElementById(xfdivc).style.display='block';"

XFSCRP = XFSCRP & ""
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & Chr(13)
XFSCRP = XFSCRP & "function dothing2() {" & Chr(13)
XFSCRP = XFSCRP & "alert('3333');}" & Chr(13)
XFSCRP = XFSCRP & "function xclcom() {" & Chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    var xfnfld2   =   '';"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1')  {                         "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD;                        "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.getElementById(xfnomc);  "
XFSCRP = XFSCRP & "                       }                         "
XFSCRP = XFSCRP & "    else {                                       "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD.getElementById(xffrm);  "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.forms[xffrm].elements[xfnomc];  "
XFSCRP = XFSCRP & "          } "
XFSCRP = XFSCRP & "    "

XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1') {xclcomCD();}"
XFSCRP = XFSCRP & "    else {if (xffrmot=='2') "
'''XFSCRP = XFSCRP & "             {xfnfld.getElementById(xfnomc).value = xfwinCD.getElementById(xfnoms).value;"
''XFSCRP = XFSCRP & "             {xfwinCD.getElementById(xfnomc).value = xfwinCD.getElementById(xfnoms).value;"

XFSCRP = XFSCRP & "             {xfnfld2.value = xfwinCD.getElementById(xfnoms).value;"


XFSCRP = XFSCRP & "              xfwinCD.getElementById(xffrm).submit();}"
XFSCRP = XFSCRP & "          else {xclcomVD();}}"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & Chr(13)

XFSCRP = XFSCRP & "function xclcomCD() {" & Chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    var xfnfld2   =   '';"
XFSCRP = XFSCRP & "    var xfnfld3   =   '';"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1')  {                         "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD;                        "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.getElementById(xfnomc);  "
XFSCRP = XFSCRP & "          xfnfld3 = xfwinCD.getElementById(xftdc);  "
XFSCRP = XFSCRP & "                       }                         "
XFSCRP = XFSCRP & "    else {                                       "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD.getElementById(xffrm);  "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.forms[xffrm].elements[xfnomc];  "
XFSCRP = XFSCRP & "          xfnfld3 = xfwinCD.forms[xffrm].elements[xftdc];  "
XFSCRP = XFSCRP & "          } "
XFSCRP = XFSCRP & "    "



XFSCRP = XFSCRP & "    "
'''XFSCRP = XFSCRP & "     xfnfld.getElementById(xfnomc).value = xfwinCD.getElementById(xfnoms).value;"
''XFSCRP = XFSCRP & "     xfwinCD.getElementById(xfnomc).value = xfwinCD.getElementById(xfnoms).value;"

XFSCRP = XFSCRP & "     xfnfld2.value = xfwinCD.getElementById(xfnoms).value;"



'''XFSCRP = XFSCRP & "     xfnfld.getElementById(xftdc).innerHTML = "
''XFSCRP = XFSCRP & "     xfwinCD.getElementById(xftdc).innerHTML = "

XFSCRP = XFSCRP & "     xfnfld3.innerHTML = "

XFSCRP = XFSCRP & "     xfwinCD.getElementById(xfnoms).options[xfwinCD.getElementById(xfnoms).selectedIndex].text;"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    xclcomS();"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & Chr(13)

XFSCRP = XFSCRP & "function xclcomVD() {" & Chr(13)
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    var xfnfld2   =   '';"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1')  {                         "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD;                        "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.getElementById(xfnomc);  "
XFSCRP = XFSCRP & "                       }                         "
XFSCRP = XFSCRP & "    else {                                       "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD.getElementById(xffrm);  "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.forms[xffrm].elements[xfnomc];  "
XFSCRP = XFSCRP & "          } "
XFSCRP = XFSCRP & "    "

XFSCRP = XFSCRP & "    if (xfwinCD.getElementById(xfnoms).options[xfwinCD.getElementById(xfnoms).selectedIndex].value=='nsnc') "
XFSCRP = XFSCRP & "       { "
XFSCRP = XFSCRP & "         alert('La seleccion es incorrecta ...');"
'''XFSCRP = XFSCRP & "         xfwinCD.getElementById(xfnoms).value=xfnfld.getElementById(xfnomc).value;"
''XFSCRP = XFSCRP & "         xfwinCD.getElementById(xfnoms).value = xfwinCD.getElementById(xfnomc).value;"

XFSCRP = XFSCRP & "         xfwinCD.getElementById(xfnoms).value = xfnfld2.value;"
XFSCRP = XFSCRP & " "

XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "       } "
XFSCRP = XFSCRP & "    else "
'''XFSCRP = XFSCRP & "       {xfnfld.getElementById(xfnomc).value=xfwinCD.getElementById(xfnoms).value;"
''XFSCRP = XFSCRP & "       {xfwinCD.getElementById(xfnomc).value = xfwinCD.getElementById(xfnoms).value;"

XFSCRP = XFSCRP & "       {xfnfld2.value = xfwinCD.getElementById(xfnoms).value;"

XFSCRP = XFSCRP & "        xfwinCD.getElementById(xffrm).submit();}"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & Chr(13)

XFSCRP = XFSCRP & "function xclcomS() {" & Chr(13)
XFSCRP = XFSCRP & " "

XFSCRP = XFSCRP & "    xfwinCD.getElementById(xfdivc).style.display='none';"
XFSCRP = XFSCRP & "    xfcomv     =   0;"
XFSCRP = XFSCRP & "    xfnomc     = 'N';"
XFSCRP = XFSCRP & "    xfnoms     = 'N';"
XFSCRP = XFSCRP & "    xffrm      = 'N';"
XFSCRP = XFSCRP & "    xffrmot    = 'N';"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & Chr(13)

shtml = shtml & XFSCRP

Whtml = ""

''''''''''
xshtml= "<script language=javascript>" & chr(13)
    xshtml=xshtml & "function shwDem2(){var xif=getIFD('info');" & chr(13)
    xshtml=xshtml & " if (xif.getElementById('DEM12').style.display=='none')"
    xshtml=xshtml & "{ "
    xshtml=xshtml & "xif.getElementById('DEM12').style.display='block';" & chr(13)
    xshtml=xshtml & "}"
   xshtml=xshtml & " else"
    xshtml=xshtml & "{"
    xshtml=xshtml & "xif.getElementById('DEM12').style.display='none';" & chr(13)
   xshtml=xshtml & "} "

    xshtml=xshtml & "}" & chr(13)
    xshtml=xshtml  & "</script>" & chr(13)

sc="<script> "
    sc=sc & "function ejec(h){var xif=getIFD('info');"
    sc=sc & "xif.getElementById('xfx').action=h; xif.getElementById('xfx').submit();"
    sc=sc & "}</script>"

shtml =  sc & xshtml & shtml & "function getIFD(xelID){  "
shtml = shtml & " var xxrv = null;  "
''shtml = shtml & " // if contentDocument exists, W3C compliant (Mozilla)  "
shtml = shtml & " if (document.getElementById(xelID).contentDocument){  "
shtml = shtml & "   xxrv = document.getElementById(xelID).contentDocument;  "
shtml = shtml & " } else {  "
''shtml = shtml & "   // IE  "
shtml = shtml & "   xxrv = document.frames[xelID].document;  "
shtml = shtml & " }  "
shtml = shtml & " return xxrv;  "
shtml = shtml & "  }  " & Chr(13)



'''''''''''

shtml = shtml & " function LDCALENDAR() { "

shtml = shtml & " var XFRD1 = getIFD('info'); "
shtml = shtml & " var XFRD2 = getIFD('fesquemas'); "

shtml = shtml & "   XFRD1.getElementById('DXCALENDAR').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & Chr(13)


shtml = shtml & " function LDVENTANA(LDVID) { "

shtml = shtml & " var XFRD1 = getIFD('info'); "
shtml = shtml & " var XFRD2 = getIFD('fesquemas'); "

shtml = shtml & "   XFRD1.getElementById(LDVID).innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & Chr(13)


shtml = shtml & " function ldfobj() { "
''shtml = shtml & "   document.frames(0).document.getElementById('DX4').innerHTML = document.frames(5).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('info'); "
shtml = shtml & " var XFRD2 = getIFD('fesquemas'); "

shtml = shtml & "   XFRD1.getElementById('DX4').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & Chr(13)

shtml = shtml & " function verfrmU() { "
shtml = shtml & "  var MDUSR = '" & Whtml & "'; "
shtml = shtml & "  return  MDUSR; "
shtml = shtml & " } " & Chr(13)


shtml = shtml & " function LDFOBJRSK() { "

shtml = shtml & " var XFRD1 = getIFD('info'); "
shtml = shtml & " var XFRD2 = getIFD('fesquemas'); "

shtml = shtml & "   XFRD1.getElementById('DXA').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & Chr(13)

shtml = shtml & " function LDFOBJMENU(GETDSD,GETHSD) { "

shtml = shtml & " var XFRD1 = getIFD('info'); "
shtml = shtml & " var XFRD2 = getIFD(GETDSD); "

shtml = shtml & "   XFRD1.getElementById(GETHSD).innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & Chr(13)


Whtml = ""

shtml = shtml & " function verfrm0() { "
''shtml = shtml & "  return  document.frames(1).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Footer'); "

shtml = shtml & "  return XFRD1.body.innerHTML; "


shtml = shtml & " } " & Chr(13)

shtml = shtml & " function verfrm1() { "
''shtml = shtml & "  return  document.frames(3).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Menus'); "

shtml = shtml & "  return XFRD1.body.innerHTML; "


shtml = shtml & " } " & Chr(13)

shtml = shtml & " function verfrm2() { "
''shtml = shtml & "  return  document.frames(4).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Master'); "

shtml = shtml & "  return XFRD1.body.innerHTML; "

shtml = shtml & " } " & Chr(13)


shtml = shtml & " function verestilo() { "
''shtml = shtml & "  return  document.frames(7).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('estilo'); "

shtml = shtml & "  return XFRD1.getElementById('estil').innerHTML; "

shtml = shtml & " } " & Chr(13)

shtml = shtml & " function verfuncion() { "
''shtml = shtml & "  return  document.frames(7).document.body.innerHTML; "

shtml = shtml & " return getIFD('estilo'); "

shtml = shtml & " } " & Chr(13)



shtml = shtml & " function xclrstr() { "
'''shtml = shtml & "  xfwin    = document.frames(0);" & chr(13)

shtml = shtml & " if (document.getElementById('info').contentDocument){  "
shtml = shtml & "   xfwin = document.getElementById('info');  "
shtml = shtml & "   xfwinCD = xfwin.contentDocument;  "
shtml = shtml & " } else {  "
shtml = shtml & "   xfwin = document.frames['info'];  "
shtml = shtml & "   xfwinCD = xfwin.document;  "
shtml = shtml & " }  "

shtml = shtml & " " & Chr(13)
shtml = shtml & "  xfcomv  =  0;" & Chr(13)
shtml = shtml & "  xfnomc  = 'N';" & Chr(13)
shtml = shtml & "  xfnoms  = 'N';" & Chr(13)
shtml = shtml & "  xffrm   = 'N';" & Chr(13)
shtml = shtml & "  xffrmot = 'N';" & Chr(13)
shtml = shtml & "  xfdivc  = 'N';" & Chr(13)
shtml = shtml & "  xftdc   = 'N';" & Chr(13)
shtml = shtml & " } " & Chr(13)


shtml = shtml & "</SCRIPT>"

shtml =   shtml & "<SCRIPT FOR=window EVENT=onunload>"
''shtml = shtml & "    alert('La sesion de Cor E-Solution Suite se cerrara.' + window. event.clientX  + ' ' + window. event.clientY );"
shtml = shtml & " if (window. event.clientX < 0) { "


shtml = shtml & " myWin= open(" & Chr(34) & "CorCierra.ASP?USR=" & Replace(Session("Usuario") & "*", " ", "_")
shtml = shtml & "&NETUSR=" & Replace(Session("UsuarNT") & "*", " ", "_")
shtml = shtml & "&DMN=" & Replace(Session("DOMINIO") & "*", " ", "_")
shtml = shtml & "&LID=" & Replace(Session("LOGIDUSRDET") & "*", " ", "_")
shtml = shtml & Chr(34) & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
shtml = shtml & "menubar=no,toolbar=no,height=300,width=400" & Chr(34) & " ); "
shtml = shtml & " } "
shtml = shtml & "</SCRIPT>"
shtml = shtml & "</head>"


If Session("VIEWOPC") = "IXIND" Then

    WKINITASP = "../COREMANAGER/CorINDEXLISTA.asp"
    Session("SMPLMENU") = "S"
    Session("CorApli") = "COREM20"


ElseIf opcion = "LC" Then
    Session("SMPLMENU") = "S"

    'usuario="PEPE"
    Usuario = UserLogon()
    Session("UN") = 1
    Session("USREXPIR") = "D"

    Call getInfoUser(Usuario)
    Session("Usuario") = Session("SGU_userid")
    Session("UsrOpcion") = Session("SGU_nombre")
    Session("SkrDefT") = "0"
    Session("SkrDefE") = "0"
    Session("SkrDefI") = "0"
    Session("UsrPmin") = "0"
    Session("UsrPmax") = "0"
    Session("UsrHist") = "0"
    Session("SegUsuario") = Session("SGU_nombre")
    Session("SegNivel") = 8
    Session("SegGrp") = "DEF"
    Session("SegPerfil") = Session("Usuario")
    Session("UsuarNT") = Session("Usuario")

    Session("UsrSector") = Session("SGU_idcostcenter")
    Session("UsrSecDsc") = Session("SGU_descrisector")
    Session("UsrPuesto") = Session("SGU_codigopuesto")
    Session("UsrPstDsc") = Session("SGU_descripuesto")
    Session("UsrMail") = Session("SGU_mail")

    Call getGroupsUser(Usuario)

    Session("SegGrpOpt") = "S"

    Session("SegGrpLst") = Session("SGU_grupos")

    Call ChkGrp()
    Set cn = OFV.conectar(OFV.strconn2)
    StrSql = "Select * From Clienteplan WHERE ID ="  & SESSION("CLIEMP")
    Set RS = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not RS.EOF Then
        Session("Cliente") = RS(1)
        Session("CliLogo") = "../COREss20/Empresa/" & RS(5)
        If (RS(11) & "") <> "" Then Session("CliLogoP") = "../COREss20/Empresa/" & RS(11)
        Session("USRSET") = RS(10)
    End If
    Call OFV.cerrarconn(cn)
Else


    WKINITASP = "inicio.asp?opcion=" & opcion & ""

    Novedform = "CLIMN16317206150930885546227"
    Session("CorApli") = "COREM20"

    OFV.ObtenerAtributosNOdo Novedform, "", "TIPO", CStr("10")
    bConsultar = OFV.AConsultar
    Session("SMPLMENU") = "N"
    If bConsultar Then
        Session("SMPLMENU") = "N"
    Else
        Session("SMPLMENU") = "S"
    End If

End If

    ofv.ObtenerAtributosNOdo Novedform, "", "BOTON", "EMPRESA"
    Session("UsrCorp") = OFV.AConsultar

'Session("CorApli") = ""


shtml = shtml & "<frameset cols='300,*,0,0,0,0,0,0' FRAMESPACING=0  FRAMEBORDER=1 BORDER=1 >"

shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='estilo'    ID='estilo'    sRC='estilo.asp'    SCROLLING=no STYLE='xdisplay:none;' >"

If opcion = "LC" Then
  '    shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info'    ID='info'    SRC='http://vsd001lx1.blp.com.ar/desarrollo/dvara/Intranet/" & PGM & ".php?" & XVAR & "'  SCROLLING=no STYLE='xdisplay:none;' >"
      shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info'    ID='info'    SRC='http://www.google.com'  SCROLLING=no STYLE='xdisplay:none;' >"
Else
    shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info'      ID='info'      SRC='" & WKINITASP & "' SCROLLING=auto >"
End If

' shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='info'      ID='info'      SRC='" & WKINITASP & "' SCROLLING=auto >"
shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Footer'    ID='Footer'    SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Estad'     ID='Estad'     SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Menus'     ID='Menus'     SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='Master'    ID='Master'    SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='fesquemas' ID='fesquemas' SRC='corbase.htm'   SCROLLING=no STYLE='display:none;' >"
shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='expir'     ID='expir'     sRC='refrescar.asp' SCROLLING=no STYLE='display:none;' >"



shtml = shtml & "<noframes><body><p>Esta página usa marcos, pero su explorador no los admite.</p></body></noframes>"
shtml = shtml & "</frameset></html>"






response.write shtml


 function ChkGrp()

 '    SESSION("USREXPIR") = 50

   SESSION("USREXPIR") = "0"

   set XFCN = ofv.conectar(ofv.strconn0)

   XFStrSql = "Select id,sessexp,descripcion From Grupos where "

   session("SegGrpLst") = Session("SegGrp") & chr(9) & session("SegGrpLst")
  session("SegGrpdsc") = Session("SegGrp") & chr(9) & session("SegGrpLst")


   XMFTGRP = split(session("SegGrpLst"),chr(9))
   XMFSEP  = ""
   XFStrSql = XFStrSql & " id in ("

   for XMFI = 0 to ubound(XMFTGRP)
       XFStrSql  = XFStrSql  & XMFSEP & "'" & replace(XMFTGRP(XMFI) & "","'","''") & "'"
       XMFSEP  = ","
   next

   XFStrSql = XFStrSql & ") "
   XFStrSql = XFStrSql & " Order By id"
 ''''  response.write XFStrSql
   set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

   if  not XFrs.eof then
       XMFSEP               = ""
       session("SegGrpLst") = ""
       session("SegGrpdsc") = ""
       do until XFRS.EOF

          if XFrs(1) <> "D" and SESSION("USREXPIR") <> "D" then
             if (cint("0" & XFrs(1)) * 60) > cint(SESSION("USREXPIR")) then
                SESSION("USREXPIR") = cint("0" & XFrs(1)) * 60
             end if
           else
            SESSION("USREXPIR") = "D"
           end if

          session("SegGrpLst") = session("SegGrpLst") & XMFSEP & XFRS(0)
           session("SegGrpdsc") = session("SegGrpdsc") & XMFSEP & XFRS(2)
          XMFSEP  = chr(9)
       XFRS.MOVENEXT
       LOOP

       IF  session("SegGrpLst") = "" THEN session("SegGrpLst") = Session("SegGrp")

   else
       session("SegGrpLst") = Session("SegGrp")
       session("SegGrpdsc") = Session("SegGrp")
   end if
   call ofv.cerrarconsulta(XFrs)
   call ofv.cerrarconn(XFCN)



end function

%>
