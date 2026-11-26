<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0
nCodigo   = request.querystring("Codigo")
sTipo     = request.querystring("Tipo")
sPerfil   = request.querystring("Perfil")
nFuncion  = request.querystring("Funcion")
sSClick   = request.querystring("SSClick")
sNIV1     = request.querystring("NIV1")
sNIV2     = request.querystring("NIV2")
sNIV3     = request.querystring("NIV3")
sNIV4     = request.querystring("NIV4")
sNIV5     = request.querystring("NIV5")
sACC      = request.querystring("ACC")
sMAINAPL  = request.querystring("MAINAPL")

IF SMAINAPL = "" THEN SMAINAPL = "EM"

sDatos    = ""
if request.querystring("Datos") <> "" then
  sDatos  = ofv.convertircar(request.querystring("Datos"),"*","&")
end if

SEGCKTPSC = request.querystring("SEGCKTPSC") & ""

if SEGCKTPSC = "" then
   session("SEGCKTPSC") = "NO"
else
   session("SEGCKTPSC") = SEGCKTPSC
end if

if  ssclick = "" then ssclick = 99

      IF SNIV1 = "" THEN SNIV1 = "0"
      IF SNIV2 = "" THEN SNIV2 = "0"
      IF SNIV3 = "" THEN SNIV3 = "0"
      IF SNIV4 = "" THEN SNIV4 = "0"
      IF SNIV5 = "" THEN SNIV5 = "0"

IF SNIV1 <> "0" THEN
   ssclick = SNIV1
END IF


Solapas = 8

  set cn = ofv.conectar(ofv.strconn2)

 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

   IF SMAINAPL = "OR" THEN
      sEsquema = "CIRCUITO AUTOEVALUACION"
   ELSE
      sEsquema = "PLAN DE CONTINGENCIAS"
   END IF



  xestado = 1

  Strsql = "Select estado,periodo From ORRSKCIRCUITO Where OBJID = 0" & nCodigo
  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx1.eof then
     xestado = rsx1(0)
     session("ORPERIODO") = rsx1(1)
  end if
  call ofv.cerrarconsulta(rsx1)


 ELSE

  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then   sEsquema = rs(1)
  call ofv.cerrarconsulta(rs)


  Strsql = "Select * From Objperfil Where tipoperfil='" & sPerfil & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  Solapas = 1
  for z = 2 to 12
    if rs(z) Then
       Solapas = Solapas + 1
    end if
  next
  end if
  call ofv.cerrarconsulta(rs)

  xestado = 1

  Strsql = "Select estado From Objprty Where codobjeto =" & nCodigo & " and tipobjeto ='" & stipo & "' "
  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx1.eof then
     xestado = rsx1(0)
  end if
  call ofv.cerrarconsulta(rsx1)

 END IF


  call ofv.cerrarconn(cn)

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
shtml = shtml & "<PROPERTY NAME='CorASLPEX' />" & chr(13)
shtml = shtml & "<PROPERTY NAME='CorASLPANTEX' />" & chr(13)
shtml = shtml & "<METHOD   NAME='verfrm2' />"
shtml = shtml & "<METHOD   NAME='verfrm1' />"
shtml = shtml & "<METHOD   NAME='verfrm1SP' />"
shtml = shtml & "<METHOD   NAME='verfrm0' />"
shtml = shtml & "<METHOD   NAME='ldfobj' />"
shtml = shtml & "<METHOD   NAME='LDFOBJRSK' />"
shtml = shtml & "<method   name='LDFOBJMENU' />" & chr(13)
shtml = shtml & "<METHOD   name='LDCALENDAR' />" & chr(13)
shtml = shtml & "<METHOD   NAME='rbarr' />"
shtml = shtml & "<method   name='verestilo' />"
shtml = shtml & "<method   name='verfuncion' />" & chr(13)

shtml = shtml & "<method name='xopcom' />"
shtml = shtml & "<method name='xclcom' />"
shtml = shtml & "<method name='xclcomCD' />"
shtml = shtml & "<method name='xclcomVD' />"
shtml = shtml & "<method name='xclcomS' />"
shtml = shtml & "<method name='abreWin' />"
shtml = shtml & "<method name='xclrstr' />" & chr(13)

shtml = shtml & " </PUBLIC>"


shtml = shtml & "<SCRIPT language=javascript >"

XFSCRP = "" & chr(13)
XFSCRP = XFSCRP & " var xfcomv  =  0;" & chr(13)
XFSCRP = XFSCRP & " var xfnomc  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfnoms  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xffrm   = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xffrmot = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfdivc  = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xftdc   = 'N';" & chr(13)
XFSCRP = XFSCRP & " var xfnfld;" & chr(13)
XFSCRP = XFSCRP & " " & chr(13)


XFSCRP = XFSCRP & "function xopcom(xnomc,xnoms,xfrm,xfrmot) {" & chr(13)
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xfcomv== 1) {xclcomS();}"
XFSCRP = XFSCRP & "    xfcomv   =   1;"
XFSCRP = XFSCRP & "    xfnomc   = xnomc;"
XFSCRP = XFSCRP & "    xfnoms   = xnoms;"
XFSCRP = XFSCRP & "    xffrm    = xfrm;"
XFSCRP = XFSCRP & "    xffrmot  = xfrmot;"
XFSCRP = XFSCRP & "    xfdivc   = 'DF@' + xfnoms;"
XFSCRP = XFSCRP & "    xftdc    = 'TD@' + xfnoms;"
''XFSCRP = XFSCRP & " alert(xfwinCD.forms[xffrm].name); "
''XFSCRP = XFSCRP & " alert(xffrmot); "
''XFSCRP = XFSCRP & " alert(xnomc); "

XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "    var xfnfld2   =   '';"
XFSCRP = XFSCRP & "    "
XFSCRP = XFSCRP & "    if (xffrmot=='1')  {                         "
''XFSCRP = XFSCRP & "          alert(xfwinCD.getElementById(xnomc).name); "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD;                        "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.getElementById(xnomc);  "
XFSCRP = XFSCRP & "                       }                         "
XFSCRP = XFSCRP & "    else {                                       "
''XFSCRP = XFSCRP & " alert(xnomc); "
XFSCRP = XFSCRP & "          xfnfld  = xfwinCD.getElementById(xffrm);  "
XFSCRP = XFSCRP & "          xfnfld2 = xfwinCD.forms[xffrm].elements[xnomc];  "
XFSCRP = XFSCRP & "          } "
XFSCRP = XFSCRP & "    "

''XFSCRP = XFSCRP & " alert(xfnfld.elements.getElementById(xnomc).name); "
''XFSCRP = XFSCRP & " alert(xfwinCD.forms[xffrm].elements[xnomc].name); "
''XFSCRP = XFSCRP & " alert(xfwinCD.forms[xffrm].name); "

'''XFSCRP = XFSCRP & "    xfwinCD.getElementById(xnoms).value=xfnfld.getElementById(xnomc).value;"
''XFSCRP = XFSCRP & "    xfwinCD.getElementById(xnoms).value = xfwinCD.getElementById(xnomc).value;"

XFSCRP = XFSCRP & "    xfwinCD.getElementById(xnoms).value = xfnfld2.value;"


XFSCRP = XFSCRP & "    xfwinCD.getElementById(xfdivc).style.display='block';"

XFSCRP = XFSCRP & ""
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcom() {" & chr(13)
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
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomCD() {" & chr(13)
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
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomVD() {" & chr(13)
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
XFSCRP = XFSCRP & "}" & chr(13)

XFSCRP = XFSCRP & "function xclcomS() {" & chr(13)
XFSCRP = XFSCRP & " "

XFSCRP = XFSCRP & "    xfwinCD.getElementById(xfdivc).style.display='none';"
XFSCRP = XFSCRP & "    xfcomv     =   0;"
XFSCRP = XFSCRP & "    xfnomc     = 'N';"
XFSCRP = XFSCRP & "    xfnoms     = 'N';"
XFSCRP = XFSCRP & "    xffrm      = 'N';"
XFSCRP = XFSCRP & "    xffrmot    = 'N';"
XFSCRP = XFSCRP & " "
XFSCRP = XFSCRP & "}" & chr(13)

''XFSCRP = XFSCRP & "</script>" & chr(13) & chr(13)

shtml = shtml & XFSCRP


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

shtml = shtml & " function LDCALENDAR() { "

shtml = shtml & " var XFRD1 = getIFD('Principal'); "
shtml = shtml & " var XFRD2 = getIFD('WK4'); "

shtml = shtml & "   XFRD1.getElementById('DXCALENDAR').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } " & chr(13)

shtml = shtml & " function ldfobj() { "
'''shtml = shtml & "   document.frames(1).document.all.DX4.innerHTML = document.frames(5).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('Principal'); "
shtml = shtml & " var XFRD2 = getIFD('WK4'); "

shtml = shtml & "   XFRD1.getElementById('DX4').innerHTML = XFRD2.body.innerHTML; "

shtml = shtml & " } "

shtml = shtml & " function LDFOBJRSK() { "

shtml = shtml & " var XFRD1 = getIFD('Principal'); "
shtml = shtml & " var XFRD2 = getIFD('WK4'); "

shtml = shtml & "   XFRD1.getElementById('DXA').innerHTML = XFRD2.body.innerHTML; "

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

shtml = shtml & " function verfrm1SP() { "
shtml = shtml & " var XFRD1 = getIFD('WK1'); "
'''''shtml = shtml & " alert(XFRD1.innerHTML); "
shtml = shtml & " var XFRD1SP = XFRD1.getElementById('WKTBSP1'); "
shtml = shtml & "  return XFRD1SP.innerHTML; "

shtml = shtml & " } "

shtml = shtml & " function verfrm2() { "
'''shtml = shtml & "  return  document.frames(3).document.body.innerHTML; "

shtml = shtml & " var XFRD1 = getIFD('WK2'); "
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
shtml = shtml & "status=yes,menubar=no,toolbar=no,height=570,width=790" & Chr(34)
shtml = shtml & " ); " & Chr(13) & " } " & Chr(13)

shtml = shtml & " function xclrstr() { "

''xfwinCD
shtml = shtml & " if (document.getElementById('Principal').contentDocument){  "
shtml = shtml & "   xfwin = document.getElementById('Principal');  "
shtml = shtml & "   xfwinCD = xfwin.contentDocument;  "
shtml = shtml & " } else {  "
shtml = shtml & "   xfwin = document.frames['Principal'];  "
shtml = shtml & "   xfwinCD = xfwin.document;  "
shtml = shtml & " }  "

''shtml = shtml & " alert(xfwinCD); "

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

XSBARR1 = "BarraEN.asp?Tipo=" & sTipo & "&MAINAPL=" & SMAINAPL & "&Perfil=" & sPerfil & "&Codigo=" & nCodigo & "&ssclick=" & ssclick & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & "&Funcion=" & replace(nFuncion,"*","&") & "&datos=" & sDatos


sHTML = sHTML & "<script language='JavaScript' for='window' event='onload' >" & chr(13)
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
''sHTML = sHTML & " alert(verfuncion() + 'editardoc');" & chr(13)
'sHTML = sHTML & "     var festil = verfuncion(); " & chr(13)
''sHTML = sHTML & " alert(festil.id);" & chr(13)
'sHTML = sHTML & "     festil.getElementById('dispara').action = '" & XSBARR1 & "'; " & chr(13)
'sHTML = sHTML & "     festil.getElementById('dispara').target = 'Barra'; " & chr(13)
'sHTML = sHTML & "     festil.getElementById('dispara').submit(); " & chr(13)
'sHTML = sHTML & " alert(festil.getElementById('dispara').id);" & chr(13)

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

 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

   IF SMAINAPL = "OR" THEN

      sHTML = sHTML & "OPC - CIRCUITO AUTOEVALUACION"
   ELSE

      sHTML = sHTML & "EPLANNING - PLAN DE CONTINGENCIAS"
   END IF
 ELSE
   sHTML = sHTML & "COR EManager - Edicion de Documentos"
 END IF
   sHTML = sHTML & chr(34) & ";" & chr(13)
end if

sHTML = sHTML & "</script>" & chr(13)





if session("ModMain") = "WK" then
   session("tcliente") = stipo
   session("ncliente") = ncodigo
end if



sHTML = sHTML & "<frameset rows=100,100,*,100,100,100,100 FRAMESPACING=0  FRAMEBORDER=1 BORDER=1 >"

if sACC = "XP" then

shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='estilo2'    ID='estilo2'    sRC='../coress20/estilo.asp'    SCROLLING=no STYLE='xdisplay:none;' >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Barra     ID='Barra' SRC='" & XSBARR1 & "' SCROLLING=no STYLE='xdisplay:none;' > "


else

shtml = shtml & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME='estilo2'    ID='estilo2'    sRC='corbase.htm'    SCROLLING=no STYLE='dxisplay:none;' >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Barra     ID='Barra' SRC='" & XSBARR1 & "' SCROLLING=no STYLE='xdisplay:none;' > "

end if

sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Principal ID='Principal' SRC='CorCharge.asp' SCROLLING=auto > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK1       ID='WK1' SRC='' SCROLLING=no  STYLE='xdisplay:none;' > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK2       ID='WK2' SRC='' SCROLLING=no  STYLE='xdisplay:none;' > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK3       ID='WK3' SRC='' SCROLLING=no  STYLE='xdisplay:none;' > "
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=WK4       ID='WK4' SRC='' SCROLLING=no  STYLE='xdisplay:none;' > "


sHTML = sHTML & "</frameset>"

sHTML = sHTML & "<body  ><head><title>COR EManager - Edicion de Esquemas Normativos</title></head></body>"




sHTML = sHTML & "</html>"

response.write sHTML

%>