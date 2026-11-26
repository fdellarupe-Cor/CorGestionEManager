<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

sElemento = request.querystring("ID")
'''response.write sElemento & "<br>"
n = instr(sElemento,"-")
nCodobjeto = left(sElemento,n-1)
n2 = instr(n+1,sElemento,"-")
sTipobjeto = mid(sElemento,n+1,n2-(n+1))
sEstado = mid(sElemento,n2+1)
sEsqSup = trim(sTipobjeto)
sEsqSup = mid(sEsqSup,len(sEsqSup))
sTipobjeto = left(sTipobjeto,len(sTipobjeto)-1)

SPTIPO = request.querystring("SPTIPO")
SPRCSUP = request.querystring("PRCSUP")

if SPTIPO = "" then SPTIPO = "SP"
if SPRCSUP = "" then SPRCSUP = "SP"


sHTML = "<html xmlns:v='urn:schemas-microsoft-com:vml'  ><title>Incorporar Procesos y Subprocesos</title>"
sHTML = sHTML & "<STYLE>"
sHTML = sHTML & " input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;}"
sHTML = sHTML & " td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;padding-left:20px;}"
sHTML = sHTML & " td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE",Session("CMFPNTIT"),0) & ";}"
sHTML = sHTML & " td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;}"
sHTML = sHTML & " td.cc{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
sHTML = sHTML & "</STYLE>"
sHTML = sHTML & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE><body " & ofv.wbody & " >"


  xs = "<script language=JavaScript>" & chr(13)
  xs = xs & "function AbrirAgregar(Datos) {" & chr(13)
  xs = xs & " window.dialogArguments.open(Datos"
  xs = xs & "," & chr(34) & "edi" & chr(34) & "," & chr(34)
  xs = xs & "menubar=no,toolbar=no" & chr(34)
  xs = xs & " ); " & chr(13)
  xs = xs & "window.close(); " & chr(13)
  xs = xs & " } " & chr(13)
  xs = xs & "</script>" & chr(13)

sHTML = sHTML & xs



if SPTIPO = "CR" then

          sHTML = sHTML & "<script language='javascript' >"
          sHTML = sHTML & "AbrirAgregar('emaddproceso.asp?SPTIPO=CR&Tipo=RGSPCR"
          sHTML = sHTML & "&Nombre=EVENTO-EXPLICITO&Elemento="
          sHTML = sHTML & sElemento & "');"
          sHTML = sHTML & "</script>"

          sHTML = sHTML & "</BODY></html>"

elseif SPTIPO = "AI" then

          sHTML = sHTML & "<script language='javascript' >"
          sHTML = sHTML & "AbrirAgregar('emaddproceso.asp?SPTIPO=AI&Tipo=RGSPAI"
          sHTML = sHTML & "&Nombre=EVENTO-EXPLICITO&Elemento="
          sHTML = sHTML & sElemento & "');"
          sHTML = sHTML & "</script>"

          sHTML = sHTML & "</BODY></html>"


elseif SPTIPO = "PL" then

          sHTML = sHTML & "<script language='javascript' >"
          sHTML = sHTML & "AbrirAgregar('emaddproceso.asp?SPTIPO=PL&Tipo=PLSPCIRC"
          sHTML = sHTML & "&Nombre=SOPORTE-CIRCUITO-CONTINGENCIAS&Elemento="
          sHTML = sHTML & sElemento & "');"
          sHTML = sHTML & "</script>"

          sHTML = sHTML & "</BODY></html>"


else  


set cn = ofv.conectar(ofv.strconn2)

        Strsql = "Select * From Tipobjeto WHERE APLIC = 'RG' "


      '  response.write strsql
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


     If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
        xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
        xGenConta = ofv.FTitDat1

        xGenCont = xGenConta

        xGenConti = "agregar"

        xGenCont = Replace(xGenCont, "XXXDATOXXX", xGenConti)
        xGenCont = Replace(xGenCont, "XXXWIDTHXXX", "98%")
        xGenCont = Replace(xGenCont, "XXXIDXXX", "DBML1R")

        sInput = xGenCont

        xsclase = "xts"

     else

        sInput = "agregar"

        xsclase = "ttr"

     End If




        sHTML = sHTML & "<table width=100% ><tr>"
        sHTML = sHTML & OFV.gencelda("", xsclase, "", "", "", "", "", sInput, "si")
        sHTML = sHTML & "</tr></table>"


        sHTML = sHTML & "<div id='ed2'  style='height:150px;overflow:auto;" & styl & "' >"
        sHTML = sHTML & "<table width=100% >"

       do until rs.eof

          sHTML = sHTML & "<tr>"
          sHTML = sHTML & "<td  class=UT width=15% align='center'  >&nbsp</td>"
          sHTML = sHTML & "<td  class=UT width=70% align='center'  ><BR>"
          sHTML = sHTML & "<p align='justify'  >Confirma la incorporacion de un nuevo proceso/subproceso?</p>"
          sHTML = sHTML & "</td>"
          sHTML = sHTML & "<td  class=UT width=15% align='center'  >&nbsp</td>"
          sHTML = sHTML & "</tr>"

          sHTML = sHTML & "<tr>"
          sHTML = sHTML & "<td colspan=3 id='TD" & rs(1) & "' align='center' class=UT  STYLE='cursor:pointer;border-style:none;border-width:1;border-color:#f7aa22;' "
          sHTML = sHTML & " onmouseover=""this.style.borderStyle='solid';this.className='vti';"" onmouseout=""this.style.borderStyle='none';this.className='uti';"" "
          sHTML = sHTML & " onclick=" & chr(34)
          sHTML = sHTML & "AbrirAgregar('" & "emaddproceso.asp?SPTIPO=" & SPRCSUP & "&Tipo=" & RS(0)
          sHTML = sHTML & "&Nombre=" & ofv.convertircar(rs(1)," ","_") & "&Elemento="
          sHTML = sHTML & sElemento & "');" & chr(34)
          sHTML = sHTML & " ><BR><BR><BR><input type='button' value='CREAR' class=bt> " & chr(13)
          sHTML = sHTML & "</td>"
          sHTML = sHTML & "</tr>"
         exit do
       loop



'       do until rs.eof
'           sxiconv = "<font  face='corflaticon' size=4 color='#990000'>&#0113;</font>"
'          sHTML = sHTML & "<tr><td width=5%  class=UTS >" & sxiconv
'          sHTML = sHTML & "</td>"
'          sHTML = sHTML & "<td id='TD" & rs(1) & "' class=UTI width=95% STYLE='cursor:pointer;border-style:none;border-width:1;border-color:#f7aa22;' "
'          sHTML = sHTML & " onmouseover=""this.style.borderStyle='solid';this.className='vti';"" onmouseout=""this.style.borderStyle='none';this.className='uti';"" "
'          sHTML = sHTML & " onclick=" & chr(34)
'          sHTML = sHTML & "AbrirAgregar('" & "emaddproceso.asp?SPTIPO=" & SPRCSUP & "&Tipo=" & RS(0)
'          sHTML = sHTML & "&Nombre=" & ofv.convertircar(rs(1)," ","_") & "&Elemento="
'          sHTML = sHTML & sElemento & "');" & chr(34)
'          sHTML = sHTML & " >" & rs(1) & chr(13)
'          sHTML = sHTML & "</td></tr>"
'          rs.movenext
'        loop



        sHTML = sHTML & "</table>"
        sHTML = sHTML & "</div>"
        sHTML = sHTML & "<hr size=1 color=#000099>"
        call ofv.cerrarconsulta(rs)


call ofv.cerrarconn(cn)





sHTML = sHTML & "</BODY></html>"


end if


response.write sHTML

%>