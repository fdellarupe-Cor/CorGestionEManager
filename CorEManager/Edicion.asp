<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

smodal = request.querystring("modal")

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

ofv.ObtenerAtributos Session("Form2"),""
bAgregar   = ofv.AAgregar
bModificar = ofv.AModificar
bSuprimir  = ofv.AEliminar
bConsultar = ofv.AConsultar

set cn = ofv.conectar(ofv.strconn2)

scodigo = nCodobjeto
stipo   = sTipobjeto
strsql = "SELECT tipocpt FROM tipobjeto  where tipobjeto ='" & stipo & "'"


 set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 if not rs.eof then
  tipoPedido=rs(0)
 else
  tipoPedido=""
 end if
 call ofv.cerrarconsulta(rs)

srestric = "N"


'xxxxxxxxxxxxxxxxxxxxxxxx

if sEsqSup <> "Y" then
    redim preserve sPath(0)
    i = 1
    do

       strsql = "select * from csbarr where codojasoc = " & sCodigo

       strsql = strsql & " and tipobjasoc = '" & sTipo & "'"
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if not rs.eof then
        redim preserve sPath(i)
        sCodigo = rs(1)
        sTipo = rs(2)
        if not isnull(isnull(rs(0))) then
          sPath(i) = rs(0)
        else
          i = i + 1
          exit do
        end if
        i = i + 1
      else
        exit do
      end if
      call ofv.cerrarconsulta(rs)
    loop

end if

    strsql = "SELECT * FROM PRODSERV where codprodserv = " & sCodigo

    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if not rs.eof then
       if rs("segaceso") <> 0 or rs("segactual") <> 0 then srestric = "S"

    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
            StrSql = "Select * From subtemaskr "
            StrSql = StrSql & "Where codprodserv = " & sCodigo & " and tiposeg = 'T' "
            StrSql = StrSql & " and (grpusr = 'USR-" & Ucase(session("usuario")) & "' "

'''''''''''''''''''''''''''''''''''''''''''
            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = ""
                StrSql = StrSql & " or grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = ","
                next

                StrSql = StrSql & ") "

            else
                StrSql = StrSql & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
            end if


''''''''''''''''''''''''''''''''''''''''''

            StrSql = StrSql & ")"

       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs.eof then
          srestric = "N"
       else
          srestric = "S"
       end if
       call ofv.cerrarconsulta(rs)
    end if



if srestric   = "S" then
   bAgregar   = 0
   bModificar = 0
   bSuprimir  = 0
end if


'xxxxxxxxxxxxxxxxxxxxxxxx

'v TEMAS         CLIMS2751557214332898409701  CLIMS28915558161334435117553 FORMULARIOS VIGENTES CLIMS46712740044434860617908
'd DESARROLLO    CLIMS2761557214333698409701  CLIMS29015558161341435117553 FORMULARIOS DESARROLLO CLIMS4681274004444386061790
'a PARA APROBAR  CLIMS2771557214334498409701  CLIMS29115558161348435117553 FORMULARIOS PARA APROBAR CLIMS48112739033140999275609
'n ELIMINADAS    CLIMS2781557214335198409701  CLIMS29215558161356435117553 FORMULARIOS ELIMINADAS CLIMS48212739033350999275609




session("Nomodificar") = "N"

if smodal <> "PROCESOS" then


if isnumeric(sEstado) then
'and  Session("Form2") <> "DESARROLLO"
if Session("Form2") <> "CLIMS46712740044434860617908" AND Session("Form2") <> "TEMAS" And Session("Form2") <> "CLIMS2751557214332898409701" And Session("Form2") <> "CLIMS28915558161334435117553" And  sEstado = 3 then
  bSuprimir = False
  session("Nomodificar") = "S"
end if
if  Session("Form2") <> "CLIMS4681274004444386061790" AND Session("Form2") <> "DESARROLLO"    And Session("Form2") <> "CLIMS2761557214333698409701" And Session("Form2") <> "CLIMS29015558161341435117553" And sEstado = 1 then
  bSuprimir = False

  session("Nomodificar") = "S"
end if
if Session("Form2") <> "CLIMS48112739033140999275609" AND Session("Form2") <> "PARA APROBAR"  And Session("Form2") <> "CLIMS2771557214334498409701" And Session("Form2") <> "CLIMS29115558161348435117553" And sEstado = 2 then
  bSuprimir = False
  session("Nomodificar") = "S"
end if
end if

if Session("Form2") = "CLIMS48212739033350999275609" OR Session("Form2") = "TEMAS" or Session("Form2") = "ELIMINADAS" or Session("Form2") = "CLIMS2751557214332898409701" or Session("Form2") = "CLIMS2781557214335198409701" or Session("Form2") = "CLIMS28915558161334435117553"  or Session("Form2") = "CLIMS29215558161356435117553" then
  bAgregar = False
  'session("Nomodificar") = "S"
end if

else

  bSuprimir = False
  bAgregar  = False

end if




session("EsqSup") = sEsqSup



Session("Segct") = "S"
Session("Segmo") = "S"
Session("Segco") = "S"
Session("Segce") = "S"

if sEsqSup = "Y" then
   xtipo = "SUBTEMAS"
    strsql = "SELECT DESCRI FROM PRODSERV where codprodserv = " & NCODOBJETO
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rs.eof then sNombredOC = rs(0)


  Strsql = "SELECT SUBTEMA.* FROM PRODSERV, SUBTEMA where "
  Strsql = Strsql & "PRODSERV.marcaps = SUBTEMA.tipobjeto and "
  Strsql = Strsql & " prodserv.marcaps = '" & sTipobjeto & "'"
else
   xtipo = "ESQUEMAS"
   strsql = "SELECT DESCRI,objun FROM OBJETOS where CODOBJETO = " & NCODOBJETO
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   PuedeCopiarDoc="S"
   if not rs.eof then
    if rs("objun")=1 and session("UN") <> "1" then PuedeCopiarDoc="N"
    sNombredOC = rs(0)
   end if

  Strsql = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"
end if
'ss=strsql
set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if sEsqSup <> "Y" then
  TIPOCPT=rs("TIPOCPT")
else
TIPOCPT=""
end if

x = 0
redim Niveles(1)
if not rs.eof then
  sNombre  = rs(1)
  sTIPOdOC = rs(1)
  sPerfil = rs(3)
  for z = 5 to 14
    if rs(z) Then
      x = x + 1
      redim preserve Niveles(x)
      Niveles(x) = right(rs(z).name,1)
    end if
  next
end if

call ofv.cerrarconsulta(rs)


sHTML = "<html xmlns:v='urn:schemas-microsoft-com:vml'  ><title>" & sNombre & "</title>"
sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css""><STYLE>"
sHTML = sHTML & " input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";transition: all 0.2s ease-in-out;border-radius:15px;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;}"
sHTML = sHTML & " input.bt:hover{opacity: 0.8;}"
sHTML = sHTML & " td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:10px;font-weight:600;padding-left:20px;}"
sHTML = sHTML & " td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:5px;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE",Session("CMFPNTIT"),0) & ";}"
sHTML = sHTML & " td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:10px;font-weight:500;padding-left:20px;}"
sHTML = sHTML & " td.cc{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
sHTML = sHTML & "</STYLE>"
sHTML = sHTML & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE><body " & ofv.wbody & " >"


  xs = "<script language=JavaScript>" & chr(13)
  xs = xs & "function AbrirAgregar(Datos) {" & chr(13)
  xs = xs & " var h=570;w=790 ;"
  xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;}"
   if BrowserName()="IE" then
        xs = xs & " window.dialogArguments.open(Datos"
       xs = xs & "," & chr(34) & "edi" & chr(34) & "," '& chr(34)
       'xs = xs & "resizable=yes,width=500,height=500,menubar=no,toolbar=no" & chr(34)
       xs = xs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
           xs = xs & " ); " & chr(13)
       'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
       else
        xs = xs & " window.opener.open(Datos"
       xs = xs & "," & chr(34) & "edi" & chr(34) & "," '& chr(34)
     '  xs = xs & "resizable=yes,width=500,height=500,menubar=no,toolbar=no" & chr(34)
       xs = xs & "'top=0,left=0,menubar=no,toolbar=no,height='+ h +',width='+ w +',resizable=yes,status=yes'" '& Chr(34)
        xs = xs & " ); " & chr(13)
    end if

 ' xs = xs & " window.dialogArguments.open(Datos"
 ' xs = xs & "," & chr(34) & "edi" & chr(34) & "," & chr(34)
  'xs = xs & "resizable=yes,width=500,height=500,menubar=no,toolbar=no" & chr(34)
  'xs = xs & " ); " & chr(13)
  xs = xs & "window.close(); " & chr(13)
  xs = xs & " } " & chr(13)
  xs = xs & "</script>" & chr(13)

sHTML = sHTML & xs

' sHTML = sHTML & "<table width=100% ><tr><td class=cc height=18 >"
' sHTML = sHTML & "<font color=#336699 face=verdana size=1>"
' sHTML = sHTML & "<b>-" & Session("Form2") & " " &  sEstado
' sHTML = sHTML & "<br>-" & Session("ASP")
'sHTML = sHTML & "<br>-" & session("Nomodificar")
' sHTML = sHTML & "<br>-" & smodal
' sHTML = sHTML & "</b></font></td></tr>"
' sHTML = sHTML & "</table>"

Session("SegRst") = "N"

if sEsqSup = "N" then

   set cn = ofv.conectar(ofv.strconn2)

   Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipobjeto & "' and codobjeto = "  & nCodobjeto
  ' response.write strsql
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rs.eof then


      if Session("Usuario") <> rs(7) then

         if rs(3) then Session("Segct") = "N"
         if rs(4) then Session("Segmo") = "N"
         if rs(5) then Session("Segco") = "N"
         if rs(6) then Session("Segce") = "N"

         if rs(3) or rs(4) or rs(6)   then Session("SegRst") = "S"


         call ofv.cerrarconsulta(rs)
         Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipobjeto
         Strsql = Strsql & "' and codobjeto = "  & nCodobjeto
         Strsql = Strsql & " and ("
         Strsql = Strsql & "(grpusr = 'USR-" & Session("Usuario") & "') "

'''''''''''''''''''''''''''''''''''''''''''
            if  Session("DELLOG") = "S" then

                XMFTGRP = split(session("SegGrpLst"),chr(9))
                xcount  = ubound(XMFTGRP) + 1
                XMFSEP  = ""
                StrSql = StrSql & " or ( grpusr in ("

                for XMFI = 0 to ubound(XMFTGRP)
                    StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
                    XMFSEP  = ","
                next

                StrSql = StrSql & ")) "

            else
                Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
            end if


''''''''''''''''''''''''''''''''''''''''''


         Strsql = Strsql & ") "
        ' response.write strsql
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         do until rs.eof

            select case rs(3)
                   case "1"
                        Session("Segct") = "S"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

                   case "2"
                        Session("Segmo") = "S"
                        Session("Segco") = "S"

                   case "3"
                        Session("Segco") = "S"

                   case "4"
                        Session("Segco") = "S"
                        Session("Segce") = "S"

            end select
         rs.movenext
         loop
      end if
    end if
    call ofv.cerrarconsulta(rs)

    if srestric = "S" then
       Session("Segct") = "N"
       Session("Segmo") = "N"
       Session("Segce") = "N"
    end if


    if Session("Segce") = "S" and session("Nomodificar") <> "S" then
       session("Nomodificarce") = "N"
    else
       session("Nomodificarce") = "S"
    end if
'response.write  SESSION("WMATRMEMP") & "           " & SESSION("WMATRMEMPACT")
    if Session("Segct") <> "S" and Session("Segmo") <> "S" then
       session("Nomodificar") = "S"
    end if

    BCOPYSPC = "N"

IF   SESSION("WMATRMEMP") = "S" AND  SESSION("WMATRMEMPACT") = "N" THEN


   bAgregar = 0
   bSuprimir = 0
   session("Nomodificar") = "S"
   session("Nomodificarce") = "S"

   IF bModificar and Session("Segct") = "S" THEN
      BCOPYSPC = "S"
   END IF

      bModificar = 0

      Session("Segct") = "N"
      Session("Segmo") = "N"
      Session("Segce") = "N"

END IF


  if smodal <> "PROCESOS" then

     sHTML = sHTML & "<table width=100% cellspacing=0 cellpadding=4 ><tr >"

     aspeditar   = cstr("Editar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil  & "&tipocpt=" &   TIPOCPT)
     aspeliminar = cstr("asp/EliminarEsquema.asp?Codobjeto=" & nCodobjeto & "&Tipobjeto=" & sTipobjeto & "&Eliminar=NSNC")
     aspcopiar   = cstr("Copiar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil)



	 '************************************************************
	 '*********************** INC 457 INICIO *********************
	 '************************************************************


	 '1: Se agregan las opciones para eliminar PD y SN EN estado 1

	 aspeliminarpedidosPD1 = cstr("CrearPedido.asp?PEDIDO=ELIM&idpedido=" & nCodobjeto)
	 aspeliminarpedidosSN1 = cstr("CrearNormativa.asp?PEDIDO=ELIM&idpedido=" & nCodobjeto)

	 'Borra el pedido de las tablas:
	 'Pedidos
	 'Pedidosdet
	 'Pedidosarch
	 'Relaobj


	 '2: Se agregan las opciones para eliminar PD y SN EN estado 2

	 aspeliminarpedidosPD2 = cstr("CrearPedido.asp?PEDIDO=ELIM2&idpedido=" & nCodobjeto)
	 aspeliminarpedidosSN2 = cstr("CrearNormativa.asp?PEDIDO=ELIM2&idpedido=" & nCodobjeto)


	 'Cambia el estado a BAJA
	 'Mueve el archivo a la carpeta de ELIMINADOS
	 'Completa todos los pasos del WKF incluyendo el paso final 999

	 'Se agregan las acciones que ejecuta el boton Eliminar

	 'Estado 1
	 IF (TIPOPEDIDO = "PD") AND (sEstado = 1) then sHTML = sHTML & Eliminar(cstr(aspeliminarpedidosPD1))
	 IF (TIPOPEDIDO = "SN") AND (sEstado = 1) then sHTML = sHTML & Eliminar(cstr(aspeliminarpedidosSN1))

	 'Estado 2
	 IF (TIPOPEDIDO = "PD") AND (sEstado = 2) then sHTML = sHTML & Eliminar(cstr(aspeliminarpedidosPD2))
	 IF (TIPOPEDIDO = "SN") AND (sEstado = 2) then sHTML = sHTML & Eliminar(cstr(aspeliminarpedidosSN2))


	 '**********************************************************
	 '********************** INC 457 FIN ***********************
	 '**********************************************************




'	if bSuprimir = "S" and Sestado = "1" and tipopedido = "PD" or tipopedido = "SN"   then sHTML = sHTML & Eliminar(cstr(aspeliminar))

     if bConsultar and Session("Segco") = "S" then sHTML = sHTML & Editar(cstr(aspeditar)) '& aspeliminar
	 if bSuprimir and Session("Segct") = "S" AND TIPOPEDIDO <> "PD" AND TIPOPEDIDO <> "SN" then sHTML = sHTML & Eliminar(cstr(aspeliminar))



	if ((bModificar and Session("Segct") = "S") OR (BCOPYSPC = "S")) AND TIPOPEDIDO <> "PD" AND TIPOPEDIDO <> "SN" and PuedeCopiarDoc= "S" then

         sHTML = sHTML & Copiar(cstr(aspcopiar))

     END IF

 '    if (bModificar and Session("Segct") = "S") OR (BCOPYSPC = "S") then
'
'	          sHTML = sHTML & MOVER(cstr(aspmover))
'
 '    END IF


     sHTML = sHTML & "</tr></table>"

  end if

end if

'response.write  Niveles(1)
if Niveles(1) <> ""  AND TIPOPEDIDO <> "PD" AND TIPOPEDIDO <> "SN" then
      if bAgregar then
        IF SESSION("BASE")="K" THEN
         Strsql = "Select * From Tipobjeto Where TIPOCPT='K' AND (aplic = 'GLB' or aplic = '" & wkaplic & "' "
        ELSE
         Strsql = "Select * From Tipobjeto Where TIPOCPT <> 'K' AND (aplic = 'GLB' or aplic = '" & wkaplic & "' "
        END IF
          Strsql = Strsql & ") and nivel IN("
        for x = 1 to ubound(Niveles)
          Strsql = Strsql & "'" & Niveles(x) & "',"
        next
        Strsql = left(Strsql,len(Strsql)-1) & ")"

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


        sHTML = sHTML & "<div id='ed2'  style='height:150px;overflow:auto;" & styl & "' ><span id=111>"
        sHTML = sHTML & "<table width=100% >"

        do until rs.eof


           if  RS(21) > "C" then
               sxiconv = "<i class='las la-file-alt' style=color:"& session("CMFPNTIT") &";></i>"
           else
               'sxiconv = "<font  face='CorFlatIcon' size=4 color='#336699'>&#0204;</font>"
			   sxiconv = "<i  class='las la-folder' style=color:"& session("CMFPNTIT") &"></i>"
           end if


          sHTML = sHTML & "<tr><td width=5%  class=UTS >" & sxiconv
          sHTML = sHTML & "</td>"
          sHTML = sHTML & "<td id='TD" & rs(1) & "' class=UTI width=95% STYLE='cursor:pointer;border-style:none;border-width:1;border-color:" & Session("CMFPNBOT") & ";' "
          sHTML = sHTML & " onmouseover=""this.style.borderStyle='solid';this.className='vti';"" onmouseout=""this.style.borderStyle='none';this.className='uti';"" "
          sHTML = sHTML & " onclick=" & chr(34)
          sHTML = sHTML & "AbrirAgregar('" & "Agregar.asp?Tipo=" & rs(0)
          sHTML = sHTML & "&Nombre=" & ofv.convertircar(rs(1)," ","_") & "&Elemento="
          sHTML = sHTML & sElemento & "');" & chr(34)
          sHTML = sHTML & " >" & rs(1) & chr(13)
          sHTML = sHTML & "</td></tr>"


          rs.movenext
        loop
        sHTML = sHTML & "</table>"
        sHTML = sHTML & "</span></div>"
        sHTML = sHTML & "<hr size=1 color=#000099>"
        call ofv.cerrarconsulta(rs)
      end if
end if

  sHTML = sHTML & "<table width=100% cellspacing=0 cellpadding=4 ><tr >"

  if smodal <> "PROCESOS" then

  if bModificar AND TIPOPEDIDO <> "PD" AND TIPOPEDIDO <> "SN" then
     xtipo = "ESQUEMAS"
     SNOMDOC = STIPODOC & ":" & SNOMBREDOC
     SNOMDOC = REPLACE(SNOMDOC," ","_")
     asporden    = cstr("edordenar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&XTipo=" & xtipo  & "&Perfil=" & sPerfil & "&NOMDOC=" & SNOMDOC)
     sHTML = sHTML & Ordenar(cstr(asporden))
  END IF

   if ((bModificar and Session("Segct") = "S") OR (BCOPYSPC = "S")) AND TIPOPEDIDO <> "PD" AND TIPOPEDIDO <> "SN" then
   if isnumeric(sEstado) then
   if sEstado <> 4   or session("APLMDCLI")="CFA" or  session("APLMDCLI")="BC" then
    aspmover  =   cstr("Copiar.asp?EsqSup="&sEsqSup&"&MOVERESQUEMA=SI&Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil)
    sHTML = sHTML & MOVER(cstr(aspmover))
    end if
   end if
    END IF


  else

  if bModificar then
     xtipo = "ESQUEMAS"
     SNOMDOC = STIPODOC & ":" & SNOMBREDOC
     SNOMDOC = REPLACE(SNOMDOC," ","_")
     aspprocesos    = cstr("edprocesos.asp?codobjeto=" & nCodobjeto & "&Tipobjeto=" & sTipobjeto)
     sHTML = sHTML & Aprocesos(cstr(aspprocesos))
  END IF







  end if

  sHTML = sHTML & "</tr></table>"

call ofv.cerrarconn(cn)

sHTML = sHTML & "</BODY></html>"

response.write sHTML
'session("Nomodificar")="N"
'response.write Session("Form2") & TIPOPEDIDO & ("&nbsp;") & sEstado & bSuprimir
'response.write session("segckt")
%>