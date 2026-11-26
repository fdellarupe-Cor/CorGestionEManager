<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 8
  CantRegMoverx = cdbl("0" & request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

   XXXX = NEWEMPRE("NO")

   '   session("ASP")="SUBCONSULTA.ASP"

      sUN = request.querystring("EMPRE")
    'response.write "sEMPRE: " & sEMPRE
       IF sUN = ""  AND SESSION("UN") = "" THEN

         sUN = ATREMP(0, 0)
         SESSION("UN") = sUN & ""
         If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
       elseIF sUN <> "" THEN
        XXXX = NEWEMPRE(sUN)
        SESSION("UN")=sUN
       ELSE
        XXXX = NEWEMPRE(SESSION("UN"))
        sUN=SESSION("UN")
   end if



 TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


shtml = shtml & tablah


sUN = session("UN")

'response.write sUN
  'Encabezados
  Fila = ofv.Encabezados("Codigo","imp","vigencia","mensaje","dias","dist","env","&nbsp;","lista","envio")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sql="select cliente from CLIENTEPLAN order by cliente"
  set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
  clientes=""
  delim=""
  Do While (Not rs.EOF)
   clientes= clientes & delim & rs(0)
   delim="\r"
  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)

  Strsql = "Select * From novegenpar where codemp=0" & sUN & " and MSGTIPOMESS = '1' Order By enviado, objid desc"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi< CantRegMostrarx)
    Formulario = ""
    sForm = "form" & posi

    if rs(8) = "S" then clase = "cc" else clase = "cc" end if

    'Eliminar
   if bSuprimir then
     sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
     sHref = sHref & "&Tabla=novegenpar&volver=" & vuelta & "&Eliminar=NSNC"
     Celda = ofv.BotonEliminar("Mensaje: NGP" & cstr(rs(0)),sHref)
  else
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
  end if

    'codigo
    sInput = "<b>NGP" & cstr(rs(0)) & "</b>"
    Celda = Celda & ofv.GenCelda("",clase & " onclick=" & chr(34) & "document.getElementById('VNGP" & cstr(rs(0)) & "').submit();WKlevanta('NGP" & cstr(rs(0)) & "');" & chr(34) & " ","","5%","","","",sInput,"si")


    xSolasp = ""
    xSolasp = xSolasp & "novedmenuview.asp?xxmsg=NGP" & cstr(rs(0))

    xsHTML = ""
    xsHTML = xsHTML & "<form name='VNGP" & cstr(rs(0)) & "' action='" & xSolasp  & "' "
    xsHTML = xsHTML & " method=post target='fesquemas' style='display:none;' >"
    xsHTML = xsHTML & "</form>"



    'Importancia

   select case clng(rs(5))
	  case 0

		sxopc4x = "<font color=#800000 size=4><b>!</b></font>"
	  case 1

		sxopc4x = "<font color=#008000 size=4><b>*</b></font>"
	  case 2

		sxopc4x = "<font color=#000080 size=4><b>+</b></font>"

    end select

    sInput = sxopc4x
    Celda = Celda & ofv.GenCelda("",clase,"center","5%","","","",sInput,"si")

    'Fecha Vigencia
    '''  sInput = cdate(clng(rs(1)) + clng(cdate("01-01-1900")))

   '   sInput = Fechador(rs(1) )
      sInput = Fechador(rs(1) + clng(cdate("01-01-1900")) )
    Celda = Celda & ofv.GenCelda("",clase,"","30%","","","","<b>" & sInput   & "</b>","si")

    'Mensaje
    'if len(rs(4)) > 30 then
    '   sInput = left(rs(4),30) & "..."
    'else
    '   sInput = rs(4)
    'end if
    sinput= RemoveHTML(rs(4))
    Celda = Celda & ofv.GenCelda("",clase,"","40%","","","",sInput,"si")

    'dias
    sInput = rs(3)
    Celda = Celda & ofv.GenCelda("",clase,"center","5%","","","",sInput,"si")''''"",clase,"right",

    'lista
    if cint(rs(6)) = 1 then
       sInput = "TODOS"
    else
       sInput = "LISTA"
    end if

    Celda = Celda & ofv.GenCelda("",clase,"","5%","","","",sInput,"si")

    'enviado
    if rs(8) = "S" then
       sInput = "<font color=#008000><b>SI</b></font>"
    else
       sInput = "<font color=#800000><b>NO</b></font>"
    end if
    Celda = Celda & ofv.GenCelda("",clase,"","5%","","","",sInput,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=novegenpar&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")

if session("APLMDCLI") <> "GP" THEN
sql = "select contacto from novegenpar where objid=" & rs(0)
set rt = ofv.crearconsultaex(sql,cn,1,parametros)
puedeEdit=""
if rt(0) = session("Usuario") then
puedeEdit=true
else
puedeEdit=false
end if

ELSE

'sql = "select contacto from novegenpar where objid=" & rs(0)
sql = "select codemp from novegenpar where objid=" & rs(0)
set rt = ofv.crearconsultaex(sql,cn,1,parametros)
puedeEdit=""
'if rt(0) = session("Usuario") then
if rt(0) & "" = session("CLIEMP") & "" then
  puedeEdit=true
else
  puedeEdit=false
end if
END IF



    if rs(8) = "S" and puedeEdit then

       sInput = ofv.GenerarInput("","EDITAR","submit","","10","bt","")
       Celda = ofv.GenCelda("","","center","10%","","","",sInput,"si")
       sAccion = "AdnovedadesProp.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0) & "&Editar=1"
       Formulario = Formulario & ofv.GenForm("Boton" & rs(0),"",sAccion,"","post","",Celda,"si")

    '   sInput = "----"
    '   Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")
    else



       sInput = ofv.GenerarInput("","PROPIEDADES","submit","","10","bt","")
       Celda = ofv.GenCelda("","","center","10%","","","",sInput,"si")

       sAccion = "AdnovedadesProp.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0) & "&Editar=1"
       Formulario = Formulario & ofv.GenForm("Boton" & rs(0),"",sAccion,"","post","",Celda,"si")

    end if



    if cint(rs(6)) = 2 then
       sInput = ofv.GenerarInput("","LISTA","submit","","10","bt","")
       Celda = ofv.GenCelda("","","center","10%","","","",sInput,"si")
       sAccion = "Adnovedadesdist.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0)
       Formulario = Formulario & ofv.GenForm("Boton" & rs(0),"",sAccion,"","post","",Celda,"si")
    else
       sInput = "---"
       Celda = ofv.GenCelda("",clase,"center","10%","","","",sInput,"si")
       Formulario = Formulario & Celda
    end if
'response.write Session("EMPMUN")  & "  " &   rs(6)
    if rs(8) <> "S" then
      ' sInput = ofv.GenerarInput("","ENVIAR","submit","","10","bt","")
        if Session("EMPMUN") = "S" and  cint(rs(6)) = 1 then
	           wktit="SE VA A ENVIAR ESTE ESQUEMA A LAS SIGUIENTES EMPRESAS\r\r" & clientes & "\r\rSi desea modificar esta selección ingresar a Lista de Distribución "
	           sInput = ofv.GenerarInput("","ENVIAR","button","","10","bt onclick=""disp_confirm('" & wktit & "','" & "Botonw" & rs(0) & "'); """ ,"")
	         else
	          sInput = ofv.GenerarInput("","ENVIAR","submit","","10","bt" ,"")
       end if

'       if cint(rs(6)) = 2 then
'        sInput = ofv.GenerarInput("","ENVIAR","submit","","10","bt" ,"")
'       else
'        if  cint(rs(6)) = 1 and Session("EMPMUN") = "S" then
'         wktit="SE VA A ENVIAR ESTE ESQUEMA A LAS SIGUIENTES EMPRESAS\r\r" & clientes & "\r\rSi desea modificar esta selección ingresar a Lista de Distribución "
'         sInput = ofv.GenerarInput("","ENVIAR","button","","10","bt onclick=""disp_confirm('" & wktit & "','" & "Boton" & rs(0) & "'); """ ,"")
'        else
'        end if
'       end if

if session("APLMDCLI") <> "GP" THEN
'MAURO 14/08/17
sql = "select contacto, enviado from novegenpar where objid=" & rs(0)
set rt = ofv.crearconsultaex(sql,cn,1,parametros)
puedeEdit=""
if rt(0) = session("Usuario") then
puedeEdit=true
else
puedeEdit=false
end if

ELSE
'sql = "select contacto, enviado from novegenpar where objid=" & rs(0)

sql = "select codemp, enviado from novegenpar where objid=" & rs(0)
set rt = ofv.crearconsultaex(sql,cn,1,parametros)
puedeEdit=""
'if rt(0) = session("Usuario") then
if rt(0) & "" = session("CLIEMP") & "" then
puedeEdit=true
else
puedeEdit=false
end if

END IF

if puedeEdit then
       Celda = ofv.GenCelda("","","center","10%","","","",sInput,"si")

       sAccion = "AdnovedadesDESP.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0)
       Formulario = Formulario & ofv.GenForm("Botonw" & rs(0),"",sAccion,"","post","",Celda,"si")
    else
       sInput = "---"
       Celda = ofv.GenCelda("",clase,"center","10%","","","",sInput,"si")
       Formulario = Formulario & Celda
    end if
     end if

    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si") & xsHTML

    posi=posi+1
    rs.Movenext
  Loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","90%","","0","","0","0","","",filas,"si")




  Agregar = ""
  if bAgregar then
    sAccion = "AgregarNovedad.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=novegenpar"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
 ' str1 = "Select descri From tema Order By descri "    ' orden pagina
 ' str2 = "Select descri From tema Order By descri "    ' orden alfabetico
  'Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
  buscar = ""
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  wstyl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"
  wStyL = wStyL & "border-style:solid;border-width:1;border-color:#000000;background-color:#eeeeee;"

  sHTML = sHTML &  "<div id='mt2view'  style='display:none;position:absolute;top:200;left:300;height:50%;width:50%;overflow:auto;" & wstyl & "' >"

  sHTML = sHTML &  "<table width=99% >"
  sHTML = sHTML &  "<tr height=10 >"
  sHTML = sHTML &  "<td id='mt2viewl1c1' class='tt' width=90% style='border-style:solid;border-width:1;border-color:#000000;padding-left:5px;background-color:" & Session("CMFMENU") & "' >Mostrando vista de </td>"
  sHTML = sHTML &  "<td id='mt2viewl1c2'  class='uT' align=center onclick=" & chr(34) & "document.getElementById('mt2view').style.display='none';" & chr(34) & " style='cursor:pointer;border-style:outset;border-width:1;border-color:#000000;background-color:#eeeeee;' ><h5>Cerrar</h5></td>"
  sHTML = sHTML &  "</tr>"

  sHTML = sHTML &  "<tr height=100% >"
  sHTML = sHTML &  "<td id='DX4' colspan=2  class='uT' style='background-color:#ffffff;' >contenido</td>"
  sHTML = sHTML &  "</tr>"
  sHTML = sHTML &  "</table>"

  sHTML = sHTML &  "</div >"

 shtml = shtml & "<script language=javascript > "
 shtml = shtml & "function WKlevanta(wkvista) { " & chr(13)
' shtml = shtml & "  var frmna = 'V' + wkvista; "
' shtml = shtml & " alert(frmna); " & chr(13)
' shtml = shtml & "  document.all.[frmna].submit(); " & chr(13)
 shtml = shtml & "  document.getElementById('mt2viewl1c1').innerHTML='<h5>Mostrando vista de ' + wkvista + '</h5>';" & chr(13)
 shtml = shtml & "  document.getElementById('mt2view').style.display='block';   " & chr(13)
 shtml = shtml & " } " & chr(13)

     sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
     sHTML = sHTML & "  {"
     sHTML = sHTML & "  var r=confirm(disp_msg);"
     sHTML = sHTML & "  if (r==true)"
     sHTML = sHTML & "    {"
     sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
     'sHTML = sHTML & "    alert(fr.action);"
     sHTML = sHTML & "    fr.submit();"
     sHTML = sHTML & "    }"
     sHTML = sHTML & "  else"
     sHTML = sHTML & "    {"
     sHTML = sHTML & "    alert('Se cancelo la accion!');"
     sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

 shtml = shtml & "</script> "

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Function Fechador(wkfecha)


      wkfecha=clng(wkfecha)
      fec = cdate(wkfecha)
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

Function RemoveHTML( strText )
	Dim RegEx

	Set RegEx = New RegExp

	RegEx.Pattern = "<[^>]*>"
	RegEx.Global = True

	RemoveHTML = RegEx.Replace(strText, "")
End Function


function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

   Fila    = ""

   SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

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
   TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


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

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 		sHref = "adnovedades.asp?EMPRE=" & ATREMP(IAEMP,0)
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

end function

%>