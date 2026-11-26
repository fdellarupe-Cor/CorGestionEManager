<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nblk = request.querystring("nblk")
  if len(nblk) = 0 then nblk = "D"

  ord = request.querystring("ord")
  if len(ord) = 0 then ord = "T"

  nAccion = request.querystring("Accion")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vueltaC = request.querystring("volver")
  if session("volverproc") = "" then session("volverproc") = vueltaC
  volver2 = ofv.vueltaasp2("../")

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "


  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * From objaccion where objid = " & nAccion
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     ccomp = rs(4)
     ccompon = rs(6)
  end if


  if nblk = "S" then ccompS = request.form("comp1")
  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS

  if nblk = "S" and clng(ccomp2) = 0 then
     ccomp2 = ccomp
     ccompS = ccomp
  end if

   sHTML = ofv.menuheader("","")
    '----------------------------------------------

      shtml = shtml & CorPnlFnc("N","")

      sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=1111>"

'-------------------------------------------------

  'sHTML = sHTML & "<center><br>"


  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=ttt align=center >"
  sHTML = sHTML & "Acciones&nbsp;disponibles "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center  >"
'-------------
  if bModificar or true then
      sAccion = "EdProcedACC.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=0 "
  if (clng(ccomp) = 0  and nblk = "D") or clng(ccomp2) = 0 then
     sHTML = sHTML & " selected "
  end if
  sHTML = sHTML & " >---------- Sin Asociar ----------</option>"

  if ord = "T" then
     Strsql = "SELECT tipoaccionalt,descri From TIPOACCALT where estado=3 Order By descri"
  else
     ccomp3 = request.form("comp3")
     if len(ccomp3) = 0 then
        Strsql = "SELECT tipoaccionalt,descri From TIPOACCALT where estado=3 Order By descri"
     else
        Strsql = "SELECT tipoaccionalt,descri From TIPOACCALT  where  estado=3  and modo = '" & ccomp3 & "' Order By descri"
     end if
  end if
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof
  sHTML = sHTML & "<option value=" & rs2(0)
  if  cstr(rs2(0)) = cstr(ccomp2) then   sHTML = sHTML & " selected "
  sHTML = sHTML & " > " & rs2(1) & "</option>"

  rs2.movenext
  loop

  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center class=ttt  >"
  sHTML = sHTML & "FILTRO: "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"

  sHTML = sHTML & "<td  align=center >"
'-------------
  if bModificar or true then
      sAccion = "EdProcedACC.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=D&ord=F&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select  name=comp3 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp3.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=TODOS "
  if (clng(ccomp) = 0  and nblk = "D") or clng(ccomp2) = 0 then
     sHTML = sHTML & " selected "
  end if
  sHTML = sHTML & " >---------- TODOS ----------</option>"

  Strsql = "SELECT  tipomodo, descri FROM modo order by descri"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof

  sHTML = sHTML & "<option value=" & rs2(0)
  if cstr(rs2(0)) = cstr(ccomp3) then   sHTML = sHTML & " selected "
  sHTML = sHTML & " > " & rs2(1) & "</option>"

  rs2.movenext
  loop

  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "</table><br><br>"

  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","80%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td class=tt width=20%  >"
  sHTML = sHTML & "Anterior"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td class=uti width=80%  >"
  Strsql = "SELECT tipoaccionalt,descri,modo From TIPOACCALT  where tipoaccionalt = " & ccomp
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sAccion = rs3(1)
      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccion,i-1)
        sAcciont = mid(sAccion,i+1)
      else
        sAcciont=""
      end if
      sInput = "<font color=#000080 ><i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & replace(sAcciont," ","&nbsp;") & "</font>"

      Strsql = "SELECT archivo From modo  where tipomodo = '" & rs3(2) & "' "
      set rs4 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs4.eof then
        sImage = ofv.GenImage("","","","","Iconos/" & rs4(0),"0","")
      else
        sImage = ofv.GenImage("","","","","Iconos/ICorAImagen.jpg","0","")
      end if

      sInput = sImage & "&nbsp;&nbsp;&nbsp;&nbsp;" & sInput
  else
     sImage = ofv.GenImage("","","","","Iconos/ICorAImagen.jpg","0","")
     sInput = sImage &  "<font color=#000080 ><B>" & "&nbsp;&nbsp;&nbsp;&nbsp;SIN&nbsp;ASOCIAR" & "</b></font>"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & sInput
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td class=tt  >"
  sHTML = sHTML & "Nueva"
  sHTML = sHTML & "</td>"


  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS

  sHTML = sHTML & "<td class=uti   >"
  Strsql = "SELECT tipoaccionalt,descri,modo From TIPOACCALT  where tipoaccionalt = " & ccomp2
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sAccion = rs3(1)
      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccion,i-1)
        sAcciont = mid(sAccion,i+1)
      else
      sAcciont=""
      end if
      sInput = "<font color=#550000 ><i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & replace(sAcciont," ","&nbsp;") & "</font>"

      Strsql = "SELECT archivo From modo  where tipomodo = '" & rs3(2) & "' "
      set rs4 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs4.eof then
        sImage = ofv.GenImage("","","","","Iconos/" & rs4(0),"0","")
      else
        sImage = ofv.GenImage("","","","","Iconos/ICorAImagen.jpg","0","")
      end if

      sInput = sImage & "&nbsp;&nbsp;&nbsp;&nbsp;" & sInput

  else
     sImage = ofv.GenImage("","","","","Iconos/ICorAImagen.jpg","0","")
     sInput = sImage & "<font color=#550000 ><B>" & "&nbsp;&nbsp;&nbsp;&nbsp;SIN&nbsp;ASOCIAR" & "</b></font>"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & sInput
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>" & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 align=center >"
  sHTML = sHTML & ofv.GenTabla("panel4","aaa cellpadding=3  ","","30%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr height=80 >"
  sHTML = sHTML & "<td align=center valign=bottom >"
  if bModificar or true then
     sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
  sHTML = sHTML & ofv.GenerarInput(rs(4).name,ccomp2,"hidden","","10","","")
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center valign=bottom   >"
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=Cancel ID=Cancel action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Cancelar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"

  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

      'Estructura del Indice

     Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	 Fila = ofv.GenRow("","","","","","","",Celda,"si")
	 sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>