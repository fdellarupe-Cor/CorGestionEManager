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
     ccomp = rs(9)
     if isnull(ccomp) then ccomp = " "
     scomp = rs(8)
  end if


  if nblk = "S" then ccompS = request.form("comp1")
  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS

'''  sHTML = ofv.menuheader("","")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  sHTML = sHTML & "<center><br>"


  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 ","","80%","",0,0,0,0,"","","","no")


  IF WKAPLIC = "PL" THEN
     Strsql = "SELECT ECO,descri from ECO  WHERE ECO = " & scomp
  ELSE
     Strsql = "SELECT codsector,xdescri from csContinuador  WHERE codobjeto=" & nCodigo & " and codsector = " & scomp
  END IF



  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sInput = rs3(1)
  else
     sInput = "Sin&nbsp;asociar"
  end if

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  class=ttt align=center >"
  sHTML = sHTML & "Perfiles&nbsp;asociados&nbsp;a&nbsp;" & sInput
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center  >"
'-------------
  if bModificar then
      sAccion = "EdProcedPER.asp?" 
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=nsnc "
  if ((ccomp = " "  or ccomp = "0"   or ccomp = "nsnc") and nblk = "D") or (ccomp2 = " " or ccomp2 = "nsnc") then 
     sHTML = sHTML & " selected " 
  end if
  sHTML = sHTML & " >---------- Sin Asociar ----------</option>" 

  IF WKAPLIC = "PL" THEN
     Strsql = "SELECT A.CODPERFIL,B.descri from ECOPERFIL A "
     Strsql = Strsql & " INNER JOIN TIPOPERFIL B ON (A.CODPERFIL = B.TIPOPERFIL) "
     Strsql = Strsql & " WHERE A.codsector = " & scomp & " order by b.descri "
  ELSE
     Strsql = "SELECT codperfil,xdescri FROM csPerfiles WHERE codsector=" & scomp & " order by xdescri "
  END IF

 ''' response.write strsql & "<br>"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof
  sHTML = sHTML & "<option value=" & rs2(0)
  if  cstr(rs2(0)) = cstr(ccomp2) then   sHTML = sHTML & " selected "
  sHTML = sHTML & " > " & replace(rs2(1) & "","Perfil:","") & "</option>"

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

  sHTML = sHTML & "<tr height=25 >"
  sHTML = sHTML & "<td class=tt width=20% >"
  sHTML = sHTML & "Perfil/funcion&nbsp;anterior"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td class=uti width=80% >"

  IF WKAPLIC = "PL" THEN
     Strsql = "SELECT A.CODPERFIL,B.descri from ECOPERFIL A "
     Strsql = Strsql & " INNER JOIN TIPOPERFIL B ON (A.CODPERFIL = B.TIPOPERFIL) "
     Strsql = Strsql & " WHERE A.codsector = " & scomp & " and a.codperfil = '" & ccomp & "' "
  ELSE
     Strsql = "SELECT codperfil,xdescri FROM csPerfiles WHERE codsector=" & scomp & " and codperfil = '" & ccomp & "' "
  END IF



  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sInput = "<font face='corflaticon' size=2 >4</font>&nbsp;" & replace(rs3(1) & "","Perfil:","")
  else
     sInput = "SIN&nbsp;ASOCIAR"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & sInput
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr height=25 >"
  sHTML = sHTML & "<td class=tt  >"
  sHTML = sHTML & "Perfil/funcion&nbsp;nuevo"
  sHTML = sHTML & "</td>"

  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS
  sHTML = sHTML & "<td   class=uti   >"

  IF WKAPLIC = "PL" THEN
     Strsql = "SELECT A.CODPERFIL,B.descri from ECOPERFIL A "
     Strsql = Strsql & " INNER JOIN TIPOPERFIL B ON (A.CODPERFIL = B.TIPOPERFIL) "
     Strsql = Strsql & " WHERE A.codsector = " & scomp & " and a.codperfil = '" & ccomp2 & "' "
  ELSE
     Strsql = "SELECT codperfil,xdescri FROM csPerfiles WHERE codsector=" & scomp & " and codperfil = '" & ccomp2 & "' "
  END IF


  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sInput = "<font face='corflaticon' size=2 >4</font>&nbsp;" & replace(rs3(1) & "","Perfil:","")
  else
     sInput = "SIN&nbsp;ASOCIAR"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & "<font color=#993333 >" & sInput & "</font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>" & chr(13) & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 align=center  >"

  sHTML = sHTML & ofv.GenTabla("panel3","aaa cellpadding=5  ","","20%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr height=80 >"
  sHTML = sHTML & "<td align=center  >"
  if bModificar then
     sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
  sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")
  if ccomp2 = "nsnc" then
     sHTML = sHTML & ofv.GenerarInput(rs(9).name," ","hidden","","10","","")
  else
     sHTML = sHTML & ofv.GenerarInput(rs(9).name,ccomp2,"hidden","","10","","")
  end if
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
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

  sHTML = sHTML & "</center>"

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & "</body></html>"


end if

response.write sHTML

%>