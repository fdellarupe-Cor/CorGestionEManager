<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar


  nblk = request.querystring("nblk")
  if len(nblk) = 0 then nblk = "D"

  ord = request.querystring("ord")
  if len(ord) = 0 then ord = "T"

  opc = request.querystring("opc")
  cid = request.querystring("cid")



  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vueltaC = request.querystring("volver")
'  if session("volverproc") = "" then session("volverproc") = vueltaC
  volver2 = ofv.vueltaasp2("../")

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

        ccomp = 0
        tcomp =  sTipo

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "Select * From CompCompuesto "
  if opc = "NO" then
     strsql = strsql & " Where CompMaster = " & nCodigo & " and id = " & cid
  end if 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     if opc = "NO" then
        ccomp = rs(3)
        tcomp = rs(4)
     end if  
  end if


  if nblk = "S" then ccompS = request.form("comp1")
  if nblk = "T" then ccompT = request.form("comp2")

  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS
  if nblk = "T" then ccomp2 = ccompT

  sHTML = ofv.FormBody & "<center><br>"

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"

  sHTML = sHTML & "<td width=50%  class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Complementos&nbsp;disponibles"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"

  sHTML = sHTML & "<td width=50% align=center style={border-style:outset;border-width:2;} >"
'-------------
  if bModificar then
      sAccion = "EdCompuestoCMP.asp?" 
      sAccion = sAccion & "nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp2 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp2.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=0 "
  if (clng(ccomp) = 0  and nblk = "D") or clng(ccomp2) = 0 then 
     sHTML = sHTML & " selected " 
  end if
  sHTML = sHTML & " >*** Sin Asociar ***</option>" 

  if ord = "T" then  
     Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE order by nombre"
  else
     ccomp3 = request.form("comp3")
     if len(ccomp3) = 0 then
        Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE order by nombre"
     else
        Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where tipocomp = '" & ccomp3 & "' order by nombre"
     end if
  end if
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof

  sHTML = sHTML & "<option value=" & rs2(0)
  if cstr(rs2(0)) = cstr(ccomp2) then   sHTML = sHTML & " selected "
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

  sHTML = sHTML & "<td width=50% align=center class=tt style={border-style:outset;border-width:2;}  >"
  sHTML = sHTML & "FILTRO: "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=50% align=center style={border-style:outset;border-width:2;} >"
'-------------
  if bModificar then
      sAccion = "EdCompuestoCMP.asp?" 
      sAccion = sAccion & "nblk=D&ord=F&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid
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
  sHTML = sHTML & " >*** TODOS ***</option>" 

  Strsql = "SELECT  tipo, descri FROM tipocomp order by descri"
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




  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:outset;border-width:2;} ","","80%","",0,0,0,0,"","","","no")

if opc = "NO" then

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Complemento&nbsp;anterior"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2  class=cc   >"
  Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp = " & ccomp
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     ncomp = rs3(1)
  else
     ncomp = "SIN&nbsp;ASOCIAR"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & "<font color=#000080 ><B>" & ncomp & "</b></font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2  >"
  sHTML = sHTML & "&nbsp;"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

end if

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 class=tt align=center style={border-style:outset;border-width:2;} >"
  sHTML = sHTML & "Complemento&nbsp;nuevo"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS
  if nblk = "T" then ccomp2 = ccompT

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2  class=cc   >"
  Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp = " & ccomp2
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     ncomp = rs3(1)
     ccompx = rs3(2)
  else
     ncomp = "SIN&nbsp;ASOCIAR"
     ccompx = sTipo
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & "<font color=#993333 ><B>" & ncomp & "</b></font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>" & chr(13) & chr(13)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td align=center  >"
  if bModificar then
'     sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
'     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    if opc = "NO" then
       sAccion = "asp/ActualizarComp2.asp?Tabla=Compcompuesto&volver=../edCompuesto2.asp&Codcomp="
       sAccion = sAccion & nCodigo & "&Tipocomp=" & sTipo & "&ID=" & cid
    else
       sAccion = "asp/AgregarComp2.asp?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
       sAccion = sAccion & "&volver=../edCompuesto2.asp"
    end if


  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & ofv.GenerarInput(rs(3).name,ccomp2,"hidden","","10","","")
  sHTML = sHTML & ofv.GenerarInput(rs(4).name,ccompx,"hidden","","10","","")
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sAccion =  "edCompuesto2.asp?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo & "&OpcAgregar=no"

  sHTML = sHTML & "<FORM  name=Cancel ID=Cancel action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Cancelar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>