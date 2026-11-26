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

  sopcion = request.querystring("opcion")

  nAccion = request.querystring("id")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ccompdf = request.querystring("ccompdf")
  tcompdf = request.querystring("Tcompdf")
  vueltaC = request.querystring("volver")
  if session("volverproc") = "" then session("volverproc") = vueltaC
  volver2 = ofv.vueltaasp2("../")

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' busco el parrafo y la asociacion creada
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     ccomp = 0
     tcomp = "NO"

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * From OBJESTDET where objid = " & nAccion
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     ccompKL = rs(9)
     tcodaKL = nAccion
     tcodpKL = nAccion

  end if

          Strsql = "Select codobjetol,tipobjetol From objlnk Where ltipo = 'KL' and objid = 0" & ccompKL
          set rsCompoM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsCompoM.eof then

              ccomp = rsCompoM(0)
              tcomp = rsCompoM(1)


          end if
          call ofv.cerrarconsulta(rsCompoM)


'  if ccomp <> "" then
'          Strsql = "SELECT   nombre,TIPOCOMP FROM CSACOMPLE where codcomp = 0" & ccomp
'          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'          if not rsComponentes.eof THEN
'            '' ccomp = rsComponentes(0)
'             tcomp = rsComponentes(1)
'          END IF
'          call ofv.cerrarconsulta(rsComponentes)
'  else
'     ccomp = 0
'     tcomp = "NO"
'  end if


  if nblk = "S" then ccompS = request.form("comp1")
  if nblk = "T" then ccompT = request.form("comp2")

  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS
  if nblk = "T" then ccomp2 = ccompT

  sHTML = ofv.menuheader("","")

  sHTML = sHTML & "<center><br>"

  if sopcion = "OK" then

     Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp = " & ccompdf
     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs3.eof then
        tcompdf = rs3(2)

     end if
     call ofv.cerrarconsulta(rs3)

     if ccomp = 0 then

        wlinkN1 = ""
        wlinkN2 = ""
        wlinkN3 = ""


        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
        strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & nCodigo & ",'" & sTipo & "',"
        strsql = strsql & "'" & tcodpKL & "','" & tcodaKL & "',"
        strsql = strsql & ccompdf & ",'" & tcompdf & "','','','KL','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
        strsql = strsql & ")"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where iden1 = '" & tcodpKL & "' and iden2 = '" & tcodaKL & "' and codobjeto = " & nCodigo
        Strsql = Strsql & " and tipobjeto = '" & sTipo & "' "
        set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsx.eof then
           ccompKL = rsx(0)
        end if
        call ofv.cerrarconsulta(rsx)

     else

       if ccompdf > 0 then

        Strsql = "update OBJLNK set codobjetol = " & ccompdf & ", tipobjetol = '" & tcompdf & "' "
        Strsql = Strsql & " where objid = " & ccompKL
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

       else

        Strsql = "delete from  OBJLNK  "
        Strsql = Strsql & " where objid = " & ccompKL
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        ccompKL = 0


       end if

     end if

     sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objestdet&volver=" & vueltaC
     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


     sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
     sHTML = sHTML & ofv.GenerarInput(rs(9).name,ccompKL,"hidden","","10","","")
'''  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"
     sHTML = sHTML & "</FORM>" & chr(13)



     sHTML = sHTML & "<SCRIPT language=javascript >" & chr(13)
     sHTML = sHTML & " var dxformwk   = document.getElementById('acep'); " & chr(13)
     sHTML = sHTML & " dxformwk.submit(); " & chr(13)
     sHTML = sHTML & "</SCRIPT>" & chr(13)

  else

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td align=center valign=top   >"

  if nblk = "S" or (nblk = "D" and ord = "T") then
     T1styl = " style='display:block;' "
     T2styl = " style='display:none;' "
  else
     T1styl = " style='display:none;' "
     T2styl = " style='display:block;' "
  end if

''''''''' tabla 1

  sHTML = sHTML & ofv.GenTabla("panelT1","aaa cellpadding=3 " & T1styl & " ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td class=ttt align=center  >"
  sHTML = sHTML & "Complementos&nbsp;asociados&nbsp;al&nbsp;Documento"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td class=ttt align=center width=5% style='cursor:pointer;' title='Ver todos los complementos disponibles.' "
  sHTML = sHTML & " onclick=""document.getElementById('panelT1').style.display='none';document.getElementById('panelT2').style.display='';"" >"
  sHTML = sHTML & "<font face='corflaticon' size=3 color=#f7aa22 style='border-width:1;border-style:solid;border-color:#f7aa22;' >4</font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 align=center  >"
'-------------



  if bModificar then
      sAccion = "EdEstrCMP.asp?"
      sAccion = sAccion & "id=" & nAccion  & "&nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp1.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=0 "

  if (cstr(ccomp) = "0"  and nblk = "D") or cstr(ccomp2) = "0" then
     sHTML = sHTML & " selected "
  end if
  sHTML = sHTML & " >---------- Sin Asociar ----------</option>"

 '' Strsql = "SELECT distinct detimagen From Objestdet where codobjeto = " & nCodigo
 '' Strsql = Strsql & " and tipobjeto = '" & sTipo & "' and dettipo = 'J' and not detimagen is null "

  Strsql = "Select distinct codobjetol,tipobjetol From objlnk Where ltipo = 'KL' and codobjeto = " & nCodigo
  Strsql = Strsql & " and tipobjeto = '" & sTipo & "'  order by tipobjetol "


 '' response.write strsql & "<br>"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof

  Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp = 0" & rs2(0)
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     ncomp = rs3(1)
  end if
  call ofv.cerrarconsulta(rs3)

  sHTML = sHTML & "<option value=" & rs2(0)
  if  cstr(rs2(0)) = cstr(ccomp2) then   sHTML = sHTML & " selected "
  sHTML = sHTML & " > " & ncomp & "</option>"

  rs2.movenext
  loop

  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"
  sHTML = sHTML & "</FORM>" & chr(13)
'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"


''''''''' tabla 2

  sHTML = sHTML & ofv.GenTabla("panelT2","aaa cellpadding=3 " & T2styl & " ","","80%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td class=ttt align=center width=5% style='cursor:pointer;' title='Ver solo los complementos asociados.' "
  sHTML = sHTML & " onclick=""document.getElementById('panelT2').style.display='none';document.getElementById('panelT1').style.display='';"" >"
  sHTML = sHTML & "<font face='corflaticon' size=3 color=#f7aa22 style='border-width:1;border-style:solid;border-color:#f7aa22;' >3</font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td class=ttt align=center  >"
  sHTML = sHTML & "Complementos&nbsp;disponibles"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 align=center  >"
'-------------
  if bModificar then
      sAccion = "EdEstrCMP.asp?"
      sAccion = sAccion & "id=" & nAccion  & "&nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select size=10  name=comp2 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp2.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=0 "
  if (cstr(ccomp) = "0"  and nblk = "D") or cstr(ccomp2) = "0" then
     sHTML = sHTML & " selected "
  end if
  sHTML = sHTML & " >---------- Sin Asociar ----------</option>"

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
  sHTML = sHTML & "<td  colspan=2 align=center class=ttt   >"
  sHTML = sHTML & "FILTRO: "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td colspan=2 align=center  >"
'-------------
  if bModificar then
      sAccion = "EdEstrCMP.asp?"
      sAccion = sAccion & "id=" & nAccion  & "&nblk=D&ord=F&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & chr(34) &  sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<font class=at><select  name=comp3 style={border-style:outset;border-width:4;} "
  sHTML = sHTML & " onchange=" & chr(34) & "fcomp3.submit();" & chr(34)

  sHTML = sHTML & " >"

  sHTML = sHTML & "<option value=TODOS "
  if (cstr(ccomp) = "0"  and nblk = "D") or cstr(ccomp2) = "0" then
     sHTML = sHTML & " selected "
  end if
  sHTML = sHTML & " >---------- TODOS ----------</option>"

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
  sHTML = sHTML & "</table>"



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table><br><br>"




  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","80%","",0,0,0,0,"","","","no")
  sHTML = sHTML & "<tr height=25 >"
  sHTML = sHTML & "<td class=tt width=20%  >"
  sHTML = sHTML & "Anterior"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "<td class=uti width=80%  >"
  Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp = 0" & ccomp
 '' response.write strsql & "<br>"
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sInput = "<font face='corflaticon' size=2 >4</font>&nbsp;" & replace(rs3(1) & "","Continuador:","")
  else
     sInput = "SIN&nbsp;ASOCIAR"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & sInput
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr height=25 >"
  sHTML = sHTML & "<td class=tt  >"
  sHTML = sHTML & "Nuevo"
  sHTML = sHTML & "</td>"


  if nblk = "D" then ccomp2 = ccomp
  if nblk = "S" then ccomp2 = ccompS
  if nblk = "T" then ccomp2 = ccompT

  sHTML = sHTML & "<td class=uti   >"
  Strsql = "SELECT  codcomp, nombre, tipocomp FROM CSACOMPLE where codcomp = 0" & ccomp2
  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs3.eof then
     sInput = "<font face='corflaticon' size=2 >4</font>&nbsp;" & replace(rs3(1) & "","Continuador:","")
  else
     sInput = "SIN&nbsp;ASOCIAR"
  end if
  call ofv.cerrarconsulta(rs3)
  sHTML = sHTML & "<font color=#993333 >" & sInput & "</font>"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>" & chr(13)

  sHTML = sHTML & "<tr >"
  sHTML = sHTML & "<td colspan=2 align=center >"

    sHTML = sHTML & ofv.GenTabla("panel4","aaa cellpadding=5  ","","20%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr height=80 >"
  sHTML = sHTML & "<td align=center valign=top >"
  if bModificar then
 ''    sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objestdet&volver=" & vueltaC
 ''    sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

      sAccion = "EdEstrCMP.asp?"
      sAccion = sAccion & "opcion=OK&ID=" & nAccion  & "&ccompdf=" & ccomp2 & "&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC


  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & chr(13)
''  sHTML = sHTML & ofv.GenerarInput(rs(9).name,ccomp2,"hidden","","10","","")
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center valign=top >"
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & "<FORM  name=Cancel ID=Cancel action=" & sAccion & "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Cancelar' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)

    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"
    sHTML = sHTML & "</table>"


  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sHTML = sHTML & "</table>"

 end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>