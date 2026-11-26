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
  if len(nblk) = 0 then nblk = "S"

  nAccion = request.querystring("Accion")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vueltaC = request.querystring("volver")

  if session("volverproc") = "" then session("volverproc") = vueltaC
  if vueltaC = "" then vueltaC = session("volverproc")

 ' volver2 = ofv.vueltaasp2("../") & "_Accion=" & nAccion
  volver2 = ofv.vueltaaspEx("../","0","0") & "_Accion=" & nAccion

  esttxt = " STYLE={font-family:tahoma;color:#4466bb;font-size:10pt;font-weight:bold;background-color:#ffffff;} "

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * From objaccion where objid = " & nAccion
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     wproced = rs(3)
     waccion = rs(18)
     wlink = rs(12)
     if isnull(wlink) then wlink = 0
 '    wtxtl = rs(20)
  end if

  call ofv.cerrarconsulta(rs)

 '' sHTML = ofv.menuheader("","")


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  sHTML = sHTML & "<center><br>"



' inserta link local

  if nblk = "IL" then

     strsql = "delete from OBJLNK "
     strsql = strsql & "where  Ltipo='R' and iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
     call ofv.crearconsultaEx(StrSql,cn,1,volver)
    'response.write "anular link:nblk IL " & strsql & "<br>"

     wlink = request.form("comp1")
     widen4 = request.form("iden4")&""
     'response.write  widen4
     if wlink <> "" then

        strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo, iden3L ,iden4"
        strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',"
        strsql = strsql & "'" & wproced & "','" & waccion & "',"
        strsql = strsql & nCodigo & ",'" & sTipo & "','','','R','" & wlink & "'"
        strsql = strsql & ",'" & widen4 &"')"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "SELECT * From OBJLNK where  Ltipo='R' and iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
       ' response.write Strsql
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then
           wlnkid = rs(0)
           wLtipor=rs("Ltipo")
        else
           wlnkid = 0
           wlink = 0

        end if
        call ofv.cerrarconsulta(rs)

        if cdbl(wlnkid) > 0 then
           strsql = "update OBJaccion set tiempo = " & wlnkid
           strsql = strsql & " where  objid = " & nAccion
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if
     end if

     nblk = "S"
  end if








' agrega link local 1

  if nblk = "AL" then
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

   if bModificar then
      sAccion = "EdProcedrel.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=AL2&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","50%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=ttt align=center  >"
  sHTML = sHTML & "Relacionar&nbsp;circuito "
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  

'-------------
 
  'sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)


  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 align=center  >"

  sHTML = sHTML & "<font class=at><select size=10  name=comp1  "
  sHTML = sHTML & " >"

    Strsql = "SELECT codproced, nombre "
    Strsql = Strsql & " from CSCEventoSectorPerfil Where Obcodobjeto = " & nCodigo
    Strsql = Strsql & " order by secuen, secuencia, codproced "
      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof
     sHTML = sHTML & "<option value=" & rs2(0)
     sHTML = sHTML & " > " & rs2(1) & "</option>"

     rs2.movenext
  loop
  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"

'--------



  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  'sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</table><br><br>"
  sHTML = sHTML & "</FORM>" & chr(13)


     nblk = "NO"

  end if


' agrega link local 2

  if nblk = "AL2" then

   sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

   if bModificar then
      sAccion = "EdProcedrel.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=IL&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 ","","50%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 class=TTT align=center  >"
  sHTML = sHTML & "Relacionar&nbsp;circuito&nbsp;-&nbsp;paso:"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
 
'-------------
 

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=2 align=center  >"

  sHTML = sHTML & "<font class=at><select size=10  name=comp1  "
  sHTML = sHTML & " >"

     wlink = request.form("comp1")

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

    strsql = " select * from objaccion "
    strsql = strsql & "WHERE codproced = '" &  wlink & "' "
    if bVersion = "NO" then
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    end if
    strsql = strsql & " Order By secuen"
    'response.write strsql
      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do while not rs2.eof

      sAccion = "S/D"

      if (rs2(4) & "") <> "" then
         Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs2(4)
      '''response.write Strsql & " Strsql " & chr(13)
         set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rsAcciones.eof then
             sAccion = rsAcciones(0)
         end if
         call ofv.cerrarconsulta(rsAcciones)
      end if


     sHTML = sHTML & "<option value=" & rs2(18)
     sHTML = sHTML & " > " & rs2(10) & "-" & sAccion & "</option>"

     rs2.movenext
  loop
  call ofv.cerrarconsulta(rs2)
  sHTML = sHTML & "</select></font>"

'--------
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=RIGHT>"
  sHTML = sHTML & "GRAFICA CONECTOR</TD><TD>"

  scomK = "S" & chr(9) & "N"
  scomD = "SI" & chr(9) & "NO"

  sInputCo=ofv.gencomboFX("fcomp1","iden4","S",scomK,scomD,"no")


  sHTML = sHTML & sInputCo & "</td></tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  
  sHTML = sHTML & "</table><br><br>"
  sHTML = sHTML & "</FORM>" & chr(13)


     nblk = "NO"

  end if

''response.write sAccion & " sAccion " & chr(13)

' Alta 1era vez


  if cdbl(wlink) = 0 and nblk = "S" then

  set cn2 = ofv.conectar(ofv.strconn2)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''Para obtener la descripcion de la Accion
  ''Primero obtiene el tipo de accion de la tabla objaccion
  Strsql = "SELECT * From OBJACCION "
  Strsql = Strsql & " WHERE codproced = '" & wproced & "' "
  Strsql = Strsql & " AND objid = '" & nAccion & "' "
  set rsz = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rsz.eof then
	 nTAccion  = rsz(4)
  end if
  call ofv.cerrarconsulta(rsz)

  ''Despues obtiene la descri de la accion de la tabla Tipoaccalt
  Strsql = "Select descri From Tipoaccalt "
  Strsql = Strsql & " Where tipoaccionalt = '" & nTAccion & "' "

  set rsy = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rsy.eof then
  nTdescri  = rsy(0)
  end if
  call ofv.cerrarconsulta(rsy)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=5  ","","50%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=3 class=ttt align=center  >"
  sHTML = sHTML & "Relacion a otro Circuito - "
  sHTML = sHTML & nTdescri
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

'-------------

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  COLSPAN=3 align=center   >&nbsp;</TD>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  align=center   >"

  if bModificar then
      sAccion = "EdProcedrel.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=AL&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='RELACIONAR' class=bt >"

  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"



  sHTML = sHTML & "<td  align=center   >"
  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "<input  type=submit  size=10%  value='CANCELAR' class=bt >"
  sHTML = sHTML & "</FORM>" & chr(13)
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  COLSPAN=3 align=center   >&nbsp;</TD>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "</table><br><br>"


     nblk = "NO"


  end if


''response.write sAccion & " sAccion " & chr(13)


' modicar link local


  if nblk = "ML" then

     wlink = request.form("comp1")

     Strsql = "SELECT * From objetos where codobjeto = " & wlink
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        wtipoL = rs(2)
     else
        wtipoL = "NO"
        wlink = 0
     end if
     call ofv.cerrarconsulta(rs)

     if clng(wlink) > 0 then

        strsql = "update OBJLNK set codobjetoL = " & wlink & ", tipobjetoL = '" & wtipoL & "' "
        strsql = strsql & "where iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

     end if

     nblk = "S"
  end if





' anular link


  if nblk = "A" then

 '    response.write "anular link: " & "<br>"

     sAccion =  replace(vueltaC,"../","")
     sAccion =  replace(sAccion,"_","&")

     strsql = "delete from OBJLNK "
     strsql = strsql & "where  Ltipo='R' and iden1 = '" & wproced & "' and iden2 = '" & waccion & "' "
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

    ' response.write "anular link nblk A: " & strsql & "<br>"

     strsql = "update OBJaccion set tiempo = 0 "
     strsql = strsql & " where  objid = " & nAccion
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

     response.redirect sAccion

  '   nblk = "S"
     nblk = "NO"
  end if




' Mostrar link


  if nblk = "S" then

  Strsql = "SELECT * From objaccion where objid = " & nAccion

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     wproced = rs(3)
     waccion = rs(18)
     wlink = rs(12)
     if isnull(wlink) then wlink = 0



     wtipoL = 0
     wUrl = "NO"
     wUrlDesc = "NO"
     wtipo = "NO"

     Strsql = "SELECT * From objlnk where objid = " & wlink
     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs1.eof then
        wtipoL = rs1(6)
        wUrl = rs1(8)
        wUrlDesc = rs1(9)
        wtipo = rs1(10)
        waccionx = rs1(11)
        wprocedx = rs1(3)
        winde4=rs1("iden4")
     end if
     call ofv.cerrarconsulta(rs1)

 'response.write  wtipoL &  wUrl & wUrlDesc & wtipo &  Strsql & "<br>"

  sAccion =  replace(vueltaC,"../","")
  sAccion =  replace(sAccion,"_","&")
  sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

  if bModificar then
      sAccion = "EdProcedrel.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=A&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if
  sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)
  sHTML = sHTML & "</FORM>" & chr(13)

  if bModificar then
      sAccion = "EdProcedrel.asp?"
      sAccion = sAccion & "Accion=" & nAccion  & "&nblk=AL&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
  else
     sAccion = ""
  end if

  sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & chr(34) & sAccion  & chr(34) &  "  method=post >" & chr(13)

  sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","50%","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td  colspan=3 class=ttt align=center  >"
  sHTML = sHTML & "Relacion&nbsp;a&nbsp;otro&nbsp;Circuito"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"



' muestra link local





    strsql = " select * from objaccion "
    strsql = strsql & "WHERE codaccion = '" &  waccionx & "' "
    if bVersion = "NO" then
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    end if
    strsql = strsql & " Order By secuen"
      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs2.eof then
        wadesc = "No Disponible"
        wprocedL = "NO"
     else

      sAccion = "S/D"

      if (rs2(4) & "") <> "" then
         Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs2(4)
         set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rsAcciones.eof then
             sAccion = rsAcciones(0)
         end if
         call ofv.cerrarconsulta(rsAcciones)
      end if


        wadesc = rs2(10) & "-" & sAccion
        wprocedL = rs2(3)
     end if
     call ofv.cerrarconsulta(rs2)

    if  wprocedL = "NO" then
        wpdesc = "No Disponible"
    else
      Strsql = "SELECT codproced, nombre "
      Strsql = Strsql & " from CSCEventoSectorPerfil Where codproced = '" & wprocedL & "' "
      Strsql = Strsql & " order by secuen, secuencia, codproced "
      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if rs2.eof then
         wpdesc = "No Disponible"
      else
         wpdesc = rs2(1)
      end if

      call ofv.cerrarconsulta(rs2)
    end if


  sHTML = sHTML & "<tr height=100 >"
  sHTML = sHTML & "<td  colspan=3 class=uti valign=bottom >"
  sHTML = sHTML & "Circuito: " & wpdesc
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr height=22 >"
  sHTML = sHTML & "<td  colspan=3 class=uti >"
  sHTML = sHTML & "Paso: " & wadesc
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  if wtipo="R" then
   sHTML = sHTML & "<tr height=22 >"
   sHTML = sHTML & "<td  colspan=3 class=uti >"

   if winde4 <> "N" then
     wadesc="SI"
   else
    wadesc="NO"
   end if
   sHTML = sHTML & "GRAFICA CONECTOR: " & wadesc
   sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
 end if

  sHTML = sHTML & "<tr height=80 >"
  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=SUBMIT  value='Modificar Link' class=bt >"

''''  sHTML = sHTML & "</FORM>" & chr(13)

  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"

  sHTML = sHTML & "<input  type=Button  value='Anular Link' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp3.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"

  sHTML = sHTML & "<td  align=center   >"
  sHTML = sHTML & "<input  type=Button  value='Volver' class=bt "
  sHTML = sHTML & " onclick=" & chr(34) & "fcomp2.submit();" & chr(34) &  " >"
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"



  sHTML = sHTML & "</table><br><br>"
  sHTML = sHTML & "</FORM>" & chr(13)
  end if

  call ofv.cerrarconsulta(rs)

  end if



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center>"

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no","no","no","") & "</body></html>"

end if
  if nblk <> "A" then response.write sHTML

%>