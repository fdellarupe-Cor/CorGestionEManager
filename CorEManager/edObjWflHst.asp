<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sestado = 1
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar


  vuelta = ""


 ' response.write "modif " & bModificar
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../") & vuelta
  vuelta2 = ofv.vueltaasp2("")
  sestado = 9

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * FROM OBJprty  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sestado = rs("estado")
  call ofv.cerrarconsulta(rs)

  if CSTR(sestado) > "2" then
     if session("Nomodificar") = "S" then bModificar = False
     response.redirect "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo
' Session("Form2") <> "DESARROLLO"
  elseif CSTR(sestado) = "2"  and Session("Form2") <> "DESARROLLO"  then
     response.redirect "edobjwflconf.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
 'and not (session("segct") = "S")
  else
  if not (session("segct") = "S") then
     if session("Nomodificar") = "S" then bModificar = False
  end if
  sHTML = ofv.FormBody
  Strsql = "SELECT * FROM OBJWFL  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     strsql = "Insert into OBJWFL (codobjeto, tipobjeto) "
     strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "')"

     call ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if

  call ofv.cerrarconsulta(rs)

  xActiva = 0
  xObliga = 0

  Strsql = "SELECT * FROM Defwfl"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

     xActiva = rs(1)
     xObliga = rs(2)

  end if

  call ofv.cerrarconsulta(rs)

  if not xActiva then xObliga = 0

  if xObliga then

     strsql = "update OBJWFL set esqauto = 1, ordenest = 1, chat = 1  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

     call ofv.crearconsultaEx(StrSql,cn,1,volver)

 end if



     Strsql = "SELECT * FROM OBJWFL  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    Celda = ofv.gencelda("","tt colspan=4","center","2%","","","","Circuito de Aprobaci&oacute;n","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")


  if not rs.eof then

    sForm = "wfl"
    Celda = ofv.gencelda("","cc colspan=2","","70%","","","","Lista de Usuarios:","si")

    if bModificar and Cint(sestado) = 1 and not xobliga then
       sInput = ofv.combosino(sForm & "1",rs(3),rs(3).name)
       sInput = ofv.gencelda("","","","2%","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "1","",sAccion,"","post","",sInput,"si")
    else
       sval = "NO"
       if rs(3) then sval = "SI"
       sinput = ofv.gencelda("","cc","","2%","","","",sval,"si")
    end if

    Celda = Celda & sinput

'    response.write "rs(3) " & rs(3) & "| "
 '   if not rs(3)  or Cint(sestado) = "2" then


    if rs(3) and Cint(sestado) = 1  then
       sInput = ofv.GenerarInput("boton1","Lista","submit",sForm & "1a","5","bt","")
       sInput = ofv.gencelda("","cc","center","2%","","","",sInput,"si")
       sAccion = "edobjwfldet.asp?volver=" & vuelta2
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "1a","",sAccion,"","post","",sInput,"si")
    else

       sInput = ofv.gencelda("","cc","","2%","","","","---","si")


    end if

    Celda = Celda & sinput
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    if rs(3) then
    Celda = ofv.gencelda("","cc colspan=2","","60%","","","","Orden Estricto:","si")

    if bModificar  and Cint(sestado) = 1  and not xobliga then
       sInput = ofv.combosino(sForm & "2",rs(4),rs(4).name)
       sInput = ofv.gencelda("","","","2%","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "2","",sAccion,"","post","",sInput,"si")
    else
       sval = "NO"
       if rs(4) then sval = "SI"
       sinput = ofv.gencelda("","cc","","2%","","","",sval,"si")
    end if

    Celda = Celda & sinput


    sInput = ofv.gencelda("","cc","","2%","","","","&nbsp;","si")


    Celda = Celda & sinput
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
   end if

    if rs(3) then
    Celda = ofv.gencelda("","cc colspan=2","","60%","","","","Enviar Mensaje:","si")

    if bModificar  and cint(sestado) = 1  and not xobliga then
       sInput = ofv.combosino(sForm & "4",rs(5),rs(5).name)
       sInput = ofv.gencelda("","","","2%","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "4","",sAccion,"","post","",sInput,"si")
    else
       sval = "NO"
       if rs(5) then sval = "SI"
       sinput = ofv.gencelda("","cc","","2%","","","",sval,"si")
    end if

    Celda = Celda & sinput

    sInput = ofv.gencelda("","cc","","2%","","","","&nbsp;","si")

    Celda = Celda & sinput
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


    'Texto del Mensaje
  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:#f7eedd;} "
    if rs(5) then

       Celda = ofv.gencelda("","tt colspan=4","center","100%","","","","Mensaje:","si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

       if bModificar  and Cint(sestado) = 1 then

          If BrowserName() = "IE" OR BrowserName() = "CHROME" then
             nCol = 90
          else
             nCol = 68
          end if
          t3b = "<font class=at><textarea CLASS=DT "
          t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=4 "
          if bModificar then
             t3b = t3b & " onblur=" & sForm & "3.submit() "
             t3b = t3b & " onkeypress=" & chr(34)
             t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
             t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
          else
             t3b = t3b & " readonly>"
          end if
          t3b = t3b & rs(7) & "</textarea></font>"
          sInput = ofv.gencelda("","cc colspan=4","","100%","","","",t3b,"si")
          sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
          sInput =  ofv.GenForm(sForm & "3","",sAccion,"","post","",sInput,"si")
       else
          sinput = ofv.gencelda("","cc colspan=4","center","100%","","","",rs(7),"si")
       end if

       Celda = sinput
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    end if
   end if


 end if
  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & "<br><br><br><br><center>"
  sHTML = sHTML & ofv.GenTabla("","","","80% border=0 ","","0","0","1","0","#bbbbbb","",fila,"si")


     Strsql = "SELECT OBJPRTY.*, Estados.estado FROM OBJPRTY, Estados "
     Strsql = Strsql & "WHERE OBJPRTY.estado = Estados.objid AND codobjeto=" & nCodigo
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        if session("segct") = "S" then
       Celda = ofv.gencelda("","tt colspan=4","center","60%","","","","Cambio de Estado","si")
       Fila =  ofv.GenRow("","","","","","","",Celda,"si")
       Celda = ofv.gencelda("","tt ","","10%","","","","Estado&nbsp;Actual:","si")
       sinput = ofv.gencelda("","cc ","","40%","","","",rs(20),"si")
       Celda = celda & sinput
       Celda = celda & ofv.gencelda("","tt ","","10%","","","","Nuevo&nbsp;Estado:","si")
'xxxxxxxxxxxxx
    'Estado
     dim Aest(4)
     Aest(0) = 0
     Aest(1) = "Desarrollo"
     Aest(2) = "Para Aprobar"
     Aest(3) = "Aprobada"
     Aest(4) = "Eliminada"

     Strsql = "SELECT * from estados  "
     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do until rsx.eof
        Aest(rsx(0)) = Ucase(rsx(1))
        rsx.movenext
     loop
     call ofv.cerrarconsulta(rsx)

    Select case cstr(rs(9))
      case "1"
        if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S"  then
          sInput = "<font class=at><select name=" & rs(9).name
          sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
          sInput = sInput & "<option selected value=1>" & Aest(1) & "</option>"
          sInput = sInput & "<option value=2>" & Aest(2) & "</option></select></font>"
        else
          sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
        end if
      case "2"
        if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S" then
          sInput = "<font class=at><select name=" & rs(9).name
          sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
          sInput = sInput & "<option selected value=2>" & Aest(2) & "</option>"
          sInput = sInput & "<option value=1>" & Aest(1) & "</option>"
          sInput = sInput & "</select></font>"
        else
          sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
        end if
    end select
    Celda = Celda & ofv.GenCelda("","cc","","","","","",sInput,"si")
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if (bModificar   or session("Nomodificarce") <> "S") and Session("Segce") = "S" then
      sAccion = "asp/ActualizarEN.asp?opc=WFL&ID=" & rs(0) & "&Tabla=Objprty&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    Formulario2 = ofv.GenForm("edPropiedades2","",sAccion,"","post","",Fila,"si")

   end if
  end if

'xxxxxxxxxx
    Celda = ofv.gencelda("","aaa colspan=4","","10%","","","","&nbsp;","si")
    Fila =  ofv.GenRow("","","","","","","",Celda,"si")



      sHTML = sHTML & "<br>"
      sHTML = sHTML & ofv.GenTabla("","","","80% border=0 ","","0","0","1","0","#bbbbbb","",Formulario2 & fila,"si")



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  sestado = 1
  sHTML = sHTML & "</center></body></html>"
 end if
end if

  if Cint(sestado) = 1 then response.write sHTML

%>