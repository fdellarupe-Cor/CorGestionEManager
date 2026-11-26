<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sestado = 2
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")

  set cn = ofv.conectar(ofv.strconn2)
  sestado = 9
  Strsql = "SELECT * FROM OBJprty  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sestado = rs("estado")
  call ofv.cerrarconsulta(rs)

  if cint(sestado) > 2 then
     response.redirect "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo

  elseif cint(sestado) = 1  then

      sHTML = ofv.FormBody

      Celda = ofv.gencelda("","tt colspan=4","center","2%","","","","Circuito de Aprobaci&oacute;n","si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")
      Celda = ofv.gencelda("","tt colspan=5","center","2%","","","","Confirmaci&oacute;n&nbsp;-&nbsp;Rechazo","si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

      sinput = "El presente Esquema Normativo retorna al estado Desarrollo<br>  "
      sinput = sinput & "debido al rechazo de alguno de los autorizantes"
      Celda = ofv.gencelda("","cc colspan=4","center","100%","","","",sinput,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","90%","","0","0","1","0","#bbbbbb","",fila,"si")


  else

  sHTML = ofv.FormBody

  sHTML = sHtml & "<div width=80% >"

  Celda = ofv.gencelda("","tt colspan=3","","2%","","","","Circuito de Aprobaci&oacute;n","si")
 sInput = ofv.GenerarInput("bk1","Ver&nbsp;Observaciones","submit","","10","bt","")
 sInput = ofv.gencelda("","cc","","20%","","","",sInput,"si")
 sAccion = "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo
 sInput = ofv.GenForm(sForm & "v","",sAccion,"_blank","post","",sInput,"si")

 Celda = celda & sInput

  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","tt colspan=5","center","2%","","","","Confirmaci&oacute;n&nbsp;-&nbsp;Rechazo","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


  sesqauto = 0
  sordenest = 0
  schat = 0
  snova = 0

  Strsql = "SELECT * FROM OBJwfl  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sesqauto = rs(3)
     sordenest = rs(4)
     schat = rs(5)
  end if
  call ofv.cerrarconsulta(rs)

  if sesqauto then

     sant = ""
     susr = ""
     Strsql = "SELECT * FROM OBJwfldet  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "' order by orden,objid"

     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do until rs.eof
        if rs(4) = session("usuario") then
           susr = rs(4)
           sord = rs(3)
           exit do
        else
           sant = rs(8)
        end if
        rs.movenext
     loop
     call ofv.cerrarconsulta(rs)

     if susr <> session("usuario") then
        snova = 1
     elseif sordenest and sant = "SC" then
        snova = 2
     end if



  else

      susr = ""
     Strsql = "SELECT * FROM OBJwfldet  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "' order by orden,objid"

     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do until rs.eof
        if rs(4) = session("usuario") then
           susr = rs(4)
           exit do
        else
           Strsql = "delete OBJwfldet  "
           Strsql = Strsql & " Where objid=" & rs(0)
           call ofv.crearconsultaEx(StrSql,cn,1,volver)
        end if
        rs.movenext
     loop
     call ofv.cerrarconsulta(rs)

     if susr <> session("usuario") then

        strsql = "Insert into OBJWFLdet (codobjeto, tipobjeto,usr,conf) "
        strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','" & session("usuario")
        strsql = strsql & "','SC')"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

     end if
           sord = 0

  end if

  Celda = ofv.gencelda("","tt","","10%","","","","Usuario:","si")
  Celda = Celda & ofv.gencelda("","cc","","70%","","","",session("usuario") & "&nbsp;-&nbsp;" & session("segusuario"),"si")
  Celda = Celda & ofv.gencelda("","tt","","10%","","","","Orden:","si")
  Celda = Celda & ofv.gencelda("","cc","","10%","","","",sord,"si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  select case snova
         case 0

            Strsql = "SELECT * FROM OBJwfldet  "
            Strsql = Strsql & " Where codobjeto=" & nCodigo
            Strsql = Strsql & " And tipobjeto = '" & stipo & "' and usr = '"
            Strsql = Strsql & session("usuario") & "' "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

            if not rs.eof then
               sForm = "wfl" & rs(0)

               Celda = ofv.gencelda("","tt colspan=4","","100%","","","","Observaciones:","si")
               Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:#f7eedd;} "

               If BrowserName() = "IE" OR BrowserName() = "CHROME" then
                  nCol = 140
               else
                  nCol = 120
               end if

               wtxt1 = rs(5)

               if len(wtxt1) = 0 then
                  stexto = " "
               else
                  stexto = wtxt1
               end if

               t3b = "<font class=at>1&nbsp;-&nbsp;<textarea size=100% type=textarea "
               t3b = t3b & "name=" & rs(5).name & " cols=" & nCol & " rows=6 " & esttxt
               t3b = t3b & " onblur=" & sForm & "3.submit()"
               t3b = t3b & " onkeypress=" & chr(34)
               t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
               t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
               t3b = t3b & stexto & "</textarea></font>"

               sInput = ofv.gencelda("","cc colspan=4 valign=top","center","100%","","","",t3b,"si")
               sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver=" & vuelta
               sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
               sInput =  ofv.GenForm(sForm & "3","",sAccion,"","post","",sInput,"si")

               Celda = sinput
               Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

               wtxt1 = rs(6)

               if len(wtxt1) = 0 then
                  stexto = " "
               else
                  stexto = wtxt1
               end if

               t3b = "<font class=at>2&nbsp;-&nbsp;<textarea   "
               t3b = t3b & "name=" & rs(6).name & " cols=" & nCol & " rows=6 " & esttxt
               t3b = t3b & " onblur=" & sForm & "5.submit()"
               t3b = t3b & " onkeypress=" & chr(34)
               t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
               t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
               t3b = t3b & stexto & "</textarea></font>"

               sInput = ofv.gencelda("","cc colspan=4 valign=top","center","100%","","","",t3b,"si")
               sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver=" & vuelta
               sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
               sInput =  ofv.GenForm(sForm & "5","",sAccion,"","post","",sInput,"si")

               Celda = sinput
               Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

               wtxt1 = rs(7)

               if len(wtxt1) = 0 then
                  stexto = " "
               else
                  stexto = wtxt1
               end if

               t3b = "<font class=at>3&nbsp;-&nbsp;<textarea maxlenght=2500 "
               t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=6 " & esttxt
               t3b = t3b & " onblur=" & sForm & "6.submit()"
               t3b = t3b & " onkeypress=" & chr(34)
               t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
               t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
               t3b = t3b & stexto & "</textarea></font>"

               sInput = ofv.gencelda("","cc colspan=4 valign=top","center","100%","","","",t3b,"si")
               sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver=" & vuelta
               sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
               sInput =  ofv.GenForm(sForm & "6","",sAccion,"","post","",sInput,"si")

               Celda = sinput
               Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


               sHTML = sHTML & "<br><center>"
               sHTML = sHTML & ofv.GenTabla("","aa  border","","100%","","0","0","1","0","#bbbbbb","",fila,"si")

               Celda = ofv.gencelda("","tt","","10%","","","","Confirmaci&oacute;n:","si")

               sInput = "<font class=at><select name=" & rs(8).name
'               sInput = sInput & " onchange=document." & sForm & "4.submit(); >"
               sInput = sInput & "  >"
               ssel = ""
               if rs(8) = "SC" then
                  sInput = sInput & "<option value=SC selected >Sin Confirmar</option>"
               end if
               ssel = ""
               if rs(8) = "SI" then ssel = " selected "
               sInput = sInput & "<option value=SI " & ssel & ">Confirmado</option>"
               ssel = ""
               if rs(8) = "NO" then ssel = " selected "
               sInput = sInput & "<option value=NO " & ssel & ">Rechazado</option>"
               sInput = sInput & "</select></font>"
               sInput = ofv.gencelda("","cc","","40%","","","",sInput,"si")

               Celda = celda & sInput

               Celda = celda & ofv.gencelda("","tt","right","10%","","","","Fecha:","si")

               sfecha = now()

               sInput = ofv.GenerarInput(rs(9).name,sfecha,"hidden","","20","dt","SI")
               sInput = ofv.gencelda("","cc","","20%","","","",sInput & sfecha,"si")
               Celda = celda & sInput

               if rs(8) <> "SC" then

                  sInput = ofv.gencelda("","cc","","20%","","","","ENVIADO","si")
                  Celda = celda & sInput

               else

                  sInput = ofv.GenerarInput("bac1","Enviar","submit","","10","bt","")
                  sInput = ofv.gencelda("","cc","","20%","","","",sInput,"si")
                   Celda = celda & sInput

               end if

               sAccion = "asp/ActualizarEN.asp?opc=WFL&ID=" & rs(0) & "&Tabla=Objwfldet&volver="
               sAccion = sAccion & vuelta & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
               Formulario2 = ofv.GenForm(sForm & "4","",sAccion,"","post","",Celda,"si")
               Fila = ofv.GenRow("","","","","","","",Formulario2,"si")




               sHTML = sHTML & ofv.GenTabla("","aa  border=0","","100%","","0","0","1","0","#bbbbbb","",fila,"si")


           end if
           call ofv.cerrarconsulta(rs)



         case 1
  sinput = "No esta habilitado en la lista de autorizantes del presente Esquema Normativo"
  Celda = ofv.gencelda("","cc colspan=4","center","100%","","","",sinput,"si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","100%","","0","0","1","0","#bbbbbb","",fila,"si")

         case 2

  sinput = "Existen autorizantes sin confirmacion para el presente Esquema Normativo<br>"
  sinput = sinput & " con orden de autorizacion previo al suyo"
  Celda = ofv.gencelda("","cc colspan=4","center","100%","","","",sinput,"si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","100%","","0","0","1","0","#bbbbbb","",fila,"si")


  end select

  sHTML = sHTML & "</div>"


  end if




  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

if not cint(sestado) > 2 then response.write sHTML

%>