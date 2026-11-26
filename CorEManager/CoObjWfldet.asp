<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
function Esqnom(xtipo,xcodigo,ofv)

set xcn=ofv.conectar(ofv.strconn2)

Strsql="Select descri From tipobjeto where tipobjeto ='" & xtipo & "' "
set xrs=ofv.crearconsultaEx(StrSql,xcn,1,parametros)
if not xrs.eof then xdescri = xrs(0)
call ofv.cerrarconsulta(xrs)

Strsql="Select * From objetos where codobjeto =" & xcodigo
set xrs=ofv.crearconsultaEx(StrSql,xcn,1,parametros)
if not xrs.eof then 
   xiden = xrs(1)
   xobjdes = xrs(3)
end if
call ofv.cerrarconsulta(xrs)

call ofv.cerrarconn(xcn)

esqnom = ""
esqnom = esqnom & xdescri & ":&nbsp;" & xiden & "&nbsp;-&nbsp;" & xobjdes & ""

end function


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sVolver = request.querystring("volver")
  sCod = request.querystring("cod")
  vuelta = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")

  session("TCliente") = sTipo
  session("NCliente") = nCodigo


  set cn = ofv.conectar(ofv.strconn2)

  snombre = esqnom (sTipo,nCodigo,ofv)

    shtml = ofv.Repbody(Ucase("Circuito de Aprobacion de ") & snombre) & ofv.RepHeader2(Ucase("Circuito de Aprobacion de ") & snombre,"600")

snova = 0

if sCod = "" then
snova = 1
     Session("Segco") = "S"
     Session("Segce") = "S"

   Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipo & "' and codobjeto = " & nCodigo
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rs.eof then

      
      if Session("Usuario") <> rs(7) then

         if rs(5) then Session("Segco") = "N"
         if rs(6) then Session("Segce") = "N"

         call ofv.cerrarconsulta(rs)
         Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipo
         Strsql = Strsql & "' and codobjeto = "  & nCodigo
         Strsql = Strsql & " and ((grpusr = 'USR-" & Session("Usuario") & "') " 
         Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')) " 
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         do until rs.eof 
            
            select case rs(3)
                   case "1"
                        Session("Segco") = "S"
                        Session("Segce") = "S"
                   case "2"
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

 if Session("Segco") = "S" then snova = 0

end if

if snova = 1 then
   sHtml = sHtml & "<font face=tahoma,arial size=2 color=#336699>"   
   sHtml = sHtml & "<p align=center>El usuario: " & Session("usuario") & "&nbsp;-&nbsp" 
   sHtml = sHtml & Session("segusuario") & " no esta autorizado para acceder a este"
   sHtml = sHtml & " esquema normativo</p>"
   sHtml = sHtml & "</font>"
else 

''' Esquema nuevo Workflow ''''''''''''''''''''''''''''''''''''''''

  '''  response.write "tipo: " & session("TCliente") & " Codigo: " & session("NCliente") & "<br>"

     set cnW = ofv.conectar(ofv.strconn5)

     StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' ORDER BY COMIENZO DESC "
     set rs = ofv.crearconsultaEx(StrSql,cnW,1,parametros)
     if not rs.eof then
        response.redirect "CoHistWKF.asp"
     end if


''' Esquema viejo '''''''''''''''''''''''''''''''''''''''''''''

  Celda = ofv.gencelda("","vt colspan=6","center","2%","","","",Ucase("Lista de Autorizantes"),"si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  Celda = ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("Orden"),"si")
  Celda = Celda & ofv.gencelda("","vt","","50%","","#99aadd","",Ucase("Usuario"),"si")
  Celda = Celda & ofv.gencelda("","vt colspan=2","","2%","","#99aadd","",Ucase("Confirmacion"),"si")
  Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","",Ucase("Fecha"),"si")
  Celda = Celda & ofv.gencelda("","vt","","2%","","#99aadd","","&nbsp;","si")

  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  Strsql = "SELECT * FROM OBJwfldet  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' order by orden,objid"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

    sForm = "wfl" & rs(0)

      scolor = "<font face=corflaticon size=2 color=#336699>w</font>&nbsp;"

    'Orden
      sinput = ofv.gencelda("","ut","","1%","","","",scolor & rs(3),"si")
     Celda =  sinput

    sTabla = "Usuario"

    sBusca = rs(4)

    sGrpUsr = "S/D"
    set cx = ofv.conectar(ofv.strconn0)

    Strsql = "SELECT * FROM " & sTabla & "s"
    Strsql = Strsql & " Where ID = '" & sBusca & "' "

    set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
    if not rx.eof then sGrpUsr = rx(0) & " - " & rx(1)
    call ofv.cerrarconsulta(rx)
    call ofv.cerrarconn(cx)

    sinput = ofv.gencelda("","ut ","","50%","","#eeeeee","",sGrpUsr,"si")

    Celda = Celda & sinput

    select case Ucase(rs(8))
           case "SC"
                scolor = "<font face=corflaticon size=2 color=#f7d655>l</font>"
                sTexto = "Sin Confirmar"
           case "NC"
                scolor = "<font face=corflaticon size=2 color=#000080>l</font>"
                sTexto = "En Expiracion"
           case "NP"
                scolor = "<font face=corflaticon size=2 color=#000080>l</font>"
                sTexto = "Expirado"

           case "SI"
                scolor = "<font face=corflaticon size=2 color=#009900>l</font>"
                sTexto = "Confirmado"

           case "NO"
                scolor = "<font face=corflaticon size=2 color=#993333>l</font>"
                sTexto = "Rechazado"

    end select

    sinput = ofv.gencelda("","ut","","1%","","","",scolor,"si")

    Celda = Celda & sinput

    sinput = ofv.gencelda("","ut","","5%","","","",sTexto,"si")

    Celda = Celda & sinput

    sinput = ofv.gencelda("","ut","","15%","","","",rs(9),"si")

    Celda = Celda & sinput

    select case Ucase(rs(8))
           case "SC"
                scolor = "<img src=" & chr(34) & "imagenes/sema.gif" & chr(34) & " width=30% >"

           case "NC"
                scolor = "---"
           case "NP"
                scolor = "---"


           case "SI"
                scolor = "<img src=" & chr(34) & "imagenes/semv.gif" & chr(34) & " width=30% >"


           case "NO"
                scolor = "<img src=" & chr(34) & "imagenes/semr.gif" & chr(34) & " width=30% >"


    end select

    sinput = ofv.gencelda("","ut","center","15%","","","",scolor,"si")

    Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    if Session("Segce") = "S" then

       Celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")
   
       Celda = Celda & ofv.gencelda("","vt colspan=5","","50%","","","","Observaciones:","si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

       Celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")


               wtxt1 = rs(5)

               if len(wtxt1) = 0 then
                  stexto = "&nbsp;"
               else
                  stexto = wtxt1
               end if  

 

       sinput = ofv.gencelda("","ut ","","50%","","#eeeeee","",stexto,"si")
       Celda = Celda & sinput

       sinput = ofv.gencelda("","ut colspan=3","","1","","","","&nbsp;","si")
       Celda = Celda & sinput

       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

               wtxt1 = rs(6)

               if len(wtxt1) = 0 then
                  stexto = "&nbsp;"
               else
                  stexto = wtxt1
               end if  

       if len(wtxt1) > 0 then
'          stexto = rs(6)
          Celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")

          sinput = ofv.gencelda("","ut ","","50%","","#eeeeee","",stexto,"si")
          Celda = Celda & sinput

          sinput = ofv.gencelda("","ut colspan=4","","1","","","","&nbsp;","si")
          Celda = Celda & sinput

          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
       end if

               wtxt1 = rs(7)

               if len(wtxt1) = 0 then
                  stexto = "&nbsp;"
               else
                  stexto = wtxt1
               end if  

       if  len(wtxt1) > 0 then
   '       stexto = rs(7)
          Celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")

          sinput = ofv.gencelda("","ut ","","50%","","#eeeeee","",stexto,"si")
          Celda = Celda & sinput

          sinput = ofv.gencelda("","ut colspan=4","","1","","","","&nbsp;","si")
          Celda = Celda & sinput

          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
       end if

    end if

    sinput = "<hr size=1 color=#5588bb widht=100% >"

    Celda = ofv.gencelda("","vt colspan=6","","100%","","","",sinput,"si")
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
   rs.movenext
   loop

  sHTML = sHTML & "<br>"
  sHTML = sHTML & ofv.GenTabla("","aa  border=0","","600","","0","0","1","0","","",fila,"si")


  call ofv.cerrarconsulta(rs)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' '

end if

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>