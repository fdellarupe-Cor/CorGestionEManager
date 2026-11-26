<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

set ofr = server.createobject("corscr.Corrender")

ofr.verfonts

  sxopc1 = ""
for ii=0 to ofr.wscreen.fontcount -1
 ' sxopc1 = sxopc1 & "<option value=" & chr(34) & Ucase(ofr.wscreen.fonts(ii)) & chr(34) & " >" & Ucase(ofr.wscreen.fonts(ii)) & "</option> "
  sxopc1 = sxopc1 & "<option value=" & chr(34) & ofr.wscreen.fonts(ii) & chr(34) & " >" & replace(ofr.wscreen.fonts(ii)," ","&nbsp;") & "</option> "
next

'  sxopc1 = "<option value=Tahoma>Tahoma</option> "
'  sxopc1 = sxopc1 & "<option value=Verdana>Verdana</option> "
'  sxopc1 = sxopc1 & "<option value=Helvetica>Helvetica</option> "
'  sxopc1 = sxopc1 & "<option value=Times_new_roman>Times&nbsp;new&nbsp;roman</option> "
'  sxopc1 = sxopc1 & "<option value=Comic_Sans>Comic&nbsp;Sans</option> "

  sxopc2 = "<option value=8pt>8pt</option>"
  sxopc2 = sxopc2 & "<option value=9pt>9pt</option>"
  sxopc2 = sxopc2 & "<option value=10pt>10pt</option>"
  sxopc2 = sxopc2 & "<option value=12pt>12pt</option>"
  sxopc2 = sxopc2 & "<option value=14pt>14pt</option>"
  sxopc2 = sxopc2 & "<option value=16pt>16pt</option>"
  sxopc2 = sxopc2 & "<option value=18pt>18pt</option>"
  sxopc2 = sxopc2 & "<option value=20pt>20pt</option>"

  set cn = ofv.conectar(ofv.strconn2)
  sxopc3 = ""
  ii=0
  redim preserve ccol(0)
  redim preserve dcol(0)
  strsql = "Select * From Syscolor"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     do until rs.eof
       redim preserve ccol(ii)
       redim preserve dcol(ii)
       ccol(ii) = rs(2)
       dcol(ii) = rs(1)
       sxopc3 = sxopc3 & "<option value=" & chr(34) & rs(2) & chr(34) & ">" & rs(1)
       sxopc3 = sxopc3 & "</option> "
       ii = ii + 1
     rs.movenext
     loop
 end if
  call ofv.cerrarconsulta(rs)

 ' ofv.ObtenerAtributos Session("Form"),""
 ' bAgregar = ofv.AAgregar
 ' bModificar = ofv.AModificar
 ' bSuprimir = ofv.AEliminar
 ' bConsultar = ofv.AConsultar


  'sHTML = ofv.FormHeader(session("FormN"))
 response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
 response.write "<tr>"
 response.write "<td width=100% >"
  response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
 response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
 response.write "<tr>"
 response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
 response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
 response.write "</tr></table>"





 CantRegMostrarx = 1
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From TABPERFIL Order By tabdescri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)

  sHTML = sHTML & "<table width=80%  marginwidth=0 cellspacing=" & ofv.wspacing
  sHTML = sHTML & " cellpadding=" & 1 + ofv.wspacing & ">"
  sHTML = sHTML & "<tr><td width=20% class=tt>Perfil</td>"
  sHTML = sHTML & "<td colspan=9 width=75% class=tt>Descripcion</td></tr>"

    'sHTML = sHTML & "<form name=form" & posi & " action=asp/actualizar.asp?fname="
    'sHTML = sHTML & ofv.convertircar(rs(0)," ","*") & "&tabla=tabperfil&volver=" & vuelta
    'sHTML = sHTML & " method=post>"

    sHTML = sHTML & "<tr>"
  '  'Boton Eliminar
  '  sHelp = rs(1).value
  '  sHelp = ofv.convertircar(sHelp," ","&nbsp;")
  '  sHTML = sHTML & "<A HREF=asp/EliminarRegistro.asp?ID=" & replace(rs(0)," ","_")
  '  sHTML = sHTML & "&Tabla=tabperfil&volver=" & vuelta & "&Eliminar=NSNC&Nombre=" & sHelp
  '  sHTML = sHTML & " Title=Eliminar&nbsp;" & sHelp & ">"
  '  sHTML = sHTML & "<img src=imagenes/botdel.jpg border=0 height=15 width=15></a></td>"

    'Perfil

    sInput = rs(0)
    nperfil = rs(0)
    sHTML = sHTML & ofv.gencelda("","vt","","50","","","",sInput,"si")


    'Descripcion
    sHTML = sHTML & "<td colspan=8 class=ut>"
    sHTML = sHTML & rs(1) 'ofv.GenerarInput(rs(1).name,rs(1),"text","form" & posi,"115","dt","")
    sHTML = sHTML & "</td></tr>"

    'Combos de perfiles
    sHTML = sHTML & "<tr><td>&nbsp;</td><td width=2% class=tu>Bordes "
    if rs(2) then
     s="Si"
    else
     s="No"
    end if
    sHTML = sHTML & "&nbsp;" & s & "</td>"
    sHTML = sHTML & "<td colspan=8 class=tu><u>Bloques</u></td></tr>"
    sHTML = sHTML & "<tr><td colspan=2>&nbsp;</td><td width=2% class=tu>Color</td><td width=25%   >"

    for ii=0 to ubound(ccol)
      if ccol(ii) = rs(3) then
       sHTML = sHTML &  dcol(ii)
      exit for
      end if
    next


    sHTML = sHTML & "</td><td width=2% class=tu>Fondo</td><td width=2% >"
    'sHTML = sHTML & "<select class=at name=" & rs(4).name & " onchange=document.form" & posi
    'sHTML = sHTML & ".submit();>"
    for ii=0 to ubound(ccol)
      if ccol(ii) = rs(4) then
       color =  dcol(ii)
      exit for
      else
       color = dcol(0)
      end if
    next
    sHTML = sHTML & color
    sHTML = sHTML & "</td><td width=2% class=tu>Font</td><td width=2% >"
    'sHTML = sHTML & "<select class=at name=" & rs(5).name & " onchange=document.form" & posi
    'sHTML = sHTML & ".submit();>"
    if not isnull(rs(5)) then
      sHTML = sHTML & replace(rs(5)," ","&nbsp;")' "<option value=" & chr(34) & rs(5) & chr(34) & " selected>" & replace(rs(5)," ","&nbsp;")  & "</option>"
    end if
    sHTML = sHTML & "</td><td width=2% class=tu>Size</td><td width=2% >"
    'sHTML = sHTML & "<select class=at name=" & rs(6).name & " onchange=document.form" & posi
    'sHTML = sHTML & ".submit();>"
    if not isnull(rs(6)) then
      sHTML = sHTML & rs(6) '"<option value=" &  rs(6)  &  " selected>"  & rs(6)  &  "</option>"
    end if
    sHTML = sHTML  & "</td></tr><tr><td>&nbsp;</td><td width=2% class=tu>"
    sHTML = sHTML & "Columnas&nbsp;"
    sHTML = sHTML & rs(7) 'ofv.GenerarInput(rs(7).name,rs(7),"text","form" & posi,"2 ","dt","si")
    sHTML = sHTML & "</td><td colspan=8  class=tu><u>Titulos de Columnas</u></td></tr>"
    sHTML = sHTML & "<tr><td colspan=2>&nbsp;</td><td width=2% "
    sHTML = sHTML & "class=tu>Color<td width=2% >"
    'sHTML = sHTML & "<select class=at name=" & rs(8).name & " onchange=document.form" & posi
    'sHTML = sHTML & ".submit();>"
    for ii=0 to ubound(ccol)
      if ccol(ii) = rs(8) then
      'sHTML = sHTML & "<option value=" & chr(34) & ccol(ii) & chr(34) & " selected>"
      sHTML = sHTML &  dcol(ii) '& "</option>"
      exit for
      end if
    next
    'sHTML = sHTML & sxopc3 & "</select>"
    sHTML = sHTML & "</td><td width=2% class=tu>Fondo</td><td width=2% >"
    'sHTML = sHTML & "<select class=at name=" & rs(9).name & " onchange=document.form" & posi
    'sHTML = sHTML & ".submit();>"
    for ii=0 to ubound(ccol)
      if ccol(ii) = rs(9) then
    '  sHTML = sHTML & "<option value=" & chr(34) & ccol(ii) & chr(34) & " selected>"
      color =   dcol(ii) '& "</option>"
      exit for
      else
      color=dcol(0)
      end if
    next
    sHTML = sHTML & color 'sxopc3 & "</select>"
    sHTML = sHTML & "</td><td width=2% class=tu>Font</td><td width=2% >"
   ' sHTML = sHTML & "<select class=at name=" & rs(10).name & " onchange=document.form" & posi
   ' sHTML = sHTML & ".submit();>"
    if not isnull(rs(10)) then
      sHTML = sHTML & replace(rs(10)," ","&nbsp;") '"<option value=" & chr(34) & rs(10) & chr(34) & " selected>" & replace(rs(10)," ","&nbsp;")  & "</option>"
    end if
    sHTML = sHTML & "</td><td width=2% class=tu>Size</td><td width=2% >"
   ' sHTML = sHTML & "<select class=at name=" & rs(11).name & " onchange=document.form" & posi
   ' sHTML = sHTML & ".submit();>"
    if not isnull(rs(11)) then
      sHTML = sHTML & rs(11) '"<option value=" &  rs(11)  &  " selected>" &  rs(11)  &  "</option>"
    end if
    sHTML = sHTML & "</td></tr><tr><td colspan=2>&nbsp;</td>"
    sHTML = sHTML & "<td colspan=8  class=tu><u>Contenido de Columnas</u></td></tr>"
    sHTML = sHTML & "<tr><td colspan=2>&nbsp;</td><td width=2% class=tu>Color<td width=2% >"
   ' sHTML = sHTML & "<select class=at name=" & rs(12).name & " onchange=document.form" & posi
   ' sHTML = sHTML & ".submit();>"
    for ii=0 to ubound(ccol)
      if ccol(ii) = rs(12) then
   '   sHTML = sHTML & "<option value=" & chr(34) & ccol(ii) & chr(34) & " selected>"
      color =   dcol(ii) '& "</option>"
      exit for
      else
      color=dcol(0)
      end if
    next
    sHTML = sHTML & color 'sxopc3 & "</select>"
    sHTML = sHTML & "</td><td width=2% class=tu>Fondo</td><td width=2% >"
    'sHTML = sHTML & rs(13) '"<select class=at name=" & rs(13).name & " onchange=document.form" & posi
    'sHTML = sHTML & ".submit();>"
    for ii=0 to ubound(ccol)
      if ccol(ii) = rs(13) then
      'sHTML = sHTML & "<option value=" & chr(34) & ccol(ii) & chr(34) & " selected>"
      color =  dcol(ii) '& "</option>"
      exit for
      else
      color=dcol(0)
      end if
    next
    sHTML = sHTML & color 'sxopc3 & "</select>"
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% class=tu>Font</td><td width=2% >"
    'sHTML = sHTML & "<select class=at name=" & rs(14).name & " onchange=document.form" & posi & ".submit(); >"
    if not isnull(rs(14)) then
      sHTML = sHTML & replace(rs(14)," ","&nbsp;") '"<option value=" & chr(34) & rs(14) & chr(34) & " selected>" & replace(rs(14)," ","&nbsp;")  & "</option>"
    end if
   ' sHTML = sHTML '& sxopc1 & " </select>"
    sHTML = sHTML & "</td><td width=2% class=tu>Size</td><td width=2% >"
   ' sHTML = sHTML & "<select class=at name=" & rs(15).name & " onchange=document.form" & posi
   ' sHTML = sHTML & ".submit();>"
    if not isnull(rs(15)) then
      sHTML = sHTML & rs(15) '"<option value=" &  rs(15)  &  " selected>" &  rs(15)  &  "</option>"
    end if

    sHTML = sHTML &"</td></tr>"

    'sHTML = sHTML & "</form>"
    sHTML = sHTML & "</table><br><center>"


    sHTML = sHTML & "<table width=50% marginwidth=0 cellspacing=2"
    sHTML = sHTML & " cellpadding=2 style={border-style:ridge;border-width:3;} >"


    'sHTML = sHTML & "<form name=form" & posi & "c action=asp/actualizar.asp?fname="
    'sHTML = sHTML & ofv.convertircar(rs(0)," ","*") & "&tabla=tabperfil&volver=" & vuelta
    'sHTML = sHTML & " method=post>"

    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td align=center class=tt colspan=6 style={border-style:ridge;border-width:3;} >"
    sHTML = sHTML & "Configuracion de Columnas"
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"

    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td class=tt style={border-style:ridge;border-width:3;} >Columna</td>"
    sHTML = sHTML & "<td class=tt style={border-style:ridge;border-width:3;} >Ancho</td>"
    sHTML = sHTML & "<td class=tt style={border-style:ridge;border-width:3;} >Alinear</td>"
    sHTML = sHTML & "<td class=tt style={border-style:ridge;border-width:3;} >Columna</td>"
    sHTML = sHTML & "<td class=tt style={border-style:ridge;border-width:3;} >Ancho</td>"
    sHTML = sHTML & "<td class=tt style={border-style:ridge;border-width:3;} >Alinear</td>"
    sHTML = sHTML & "</tr>"


    'xcm1 = "<select class=at name=NN onchange=document.form" & posi & "c.submit();>"
    'xcm1 = xcm1 &  "<option value=0 >CEN</option>"
   'xcm1 = xcm1 &  "<option value=1 >IZQ</option>"
   'xcm1 = xcm1 &  "<option value=2 >DER</option>"
   'xcm1 = xcm1 &  "</select>"

    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td width=2% class=tu align=right >1</td><td width=2% >"
    sHTML = sHTML & rs(16) 'ofv.GenerarInput(rs(16).name,rs(16),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(16)) > 0 then
   ' XCM2 = replace(xcm1,"name=NN","name=" & rs(26).name)
   ' XCM2 = replace(xcm2,"value=" & cstr(rs(26)),"value=" & cstr(rs(26)) & " selected ")
    select case clng(rs(26))
         case 0
              alg="CEN"
         case 1
              alg="IZQ"
         case 1
              alg="DER"
         case else
          alg="&nbsp;"
    end select

    else
      XCM2 = "&nbsp;"
    end if



    xcm2=alg
    sHTML = sHTML & XCM2 '& rs(26) & 222222
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% class=tu align=right >2</td><td width=2% >"
    sHTML = sHTML & rs(17) 'ofv.GenerarInput(rs(17).name,rs(17),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(17)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(27).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(27)),"value=" & cstr(rs(27)) & " selected ")
    select case clng(rs(27))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
     xcm2=alg

    else
      XCM2 = "&nbsp;"
    end if

    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"

    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td width=2% class=tu align=right >3</td><td width=2% >"
    sHTML = sHTML & rs(18) 'ofv.GenerarInput(rs(18).name,rs(18),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(18)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(28).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(28)),"value=" & cstr(rs(28)) & " selected ")
    select case clng(rs(28))
         case 0
              alg="CEN"
         case 1
              alg="IZQ"
         case 1
              alg="DER"
         case else
          alg="&nbsp;"
    end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% class=tu align=right >4</td><td width=2% >"
    sHTML = sHTML & rs(19) 'ofv.GenerarInput(rs(19).name,rs(19),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(19)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(29).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(29)),"value=" & cstr(rs(29)) & " selected ")
    select case clng(rs(29))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"


    sHTML = sHTML & "<tr><td width=2% class=tu align=right >5</td>"
    sHTML = sHTML & "<td width=2% >"
    sHTML = sHTML & rs(20) 'ofv.GenerarInput(rs(20).name,rs(20),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(20)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(30).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(30)),"value=" & cstr(rs(30)) & " selected ")
    select case clng(rs(30))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% class=tu align=right >6</td><td width=2% >"
    sHTML = sHTML & rs(21) 'ofv.GenerarInput(rs(21).name,rs(21),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(21)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(31).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(31)),"value=" & cstr(rs(31)) & " selected ")
    select case clng(rs(31))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"

    sHTML = sHTML & "<tr>"
    sHTML = sHTML & "<td width=2% class=tu align=right >7</td><td width=2% >"
    sHTML = sHTML & rs(22) 'ofv.GenerarInput(rs(22).name,rs(22),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(22)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(32).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(32)),"value=" & cstr(rs(32)) & " selected ")
    select case clng(rs(32))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% class=tu align=right >8</td><td width=2% >"
    sHTML = sHTML & rs(23) 'ofv.GenerarInput(rs(23).name,rs(23),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(23)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(33).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(33)),"value=" & cstr(rs(33)) & " selected ")
    select case clng(rs(33))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"

    sHTML = sHTML & "<tr><td width=2% class=tu align=right >9</td>"
    sHTML = sHTML & "<td width=2% >"
    sHTML = sHTML & rs(24) 'ofv.GenerarInput(rs(24).name,rs(24),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(24)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(34).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(34)),"value=" & cstr(rs(34)) & " selected ")
    select case clng(rs(34))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% class=tu align=right >10</td><td width=2% >"
    sHTML = sHTML & rs(25) '0fv.GenerarInput(rs(25).name,rs(25),"text","form" & posi & "c","2 ","dt","si")
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "<td width=2% align=center>"
    if clng(rs(25)) > 0 then
    'XCM2 = replace(xcm1,"name=NN","name=" & rs(35).name)
    'XCM2 = replace(xcm2,"value=" & cstr(rs(35)),"value=" & cstr(rs(35)) & " selected ")
    select case clng(rs(35))
             case 0
                  alg="CEN"
             case 1
                  alg="IZQ"
             case 1
                  alg="DER"
             case else
              alg="&nbsp;"
        end select
    xcm2=alg
    else
      XCM2 = "&nbsp;"
    end if
    sHTML = sHTML & XCM2
    sHTML = sHTML & "</td>"
    sHTML = sHTML & "</tr>"



    'sHTML = sHTML & "</form>"
    sHTML = sHTML & "</table></center><br>"

    posi = posi + 1
    rs.Movenext
  Loop



  'Boton par agregar un registro
 ' agregar = "<form name=form" & posi & " action=AgregarPerfilesdeTablas.asp?"
 ' agregar = agregar & "Regisvuelta=" & CantRegMoverx & "&tabla=tabperfil method=post>"
 ' agregar = agregar & "<td><input class=bt type=submit size=10% value=Agregar></td></form>"

  'Boton test
 ' test = "<form name=form" & posi & "test action=PerfilesdeTablasTest.asp?"
 ' test = test & "nperfil=" & nperfil & " target=_blank method=post>"
 ' test = test & "<td><input class=bt type=submit size=10% value=Test></td></form>"


 ' str1 = "Select tabdescri From TABPERFIL Order By tabdescri"    ' orden pagina
 ' str2 = "Select tabdescri From TABPERFIL Order By tabdescri"    ' orden alfabetico
 ' Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
 ' sinput = "<a  href=PerfilesdeTablasCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
 ' Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

 ' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,test & celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>