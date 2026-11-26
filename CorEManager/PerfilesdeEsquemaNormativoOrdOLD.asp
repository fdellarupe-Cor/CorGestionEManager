<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = request.querystring("CantRegMostrar")
  CantRegMoverx = request.querystring("CantRegMover")


  wperfil = request.querystring("wperfil")
  welem = request.querystring("welem")

  wOPC = request.querystring("OPC")

''response.write wOPC & " wOPC " & "<br>"

  if len(welem) = 0 then welem = 0


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""


  dim stabla(15)
  dim stablaE(15)
  dim stablaS(15)



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0


  Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     fila = ""


      Celda = ofv.GenCelda("","tt  ","","","","","","Perfil: ","si")
      sInput = rs(0) & "&nbsp;-&nbsp;" & rs(1)

      Celda = celda & ofv.GenCelda("","cc  ","","90%","","","",sInput,"si")

      sinput = "<a  href=PerfilesdeEsquemaNormativoProp.asp?CantRegMover=" & CantRegMoverx
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil & " ><b>Volver</b></a>"
      Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","aaa colspan=3  ","","","","","","&nbsp;","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      sInput = "&nbsp;Orden&nbsp;de&nbsp;Solapas"
      Celda = ofv.GenCelda("","tt colspan=2 ","","","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Formulario = Fila



'      sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

'      Formulario = ""

response.write wOPC & "<br>"

if wOPC= "SEQ" then

 winput = request.form(1) &""

response.write winput & " winput "& "<br>"

' Pasa valores a tabla de entrada
' Indice 1 = secuencia de entrada
' Indice 2 = valor en tabla del campo que se muestra
    dim stablaZS(16,16)
    xx = 0
    ii = 1
    yy = 0
    nn = 0
    for xx = 1 to 15
        if clng(stabla (xx)) < 99 then
           yy = yy + 1
           stablaE(stabla(xx)) = xx
           ii = xx
           nn = xx
          'response.write ii & " ii " & "<br>"
          'response.write stablaZS(xx,ii) & " xx + ii linea 104 " & "<br>"
           stablaZS(xx,nn) = xx
           stablaZS(nn,ii) = rs(xx + 1 + 30)
          'response.write xx & " xx " & "<br>"
          'response.write nn & " nn " & "<br>"
          'response.write stablaZS(xx,nn) & " stablaZ(xx,nn) " & "<br>"
          'response.write stablaZS(nn,ii) & " stablaZ(nn,ii) " & "<br>"
          'response.write rs(xx + 1 + 30) & " tyy linea 111 " & "<br>"
          'response.write yy & " yy " & "<br>"

          'response.write stablaZS(xx,ii) & " TablaZ(xx,ii) linea 114 " & "<br>"

        end if
    next
    tyy = yy

    xx = 0
    ii = 1
    yy = 0
    nn = 0

for xx = 2 to 16
    stabla (xx - 1) = rs(xx + 30)

  ''response.write stabla(xx) & " linea 95 " & "<br>"
  ''response.write xx & " xx - linea 96 " & "<br>"

next

   xmax = 15
   tyy  = 15
   xmin = 1
   xx  = 1

   zzy = 0

   dim stablaZZ(16,16)
   xx = 0
   ii = 0
   yy = 0
   for xx = 1 to 15
       if clng(stabla (xx)) < 99 then

          zzy = zzy + 1

         'response.write stablaZ(xx,ii) & " xx + ii linea 149 " & "<br>"

          ii = xx
          nn = xx

          stablaZZ(xx,nn) = xx
          stablaZZ(nn,ii) = rs(xx + 1 + 30)
         'response.write xx & " xx " & "<br>"
         'response.write rs(xx + 1 + 30) & " linea 157 " & "<br>"

        'response.write stablaZZ(xx,ii) & " TablaZZ(xx,ii) linea 159 " & "<br>"

       end if
   next
   tyy = yy


'Recorrida de ambas tablas para ver las diferencias

   comodin = winput
   xx = 1
   ii = 1

   for xx = 1 to xmax

    if xx + 1 <= 15 then
       ii = xx

      'response.write xx & " xx linea 177 " & "<br>"
      'response.write ii & " ii linea 178 " & "<br>"
      'response.write stablaZS(xx,ii)  & " linea 179 " & "<br>"
      'response.write stablaZZ(xx,ii) & " linea 180 " & "<br>"


    '' if stablae(xx) < comodin then
       if cstr(stablaZZ(xx,ii)& "") <> cstr(stablaZS(xx,ii)& "") then

        ''response.write stablaz(xx,ii) & " linea 186 " & "<br>"
        ''response.write stablazz(xx,ii) & " linea 187 " & "<br>"
          stablaE(xx,ii) = stablaZZ(xx,ii)
         'response.write stablaE(xx,ii) & " tablas -189- " & "<br>"

       else

        ''response.write stablaZZ(xx,ii) & " linea 186 " & "<br>"
        ''response.write stablaZS(xx,ii) & " linea 187 " & "<br>"

       end if

    end if

   Next

   response.write "va a hacer update objperfil - linea 204 - " & "<br>"

         Strsql = "UPDATE OBJPERFIL SET "

         for xx = 1 to tyy

             Strsql = strsql & rs(stablaZZ(xx) + 1 + 30).name & " = " & xx
             if xx <> tyy then Strsql = strsql & ", "

         next

         Strsql = strsql & " where tipoperfil = '" & wperfil & "' "
       ''response.write strsql & " UPDATE " & "<br>"
         call ofv.crearconsultaEx(StrSql,cn,1,parametros)


         Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
       ''response.write strsql & " SELECT " & "<br>"
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       ''if not rs.eof then

         tyy = 15
         ''tyy = yy + 1 + 15

else

    for xx = 2 to 16
        stabla (xx - 1) = rs(xx + 30)

    response.write rs(xx + 30) & " linea 223 " & "<br>"
    response.write stabla (xx - 1) & " linea 224 " & "<br>"
    response.write xx & " linea 167 " & "<br>"

    next

' Pasa valores a tabla de salida
' Indice 1 = secuencia de entrada
' Indice 2 = valor en tabla del campo que se muestra
    dim stablaZ(16,16)
    xx = 0
    ii = 1
    yy = 0
    nn = 0
    for xx = 1 to 15
      if clng(stabla (xx)) < 99 then
         yy = yy + 1
         stablaE(stabla(xx)) = xx
         ii = xx
         nn = xx
        'response.write ii & " ii " & "<br>"
        'response.write stablaZ(xx,ii) & " xx + ii linea 252 " & "<br>"
         stablaZ(xx,nn) = xx
         stablaZ(nn,ii) = rs(xx + 1 + 30)
        'response.write xx & " xx " & "<br>"
        'response.write nn & " nn " & "<br>"
         response.write stablaZ(xx,nn) & " stablaZ(xx,nn) " & "<br>"
         response.write stablaZ(nn,ii) & " stablaZ(nn,ii) " & "<br>"
        'response.write rs(xx + 1 + 30) & " tyy linea 259 " & "<br>"
        'response.write yy & " yy " & "<br>"

         response.write stablaZ(xx,ii) & " TablaZ(xx,ii) linea 262 " & "<br>"

      end if
    next
    tyy = yy

    response.write tyy & " tyy " & "<br>"

' aca ,modifique clng(welem) = 15
    if clng(welem) = 16 then welem = 0
'    response.write welem
    if clng(welem) > 0 then
'       if  clng(stabla(clng(welem) - 1)) < 3 then welem = 0
    end if

'response.write clng(welem) & " linea 181 " & "<br>"

    if clng(welem) > 0 then

       welemix = stabla(clng(welem) - 1)
       welemixa = clng(welemix) - 1
       welemixh = clng(welemix) - 2

       for xx = 1 to welemixh

           stablaS(xx) =  stablaE(xx)

       ''response.write stablaS(xx) & " linea 198 " & "<br>"
       ''response.write stablaE(xx) & " linea 199 " & "<br>"

       next

       stablaS(welemix) =  stablaE(welemixa)
       stablaS(welemixa) =  welem - 1

       for xx = (clng(welemix) + 1) to tyy

           stablaS(xx) =  stablaE(xx)

    '' response.write stablaS(xx) & " linea 219 " & "<br>"

       next


    '' response.write "va a hacer update objperfil - linea 215 - " & "<br>"
    '' response.write tyy & " tyy " & "<br>"
    '' response.write yy  & " yy " & "<br>"

       Strsql = "UPDATE OBJPERFIL SET "

       for xx = 1 to tyy

           Strsql = strsql & rs(stablaS(xx) + 1 + 30).name & " = " & xx
         ''response.write strsql & "<br>"
           if xx <> tyy then Strsql = strsql & ", "

       next

       Strsql = strsql & " where tipoperfil = '" & wperfil & "' "
     ''response.write strsql
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)


    else

       for xx = 1 to tyy

           stablaS(xx) =  stablaE(xx)

          response.write stablaE(xx)  & " stablaE(xx) " & "<br>"
          response.write stablaS(xx)  & " stablaS(xx) " & "<br>"

       next

    end if


  end if

 ''response.write stablaE(xx)  & " stablaE(xx) " & "<br>"
 ''response.write stablaS(xx)  & " stablaS(xx) " & "<br>"
 ''response.write stablaZ(xx,ii)  & " stablaZ(xx,ii) " & "<br>"

if wOPC= "SEQ" then

   response.write " Entro linea 343 " & "<br>"

   for xx = 1 to tyy

      stablaS(xx) =  stablaE(xx)

  'response.write stablaE(xx)  & " stablaE(xx) " & "<br>"
  'response.write stablaS(xx)  & " stablaS(xx) " & "<br>"

   next

end if

response.write xx & " xx " & "<br>"
yy = 0
response.write yy & " yy " & "<br>"
'tyy = 15
response.write tyy & " tyy " & "<br>"

response.write " Paso por aca - linea 356 " & "<br>"

     xx = 0

     response.write tyy & " tyy  linea 360 " & "<br>"
     response.write yy & " yy " & "<br>"


     for yy = 1 to tyy
         response.write tyy & " tyy  linea 368 " & "<br>"
         sform = "formxxx" & yy
         response.write yy & " yy " & "<br>"
         fila = ""

         xx = xx + 1

        ssInput = ofv.GenerarInput("S" & yy,yy * 10,"text",sForm,"10","dtn","")
      ''response.write ssInput & " ssInput " & "<br>"
        Celda = ofv.GenCelda("","cc  ","right","","","","",ssInput,"si")

        sInput = rs(stablaS(yy) + 1 + 15)
        response.write yy & " yy" & "<br>"
        response.write rs(stablaS(yy) + 1 + 15) & " rs(stablaS(yy) linea 381 " & "<br>"
        response.write rs(stablaS(yy) + 1)  & " rs(stablaS(yy) + 1) linea 382 " & "<br>"

   'if rs(stablaS(yy) + 1) then

      response.write xx & " xx linea 386 " & "<br>"

      if xx = 1 or xx = 82 then

         Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")

      '''Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;background-color:#008000;' ","center","","15%","","","-----","si")
         Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:none; ' ","center","","15%","","","","si")

      else

         Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")

       ''response.write (stablaS(yy) + 1) & " tabla-yy + 1 " & "<br>"
       ''response.write yy & " yy " & "<br>"

         sinput = ""
'''''''''sinput = "<a  href=PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx
'''''''''sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil    & "&Welem=" & (stablaS(yy) + 1) & "><b>Subir</b></a>"
'''''''''sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Wperfil=" & Wperfil & "&Welem=" & (stablaS(yy) + 1)


       ''Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;background-color:#993333;' ","center","","15%","","",sInput,"si")
         Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2; ' ","center","","15%","","",sInput,"si")

      end if

 'else

     ''Celda = celda & ofv.GenCelda("","cc ","","60%","","","",sInput,"si")

     ''Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2;' ","center","","15%","","","","si")

 'end if



     sAccion = ""
	   if bModificar then

	       sAccion = "PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx & "&OPC=SEQ"
           sAccion = sAccion  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil    & "&Welem=" & (stablaS(yy) + 1)
   end if

	  Fila = ofv.genform(sForm,"",sAccion,"","post","",Celda,"si") & chr(13)
	  Formulario = Formulario & ofv.genrow("","","","","","","",Fila,"si")

      posi = posi + 1


     next

  end if
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","2","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

'end if

Response.Write sHTML

%>