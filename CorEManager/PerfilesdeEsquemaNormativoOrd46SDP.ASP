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

''response.write welem & " welem "

  if len(welem) = 0 then welem = 0


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""


  dim stabla(15)
  dim stablaE(15)
  dim stablaS(15)
''dim stablaSS(15)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0


  Strsql = "Select * From Objperfil where tipoperfil = '" & wperfil & "' "
''response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  sForm = "form" & posi

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

      Celda = ofv.GenCelda("","tt colspan=3 ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Formulario = Fila


   '  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

   '  Formulario = ""




    for xx = 2 to 16
        stabla (xx - 1) = rs(xx + 30)
    next


    yy = 0
    for xx = 1 to 15
        if clng(stabla (xx)) < 99 then
           yy = yy + 1
           stablaE (stabla (xx)) = xx
        end if
    next
    tyy = yy
'   response.write " yy " & yy & "<br>"

' aca ,modifique clng(welem) = 15
    if clng(welem) = 16 then welem = 0
 '''response.write welem
    if clng(welem) > 0 then
'       if  clng(stabla(clng(welem) - 1)) < 3 then welem = 0
    end if

    if clng(welem) > 0 then


       welemix = stabla(clng(welem) - 1)
       welemixa = clng(welemix) - 1
       welemixh = clng(welemix) - 2

      response.write welemix  & " welemix  -119- " & "<br>"
      response.write welemixa & " welemixa -120- " & "<br>"
      response.write welemixh & " welemixh -121- " & "<br>"

       for xx = 1 to welemixh

           stablaS(xx) =  stablaE(xx)

     ''   response.write stablaS(xx) & " stablaS(xx) -123- " & "<br>"

       next

       stablaS(welemix) =  stablaE(welemixa)
       stablaS(welemixa) =  welem - 1

       for xx = (clng(welemix) + 1) to tyy

           stablaS(xx) =  stablaE(xx)

       response.write stablaS(xx) & " " & xx & " stablaS(xx) -137- " & "<br>"

       next

       response.write tyy & " tyy "  & "<br>"
       response.write xx & " xx "  & "<br>"

   xx = 1

       response.write rs(stablaS(xx) + 1 + 30).name  & "<br>"

       Strsql = "UPDATE OBJPERFIL SET "

       for xx = 1 to tyy

       Strsql = strsql & rs(stablaS(xx) + 1 + 30).name & " = " & xx
       if xx <> tyy then Strsql = strsql & ", "

       next

       Strsql = strsql & " where tipoperfil = '" & wperfil & "' "
       response.write strsql
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)


    else

       for xx = 1 to tyy

           stablaS(xx) =  stablaE(xx)

       next

    end if


    xx = 0

    for yy = 1 to tyy


      fila = ""


         xx = xx + 1
      ''   sInput = xx

 '''''Celda = ofv.GenCelda("","cc  ","","","","","",xx & "º ","si")

    ''response.write " xx " & xx & "<br>"
    ''response.write   rs(stablaS(yy) + 1 + 15)   & " 1 + 15 "  & "<br>"
    ''response.write   rs(stablaS(yy) + 1 + 30)   & " 1 + 30 "  & "<br>"
      response.write   rs(stablaS(yy) + 2) & " 2 " & "<br>"

      sInput = rs(stablaS(yy) + 1 + 15)

    ''response.write sInput & " sInput " & "<br>"

      if rs(stablaS(yy) + 2) then

    ''response.write rs(stablaS(yy) + 2) & " rs(stablaS(yy) + 2) " & "<br>"

        if xx = 1 or xx = 82 then

           ssInput = xx & "0 "
           Celda = ofv.GenCelda("","cc  ","right","","","","",ssInput,"si")

           Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")

        '''response.write sInput & " sInput " & "<br>"

     ''''''Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2;background-color:#008000;' ","center","","15%","","","-----","si")
           Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2; ' ","center","","15%","","","-----","si")

        else

'''''Nuevo agregado 19112014 - Version 4.6
'
           ssInput = xx & "0 "

         IF BMODIFICAR THEN
     ''''''response.write bmodificar & " bmodificar " & "<br>"

           ssInput = ofv.GenerarInput(rs(stablaS(yy) + 1 + 30).name,rs(stablaS(yy) + 1 + 30).value & "0 ","text",sForm,"10","dtn","")
           response.write rs(stablaS(yy) + 1 + 30).name & " name " & "<br>"
           response.write rs(stablaS(yy) + 1 + 30).value & " value " & "<br>"

     ''''''response.write xx & " xx " & "<br>"
     ''''''response.write yy & " yy " & "<br>"

     ''''''Celda = ofv.GenCelda("","cc style='border-style:outset;border-width:0;' ","right","","","","",ssInput,"si")
           Celda = ofv.GenCelda("","tt  ","right","","","","",ssInput,"si")

         ELSE

           Celda = ofv.GenCelda("","tt  ","right","","","","",ssInput,"si")

         END IF
'
'''''Nuevo agregado 19112014 - Version 4.6 FIN

           Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")

           sinput = "<a  href=PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx
'''''''''''sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil    & "&Welem=" & (stablaS(yy) + 1) & " ><b>Subir</b></a>"
           sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil    & "&Welem=" & (stablaS(yy) + 1) & " ><b>-----</b></a>"

           response.write  Wperfil & " Wperfil "  & chr(13) & "<br>"
           response.write  (stablaS(yy) + 1) & " stablaS + 1 "  & chr(13) & "<br>"
       ''''response.write sInput & " sInput linea 231 " & chr(13)

      '''''Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;background-color:#993333;' ","center","","15%","","",sInput,"si")
           Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2; ' ","center","","15%","","",sInput,"si")


        end if

      else

        ssInput = xx & "0 "
        Celda = ofv.GenCelda("","cc  ","right","","","","",ssInput,"si")

        Celda = celda & ofv.GenCelda("","cc ","","60%","","","",sInput,"si")
    ''''response.write sInput & " sInput " & "<br>"
        Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2;' ","center","","15%","","","-----","si")

      end if

 ''Next

 '''' response.write ssInput & " sInput linea 250 " & chr(13)

      'Para cambiar orden de los registos en la presentación de la pantalla

	  'salida normal
	  '   tyyaux = ""
	  '    zcampotablass = ""
	  '    yy = 1
	  '    yyy = 1

	  '   for yy = 1 to tyy

	  '  'tyy = maximo ocurrencias

	  '         response.write rs(stablaS(yy) + 1) & " stablaS + 1 "  & chr(13)

	  '          if rs(stablaS(yy)) > rs(stablaS(yy) + 1) then

	  '             zcampotablass = rs(stablaS(yy) + 1)
	  '             response.write zcampotablass & " zcampotablass "  & chr(13)

	  '     '''''''rs(stablaS(yyy)) = rs(stablaS(yy) + 1)
	  '             rs(stablaS(yyy)) = zcampotablass

	  '             rs(stablaS(yyy) + 1) = rs(stablaS(yy))
   	  '             response.write rs(stablaS(yyy) + 1 + 15) & " stablaS + 1 por mayor " & chr(13)
	  '             response.write rs(stablaS(yyy) + 1 + 30) & " stablaS + 2 por mayor " & chr(13)

	  '          else
	  '              if rs(stablaS(yy)) < rs(stablaS(yy) + 1) then
	  '                rs(stablaS(yyy) + 1) = rs(stablaS(yy))

	  '                response.write rs(stablaS(yyy) + 1) & " stablas + 1 por menor " & chr(13)

	  '             end if
	  '          end if

	  '     if tyy > 15 then exit for

	  '      yy = yy + 1


      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Formulario = Formulario & Fila

     next

  end if



  '''''Nuevo agregado 19112014 - Version 4.6
  '
  sAccion = ""
  if bModificar then
  '''response.write bmodificar & " bmodificar " & "<br>"
     sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
     sAccion = sAccion & "&tabla=Objperfil&volver=" & vuelta
  end if
     Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
  '
  '''''Nuevo agregado 19112014 - Version 4.6 FIN


  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","2","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>