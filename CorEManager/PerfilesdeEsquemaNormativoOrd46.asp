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

  Windice = request.querystring("Windice")

  wperfil = request.querystring("wperfil")
  welem = request.querystring("welem")

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
  sForm = "form" & posi

'1.
  if not rs.eof then
'1.
     fila = ""


      Celda = ofv.GenCelda("","tt  ","","","","","","Perfil: ","si")
      sInput = rs(0) & "&nbsp;-&nbsp;" & rs(1)

      Celda = celda & ofv.GenCelda("","cc  ","","90%","","","",sInput,"si")

      sinput = "<a  href=PerfilesdeEsquemaNormativoProp.asp?CantRegMover=" & CantRegMoverx
''''''sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Wperfil=" & Wperfil & " ><b>Volver</b></a>"
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Windice=" & Windice & "&Wperfil=" & Wperfil & " ><b>Volver</b></a>"



      Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")


      Celda = ofv.GenCelda("","aaa colspan=3  ","","","","","","&nbsp;","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "&nbsp;Orden&nbsp;de&nbsp;Solapas"

      Celda = ofv.GenCelda("","tt colspan=3 ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Formulario = Fila


'     sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

'     Formulario = ""


  '1f.
    for xx = 2 to 16
  '1f.
        stabla (xx - 1) = rs(xx + 30)
    next


    yy = 0

  '2f.
    for xx = 1 to 15
  '2f.

        if clng(stabla (xx)) < 99 then
           yy = yy + 1
           stablaE (stabla (xx)) = xx
        end if
    next
    tyy = yy
' aca ,modifique clng(welem) = 15
    if clng(welem) = 16 then welem = 0
   ' response.write welem
    if clng(welem) > 0 then
'''    if  clng(stabla(clng(welem) - 1)) < 3 then welem = 0
    end if

'2.
  if clng(welem) > 0 then
'2.

       welemix = stabla(clng(welem) - 1)
       welemixa = clng(welemix) - 1
       welemixh = clng(welemix) - 2

  '3f.
    for xx = 1 to welemixh
  '3f.
           stablaS(xx) =  stablaE(xx)

       next

       stablaS(welemix) =  stablaE(welemixa)
       stablaS(welemixa) =  welem - 1

  '4f.
    for xx = (clng(welemix) + 1) to tyy
  '4f.

           stablaS(xx) =  stablaE(xx)

       next

       Strsql = "UPDATE OBJPERFIL SET "

   '5f.
     for xx = 1 to tyy
   '5f.

           Strsql = strsql & rs(stablaS(xx) + 1 + 30).name & " = " & xx
           if xx <> tyy then Strsql = strsql & ", "

       next

       Strsql = strsql & " where tipoperfil = '" & wperfil & "' "
       'response.write strsql
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  '2.
    else
  '2.

       for xx = 1 to tyy

           stablaS(xx) =  stablaE(xx)

       next

  '2.
    end if
  '2.

  xx = 0

 '6f.
   for yy = 1 to tyy
 '6f.

      sform = "form" & yy

      fila = ""

      xx = xx + 1

    ''Celda = ofv.GenCelda("","cc  ","","","","","",xx & "º ","si")

    ''ssInput = xx & "0 "
      ssInput = ofv.GenerarInput("S" & yy,yy * 10,"text",sForm,"10","dtn","")

    ''response.write ssInput & " secuencia " & "<br>"

      Windice = "S" & yy * 10

    ''response.write Windice & " Windice " & "<br>"

      Celda = ofv.GenCelda("","cc  ","right","","","","",ssInput,"si")

      sInput = rs(stablaS(yy) + 1 + 15)

    ''response.write ssInput & sInput & " ssInput + sInput " & "<br>"

   '3.
    if rs(stablaS(yy) + 1) then
   '3.

   '4.
    if xx = 1 or xx = 82 then
   '4.

      Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")

''''''Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;background-color:#008000;' ","center","","15%","","","-----","si")
      Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2; ' ","center","","15%","","","-----","si")

   '4.
   else
   '4.

      Celda = celda & ofv.GenCelda("","tt ","","60%","","","",sInput,"si")

      sinput = "<a  href=PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx
''''''sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx   & "&Wperfil=" & Wperfil  & "&Welem=" & (stablaS(yy) + 1) & " ><b>Subir</b></a>"
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Windice=" & Windice  & "&Wperfil=" & Wperfil  & "&Welem=" & (stablaS(yy) + 1) & "-----"

''''''Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;background-color:#993333;' ","center","","15%","","",sInput,"si")
      Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2; ' ","center","","15%","","",sInput,"si")




   '4.
    end if
   '4.

 '3.
  else
 '3.

      Celda = celda & ofv.GenCelda("","cc ","","60%","","","",sInput,"si")

      Celda = Celda & ofv.gencelda("","cc style='border-style:outset;border-width:2;' ","center","","15%","","","-----","si")

 '3.
  end if
 '3.

  'Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

  next

  sAccion = ""
	   if bModificar then
	      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
	      sAccion = sAccion & "&tabla=Objperfil&volver=" & vuelta
	   end if
	   Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Formulario = Formulario & Fila

      posi = posi + 1


 '1.
  end if
 '1.

''''''''''''''''''''''''''''''''''''''''''''''''
' zcampotablass = ""
'  yy  = 1
'  yyy = 1

'  for yy = 1 to tyy

'    'response.write yy + 1 & " yy + 1 " & "<br>"
'    'response.write yy & " yy " & "<br>"

'      if yy > yy + 1 then

'         zcampotablass = yy + 1

'       ' response.write zcampotablass & " zcampotablass " & "<br>"

'         yyy = yy + 1

'         yyy = zcampotablass

'         yyy + 1 = yy
'       ' response.write yyy + 1 + 15 & " yyy + 1 por mayor " & "<br>"
'       ' response.write yyy + 1 + 30 & " yyy + 2 por mayor " & "<br>"

'      else

'	     if yy < yy + 1 then

'	        yyy = yy

'           'response.write yyy & " yyy por menor " & "<br>"

'         end if
'	  end if

'  if tyy > 15 then exit for
'
	  ''   yy = yy + 1

'     next

''''''''''''''''''''''''''''''''''''''''''''''''


 sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","2","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
'1.
end if
'1.

Response.Write sHTML

%>