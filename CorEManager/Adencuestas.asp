<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Fila = ofv.Encabezados("Codigo","imp","vigencia","mensaje","EXPIRA","dist","env","&nbsp;","lista","envio")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From emencuestas Order By encfecinicio desc, objid desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi< CantRegMostrarx)
    Formulario = ""
    sForm = "form" & posi

     clase = "ut" 

    'Eliminar
   if bSuprimir then
     sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_") 
     sHref = sHref & "&Tabla=emencuestas&volver=" & vuelta & "&Eliminar=NSNC" 
     Celda = ofv.BotonEliminar("Mensaje: NGP" & cstr(rs(0)),sHref) 
  else 
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si") 
  end if

    'codigo
    sInput = "NEC" & rs(0)
    Celda = Celda & ofv.GenCelda("",clase,"","5%","","","",sInput,"si")

    'Importancia

   select case clng(rs(7))
	  case 0

		sxopc4x = "<font color=#800000 size=4><b>!</b></font>"
	  case 1

		sxopc4x = "<font color=#008000 size=4><b>*</b></font>"
	  case 2

		sxopc4x = "<font color=#000080 size=4><b>+</b></font>"

    end select

    sInput = sxopc4x
    Celda = Celda & ofv.GenCelda("",clase,"center","5%","","","",sInput,"si")

    'Fecha Vigencia
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(3))
    end if

    sInput = dFechan & "&nbsp;"
    Celda = Celda & ofv.GenCelda("",clase,"","30%","","","","<b>" & sInput & "</b>","si")

    'titulo
    if len(rs(17)) > 30 then
       sInput = left(rs(17),30) & "..."
    else
       sInput = rs(17)
    end if
    Celda = Celda & ofv.GenCelda("",clase,"","40%","","","",sInput,"si")

    'dias
    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(4))
    end if

    sInput = dFechan & "&nbsp;"
    Celda = Celda & ofv.GenCelda("",clase,"right","5%","","","",sInput,"si")

    'lista
    if cint(rs(8)) = 1 then
       sInput = "TODOS"
    else
       sInput = "LISTA"
    end if

    Celda = Celda & ofv.GenCelda("",clase,"","5%","","","",sInput,"si")

    'enviado
    if rs(14) <> "S" then
       sInput = "NO"
    ELSE
       sInput = "SI"
    END IF
    Celda = Celda & ofv.GenCelda("",clase,"","5%","","","",sInput,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=emencuestas&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")







       sInput = ofv.GenerarInput("","PROPIEDADES","submit","","10","bt","")
       Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")

       sAccion = "AdencuestasProp.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0)
       Formulario = Formulario & ofv.GenForm("Boton" & rs(0),"",sAccion,"","post","",Celda,"si")





    if cint(rs(8)) = 2 then
       sInput = ofv.GenerarInput("","LISTA","submit","","10","bt","")
       Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")
       sAccion = "Adencuestasdist.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0)
       Formulario = Formulario & ofv.GenForm("Boton" & rs(0),"",sAccion,"","post","",Celda,"si")
    else
       sInput = "---"
       Celda = ofv.GenCelda("",clase,"center","10%","","","",sInput,"si")
       Formulario = Formulario & Celda
    end if

    if rs(14) <> "S" then
       sInput = ofv.GenerarInput("","ENVIAR","submit","","10","bt","")
       Celda = ofv.GenCelda("","","","10%","","","",sInput,"si")
       sAccion = "AdencuestasDESP.asp?Retorno=" & CantRegMoverx & "&ID=" & rs(0)
       Formulario = Formulario & ofv.GenForm("Boton" & rs(0),"",sAccion,"","post","",Celda,"si")
    else
       sInput = "---"
       Celda = ofv.GenCelda("",clase,"center","10%","","","",sInput,"si")
       Formulario = Formulario & Celda
    end if

    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    posi=posi+1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","0","0","0","","",Fila,"si")

  Agregar = ""
  if bAgregar then
    sAccion = "Agregarencuestas.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=emencuestas"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
 ' str1 = "Select descri From tema Order By descri "    ' orden pagina
 ' str2 = "Select descri From tema Order By descri "    ' orden alfabetico
  'Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
  buscar = ""
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Function Fechador(wkfecha)


      wkfecha=clng(wkfecha)
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>