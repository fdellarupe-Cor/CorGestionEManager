<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  GLOS = request.querystring("GLOS")
  IF GLOS <> "" THEN
     SESSION("GLOS") = GLOS
  ELSE
     GLOS = SESSION("GLOS")
  END IF

  VMOVER = request.querystring("VMOVER")

  IF VMOVER <> "" THEN
     SESSION("VMOVER") = VMOVER
  ELSE
     VMOVER = SESSION("VMOVER")
  END IF


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  opc = request.querystring("opc")

  set cn = ofv.conectar(ofv.strconn2)

  if opc = "SI" then


           GLOSCOD = GLOS
           GLOSWRD = "Nueva Descripcion"
           GLOSFEC = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
              xxxx = GETFECH()
           GLOSHOR = XLHORDAT

           strsql = "INSERT INTO GLOSDSC (CODWORD,FECVIG,HORVIG,DSCWORD) "
           strsql = strsql & " VALUES ('" & GLOSCOD & "','" & GLOSFEC & "','" & GLOSHOR & "','" & GLOSWRD & "') "
           call ofv.crearconsultaEx(StrSql,cn,1,volver)


           strsql = "select OBJID from GLOSDSC where CODWORD = " & GLOSCOD & " AND FECVIG = " & GLOSFEC & " AND HORVIG = " & GLOSHOR
           set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
           if NOT rs.eof then

              GLOSDSC = RS(0)
              GLOSUSR = SESSION("USUARIO")


              strsql = "INSERT INTO GLOSUSR (CODWORD,CODDSC,USRWORD,TIPMOD,FECMOD,HORMOD) "
              strsql = strsql & " VALUES ('" & GLOSCOD & "','" & GLOSDSC & "','" & GLOSUSR & "','A','" & GLOSFEC & "','" & GLOSHOR & "') "
              call ofv.crearconsultaEx(StrSql,cn,1,volver)




           END IF



  end if


  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)




  'Encabezados
  Formulario = ofv.Encabezados("DESCRIPCION","FECHA","HORA","ACTUALIZAR")



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  Strsql = "Select * From GLOSDSC WHERE CODWORD = " & GLOS & " Order By FECVIG"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=GLOSDSC&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("DESCRIPCION",sHref)

    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Tipo
    wstexto = rs(4)

      if isnull(wstexto)  then
'         response.write "longv " & len(wstexto) & " " & wstexto & "<br>"
         sInput = space(80)
         sinput = replace(sinput," ","-")
      elseif len(wstexto) > 100 then
 '        response.write "longg " & len(wstexto) & " " & wstexto & "<br>"
         sinput = left(wstexto,100) & "<font face='tahoma' size=1 color=#993333 ><b>...(+)</b></font>"
      else
         if wstexto = space(len(wstexto)) then wstexto = replace(wstexto," ","-")
 '        response.write "longc " & len(wstexto) & " " & wstexto & "<br>"
         sinput = wstexto
      end if



    Celda = Celda & ofv.gencelda("","ut","","80%","","","",sInput,"si")


    sInput = cdate(cdbl(RS(2)) + clng(cdate("01-01-1900")))
    sinput = fechaArg(sInput)
    Celda = Celda & ofv.gencelda("","ut","","5%","","","",sInput,"si")

    sInput = GETGLBHMS(RS(3))


    sInput = GLBSH & ":" & GLBSM & ":" & GLBSS

    Celda = Celda & ofv.gencelda("","ut","","5%","","","",sInput,"si")

    'sImagenP ="&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='GLOSMANAGERDSCEXT.asp?GLOSDSC=" & rs(0) & "&CantRegMover=" & CantRegMoverx
    sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    sinput = sinput  & "' title='Actualizar.' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=GLOSDSC&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  sHTML = sHTML & ofv.GenTabla("","aaa","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "GLOSMANAGERdsc.asp?CantRegMover=" & CantRegMoverx & "&Opc=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = ""

  sAccion = "GLOSMANAGER.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Buscar,Celda,Agregar)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

function fechaArg(f)
   dia      = day(f)
   mes      = month(f)
   ano      = year(f)
   fechaarg = dia &"/" & mes & "/" & ano
end function

%>