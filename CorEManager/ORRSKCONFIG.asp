<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = cdbl(request.querystring("CantRegMostrar"))
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

 
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  set cn = ofv.conectar(ofv.strconn2)
  set cn3 = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From ORRSKCONFIG where APLIC = '" & WKAPLIC & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posI = 0

    sForm = "form" & posi

     Fila = ""


 ' Encabezados 1

    'Descripcion

    sInput = rs(2)
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'AUTORIZACION
     Celda = ofv.gencelda("","tt","","20%","","","","CIRCUITO AUTORIZACION","si")

    scomK = "0" & chr(9) & "1" & chr(9) & "2"
    scomD = "NO APLICABLE" & chr(9) & "OPCIONAL" & chr(9) & "MANDATORIO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(3),"readonly","","10","dt","")
    end if
    Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    'OBSERVACIONES
     Celda = ofv.gencelda("","tt","","20%","","","","CIRCUITO OSERVACIONES","si")

    scomK = "0" & chr(9) & "1" & chr(9) & "2"
    scomD = "NO APLICABLE" & chr(9) & "OPCIONAL" & chr(9) & "MANDATORIO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(4),"readonly","","10","dt","")
    end if
    Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


   if rs(3) > 0 then

    'CIRCUITO
     Celda = ofv.GenCelda("","tt","","","","","","MODELO AUTORIZACION","si")

    SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = '" & WKAPLIC & "'  )  ) "


   ' if bModificar then
   '     StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
   '     sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5),strsql,cn3,0,1,"")

   ' else
        StrSql = "Select * From WKFCKTMASTER WHERE WKFCOD = '" & rs(5) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sniv
   ' end if

    celda = CELDA & ofv.GenCelda("","vti ","","","","","",sInput,"si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   end if

   if rs(4) > 0 then

    'CIRCUITO
     Celda = ofv.GenCelda("","tt","","","","","","MODELO OBSERVACION","si")

    SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = '" & WKAPLIC & "'  )  ) "


  '  if bModificar then
  '      StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
  '      sInput = ofv.GenerarCombo(sForm,rs(6).name,rs(6),strsql,cn3,0,1,"")

  '  else
        StrSql = "Select * From WKFCKTMASTER WHERE WKFCOD = '" & rs(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sniv
   ' end if

    celda = CELDA & ofv.GenCelda("","vti ","","","","","",sInput,"si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


  end if



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORRSKCONFIG&volver=" & vuelta
    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","-1","2","","",Formulario,"si") & "<br>"



 
 end if



  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>