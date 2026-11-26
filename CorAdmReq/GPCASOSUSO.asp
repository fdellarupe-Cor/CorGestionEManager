<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
''  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")
  XTIPO = REQUEST.QUERYSTRING("TIPO")

  IF XTIPO <> "" THEN
     SESSION("TIPOCU") = XTIPO
  ELSE
     XTIPO = SESSION("TIPOCU")
  END IF

 if sopcion = "SI" then
    GPCASOSUSO()

 end if

     XTIPON = XTIPO


  StrSql = "Select * From GPMODULOS WHERE CODMOD ='" & SESSION("TIPOCU") & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     XTIPOMOD = RS(3)
  END IF

  StrSql = "Select * From GPSISTEMAS WHERE CODSIST ='" & SESSION("TIPO") & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     XTIPOSIS = RS(2)
  END IF


    celdas = ""



  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=3","","","","","","CASOS DE USO PARA EL MODULO " & XTIPOSIS & "/" & XTIPOMOD,"si")
  formulario = formulario & ofv.Encabezados("CODIGO","DESCRIPCION")


  StrSql = "Select * From GPCASOSUSO WHERE CODMOD = '" & SESSION("TIPOCU") & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "uT"


    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=gpmodulos&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(4) & "",sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")


       sInput = rs(0) & ""
       celdas = celdas & ofv.GenCelda("","CC","","5%","","","",sInput,"si")




       if bModificar then
          sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"90","dt","")
       else
          sInput = ofv.GenerarInput("",rs(4),"readonly","","90","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")





    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=GPCASOSUSO"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "GPCASOSUSO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = "" 
 
   sAccion = "GPMODULOS.asp"
  Celda = ofv.BotonVolver(sAccion)

 
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GPCASOSUSO()


    strsql = "insert into GPCASOSUSO (CODUSO,CODSIST,CODMOD) values ('"& GTRnd("USO")&"','" & SESSION("TIPO") & "','" & SESSION("TIPOCU") & "')"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>