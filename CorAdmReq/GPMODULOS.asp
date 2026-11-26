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
     SESSION("TIPO") = XTIPO
  ELSE
     XTIPO = SESSION("TIPO")
  END IF

 if sopcion = "SI" then
    GPMODULOS()

 end if

     XTIPON = XTIPO

  StrSql = "Select * From GPSISTEMAS WHERE CODSIST ='" & XTIPO & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     XTIPON = RS(2)
  END IF


    celdas = ""

    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"

  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=4","","","","","","MODULOS PARA EL SISTEMA " & XTIPON,"si")
  formulario = formulario & ofv.Encabezados("CODIGO","DESCRIPCION","CASOS DE USO")


  StrSql = "Select * From gpmodulos WHERE CODSIST = '" & XTIPO & "'"
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
      Celdas = ofv.BotonEliminar(rs(1),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")


       sInput = rs(0) & ""
       celdas = celdas & ofv.GenCelda("","CC","","5%","","","",sInput,"si")




       if bModificar then
          sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sform,"90","dt","")
       else
          sInput = ofv.GenerarInput("",rs(3),"readonly","","90","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";'  href=GPCASOSUSO.asp?TIPO=" & rs(2)
       sinput2 = sinput2  & " TITLE='CASOS DE USO PARA " & rs(3) & "'  ><b>" & sImagenC & "</b></a>" 

       celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput2,"si")






    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=GPMODULOS"
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
    sAccion = "GPMODULOS.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = "" 
 
   sAccion = "gpsistemas.asp"
  Celda = ofv.BotonVolver(sAccion)

 
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GPMODULOS()


    strsql = "insert into GPMODULOS (CODMOD,CODSIST) values ('"& GTRnd("MOD")&"','" & XTIPO & "')"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>