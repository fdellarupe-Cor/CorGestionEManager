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


 if sopcion = "SI" then
    GPMODOCONTACTO()

 end if



    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"

    celdas = ""

   ' scomK = "0" & chr(9) & "1" & chr(9) & "2"
    'scomD = "INCIDENCIAS" & chr(9) & "TAREAS DE PROYECTO" & chr(9) & "AUSENCIAS"

  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=5","","","","","","SISTEMAS","si")
  formulario = formulario & ofv.Encabezados("CODIGO","DESCRIPCION" )


  StrSql = "Select *  From GPMODOCONTACTO "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "GPMODC" & rs(0)


       sc = "ut"

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=GPMODOCONTACTO&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(0),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



       sInput = ofv.GenerarInput("",rs(0),"readonly","","5","dtn","")
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")




       if bModificar then
          sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sform,"60","dt","")
       else
          sInput = ofv.GenerarInput("",rs(1),"readonly","","60","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


    







    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=GPMODOCONTACTO"
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
    sAccion = "GPMODOCONTACTO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GPMODOCONTACTO()
 

    strsql = "insert into GPMODOCONTACTO (DESCRI) values (NULL)"
   'CODREQZ = GENREQ(1)
   
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>