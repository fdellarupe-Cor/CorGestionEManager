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
  Volver2 = replace(vuelta,"../","")

  set cn = ofv.conectar(ofv.strconn6)
''  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")
  sPERFIL = request.querystring("PERFIL")

  IF SPERFIL <> "" THEN
     SESSION("REQPERFIL") = SPERFIL
  ELSE
     SPERFIL = SESSION("REQPERFIL")
  END IF
 


 if sopcion = "SI" then
    GENTIPO()

 end if


    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"


    celdas = ""

    scomK = "0" & chr(9) & "1"
    scomD = "INCIDENCIAS" & chr(9) & "PROYECTOS/REQUERIMIENTOS"

  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=5","","","","","","TIPOS","si")
  formulario = formulario & ofv.Encabezados("CODIGO","DESCRIPCION","MODO","INFORMACION")


  StrSql = "Select * From TIPOS WHERE TIPOREQ = " & SPERFIL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "ut"


    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=TIPOS&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(1),sHref)
    else
      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

 '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")


       sInput = ofv.GenerarInput("",rs(0),"readonly","","5","dtn","")
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")




       if bModificar then
          sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sform,"90","dt","")
       else
          sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
       end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if bModificar then
        sInput = "<font class=at><select name=" & rs(4).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(4) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >INCIDENCIAS<option value=1 "
        if rs(4) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >PROYECTOS/REQUERIMIENTOS</select></font>"

        sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")

    else
        if rs(4) = 0 then sInput = "INCIDENCIAS"
        if rs(4) = 1 then sInput = "PROYECTOS/REQUERIMIENTOS"
    end if
       celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")


       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";'  href=CORREQTIPOINFO.asp?TIPO=" & rs(0)
       sinput2 = sinput2  & " TITLE='INFORMACION DE " & rs(1) & "'  ><b>" & sImagenC & "</b></a>" 

       celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput2,"si")


    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=TIPOS"
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
    sAccion = "CORREQTIPOPROP.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  


  sAccion = "CORREQTIPO.asp"
  Celda = ofv.BotonVolver(sAccion)


  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GENTIPO()


    strsql = "insert into TIPOS (DESCRI,TIPOREQ,MODO) values ('NUEVO TIPO','" & SPERFIL & "','1')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>