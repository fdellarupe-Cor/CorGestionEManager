<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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
    GENESTADO()

 end if



    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>&#0232;</FONT></B>"

    celdas = ""

    scomK = "0" & chr(9) & "1" & chr(9) & "2"
    scomD = "INCIDENCIAS" & chr(9) & "TAREAS DE PROYECTO" & chr(9) & "AUSENCIAS"

  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=5","","","","","","TIPOS DE TAREA","si")
  formulario = formulario & ofv.Encabezados("CODIGO","TIPO","MODO","SUBTIPOS")


  StrSql = "Select * From PROYTIPOTSK "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "ut"

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=PROYTIPOTSK&volver=" & vuelta & "&Eliminar=NSNC"
      Celdas = ofv.BotonEliminar(rs(1),sHref)
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


    if bModificar then
        sInput = "<font class=at><select name=" & rs(2).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(2) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >INCIDENCIAS<option value=1 "
        if rs(2) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >TAREAS DE PROYECTO<option value=2 "
        if rs(2) = 2 then sInput = sInput & "selected"
        sInput = sInput & " >AUSENCIAS</select></font>"

        sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")

    else
        if rs(2) = 0 then sInput = "INCIDENCIAS"
        if rs(2) = 1 then sInput = "TAREAS DE PROYECTO"
        if rs(2) = 2 then sInput = "AUSENCIAS"
    end if



       celdas = celdas & ofv.GenCelda("",sc,"","40%","","","",sInput,"si")


       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";'  href=PROYTSKSUBTIPO.asp?TIPO=" & rs(0)
       sinput2 = sinput2  & " TITLE='SUBTIPOS DE " & rs(1) & "'  ><b>" & sImagenC & "</b></a>" 

       celdas = celdas & ofv.GenCelda("",sc,"CENTER","","","","",sInput2,"si")




    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=PROYTIPOTSK"
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
    sAccion = "PROYTSKTIPO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

FUNCTION GENESTADO()


    strsql = "insert into PROYTIPOTSK (DESCRIPCION,VALORACION) values ('NUEVO TIPO','1')"
    
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION



%>