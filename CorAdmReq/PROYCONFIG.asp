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



   XXXX = OWKF.GETPRYCNFG()





    celdas = ""



  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=4","","","","","","CONFIGURACION","si")



  StrSql = "Select * From PROYCONFIG "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 
  posi = 0

    sform = "REQ" & rs(0)


       sc = "ut"



    celdas = ofv.GenCelda("","TT","","25%","","","","DIAS HABILES/CORRIDOS:","si")

    scomK = "0" & chr(9) & "1"
    scomD = "DIAS CORRIDOS" & chr(9) & "DIAS HABILES"

    if bModificar then
        sInput = "<font class=at><select name=" & rs(2).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(2) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >DIAS CORRIDOS<option value=1 "
        if rs(2) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >DIAS HABILES</select></font>"

        sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")

    else
        if rs(2) = 0 then sInput = "DIAS CORRIDOS"
        if rs(2) = 1 then sInput = "DIAS HABILES"
    end if
       celdas = celdas & ofv.GenCelda("","VT","","70%","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    celdas = ofv.GenCelda("","TT","","","","","","HORAS X DIA:","si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(6).name,rs(6),"text",sform,"5","dtn","")
       else
          sInput = ofv.GenerarInput("",rs(6),"readonly","","5","dtn","")
       end if
       celdas = celdas & ofv.GenCelda("","VT","","","","","",sInput,"si")


    filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

    celdas = ofv.GenCelda("","TT","","","","","","DIAS X SEMANA:","si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(8).name,rs(8),"text",sform,"5","dtn","")
       else
          sInput = ofv.GenerarInput("",rs(8),"readonly","","5","dtn","")
       end if
       celdas = celdas & ofv.GenCelda("","VT","","","","","",sInput,"si")


    filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

    celdas = ofv.GenCelda("","TT","","","","","","MAXIMO HS. IMPUTABLES:","si")


       if bModificar then
          sInput = ofv.GenerarInput(rs(10).name,rs(10),"text",sform,"5","dtn","")
       else
          sInput = ofv.GenerarInput("",rs(10),"readonly","","5","dtn","")
       end if
       celdas = celdas & ofv.GenCelda("","VT","","","","","",sInput,"si")


    filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")


    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=PROYCONFIG"
      formulario = formulario & ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if



  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")

  END IF


  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)

  


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML





%>