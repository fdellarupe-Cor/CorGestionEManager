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
  wPerfil = request.querystring("wPerfil")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wPerfil=" & wPerfil

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sql1 = "Select * From SEGMENTO Order By descri"
  sql2 = "Select * From JERARQUIA Order By descri"
  sql3 = "Select * From NIVELJERAR Order By descri"
  sql4 = "Select * From sites where siteun = " & session("sun") & " Order By descri"


  sHTML = ofv.FormHeader(session("FormN"))

     simagen1 = "<font face='corflaticon' size=2 color=#f7aa22>4</font>"

     simagend1 = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;/FONT></B>"
     simagend2 = "<B><FONT FACE=CorFlaticon COLOR=#336699 SIZE=2>&#0232;</FONT></B>"

  Formulario = ""


  Strsql = "Select * From ECO where eco = " & wPerfil
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then
  posI = 0

    sForm = "form" & posi

    filas = ""

'descripcion

    celdas = ofv.GenCelda("","tt  ","","20%","","","","Nombre:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sInput = "<big>" & rs(1) & "</big>"

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","tt  ","","","","","","Codigo Interno:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    if bModificar  then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","100","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","tt  ","","","","","","Segmento:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,sql1,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Jerarquia:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sInput = ofv.GenerarCombo(sForm,rs(6).name,rs(6).value,sql2,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Nivel:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,sql3,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","tt  ","","","","","","Site/Sucursal:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sInput = ofv.GenerarCombo(sForm,rs(8).name,rs(8).value,sql4,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Dependencia:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sAccion = "SectorDependencias.asp?Regisvuelta=" & CantRegMoverx & "&NOM="
    sAccion = sAccion & ofv.convertircar(rs(1)," ","*") & "&ID=" & rs(0) & "&RegisPerfil=" & wPerfil

      sinput = "<a  href=" & sAccion
      sinput = sinput  & " >" & simagend1 & "</a>"

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Perfiles:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",simagen1,"si")

    sAccion = "SectorPerfiles.asp?Regisvuelta=" & CantRegMoverx & "&NOM="
    sAccion = sAccion & ofv.convertircar(rs(1)," ","*") & "&ID=" & rs(0) & "&RegisPerfil=" & wPerfil

      sinput = "<a  href=" & sAccion
      sinput = sinput  & " >" & simagend2 & "</a>"

    celdas = celdas & ofv.GenCelda("","vt ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Eco&volver=" & vuelta
    end if
    Formulario = ofv.genform(sForm,"",sAccion,"","post","",FilaS,"si")




    posi = posi + 1

  end if
  sHTML = sHTML & "<br><br><br><dir><dir>" & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f7aa22;' ","","80%","","0","","-1","2","","",Formulario,"si")


      sinput = "<a  href=SectoresUN.asp?CantRegMover=" & CantRegMoverx
      sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & " ><b>Volver</b></a>"
      Celda =  ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","","","",sInput,"si")

      Fila = ofv.genrow("","","","","","","",Celda,"si")

  sHTML = sHTML & "<br><br>" & ofv.GenTabla("","","","15%","","0","","0","0","","",Fila,"si")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>