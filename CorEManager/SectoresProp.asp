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

'response.write sql4
  sHTML = ofv.FormHeader(session("FormN"))

     simagen1 = "<font face='corflaticon' size=2 color=#f7aa22>4</font>"

     simagend1 = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"
     simagend2 = "<B><FONT FACE=CorFlaticon COLOR=#336699 SIZE=2>&#0232;</FONT></B>"
     simagend3 = "<B><FONT FACE=CorFlaticon COLOR=#005500 SIZE=2>&#0232;</FONT></B>"

  Formulario = ""


  Strsql = "Select * From ECO where eco = " & wPerfil
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then
  posI = 0

    sForm = "form" & posi

    filas = ""

'descripcion

    sInput = rs(1)

    celdas = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")



    celdas = ofv.GenCelda("","tt  ","","30%","","","","Codigo Interno:","si")


    if bModificar  then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","100","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","vti ","","70%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if session("APLMDCLI")="BC" THEN
     celdas = ofv.GenCelda("","tt  ","","30%","","","","Dep. Contable:","si")

	if bModificar  then
	    ' sInput = ofv.GenerarInput(rs(9).name,rs(9),"text",sForm,"50","dt","")
	      SINPUT="<input type=text size=50 class=dt name="&rs(9).name &" align=left value='" & rs(9) & "' onchange=document."&sForm&".submit();  >"
	   else
	     sInput = ofv.GenerarInput("",rs(9),"readonly","","100","dt","")
	 end if

	 celdas = celdas & ofv.GenCelda("","vti ","","70%","","","",sInput,"si")
     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
    end if



    celdas = ofv.GenCelda("","tt  ","","","","","","Segmento:","si")


    sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,sql1,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Jerarquia:","si")

    sInput = ofv.GenerarCombo(sForm,rs(6).name,rs(6).value,sql2,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Nivel:","si")


    sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,sql3,cn,0,1,"")

    celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","tt  ","","","","","","Site/Sucursal:","si")


    sAccion = "SectorSITES.asp?Regisvuelta=" & CantRegMoverx & "&NOM="
    sAccion = sAccion & ofv.convertircar(rs(1)," ","*") & "&ID=" & rs(0) & "&RegisPerfil=" & wPerfil

      sinput = "<a  href=" & sAccion
      sinput = sinput  & " >" & simagend3 & "</a>"

 '   celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")


    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput   = "<a href='" & sAccion
    sinput   = sinput  & "' title='Site/sucursal de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
    Celdas   = Celdas & ofv.gencelda("","fti","","","","","",sInput,"si")


    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Dependencia:","si")


    sAccion = "SectorDependencias.asp?Regisvuelta=" & CantRegMoverx & "&NOM="
    sAccion = sAccion & ofv.convertircar(rs(1)," ","*") & "&ID=" & rs(0) & "&RegisPerfil=" & wPerfil

      sinput = "<a  href=" & sAccion
      sinput = sinput  & " >" & simagend1 & "</a>"

  '  celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft2o';"" onmouseout=""this.className='ft2n';"" "
    sinput   = "<a href='" & sAccion
    sinput   = sinput  & "' title='Dependencia de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
    Celdas   = Celdas & ofv.gencelda("","fti","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","tt  ","","","","","","Perfiles:","si")


    sAccion = "SectorPerfiles.asp?Regisvuelta=" & CantRegMoverx & "&NOM="
    sAccion = sAccion & ofv.convertircar(rs(1)," ","*") & "&ID=" & rs(0) & "&RegisPerfil=" & wPerfil

      sinput = "<a  href=" & sAccion
      sinput = sinput  & " >" & simagend2 & "</a>"

   ' celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")

    'sImagenP ="&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft3o';"" onmouseout=""this.className='ft3n';"" "
    sinput   = "<a href='" & sAccion
    sinput   = sinput  & "' title='Perfiles de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
    Celdas   = Celdas & ofv.gencelda("","fti","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","tt  ","","","","","","Descripciones:","si")


    sAccion = "SectorDESCRIPCION.asp?Regisvuelta=" & CantRegMoverx & "&NOM="
    sAccion = sAccion & ofv.convertircar(rs(1)," ","*") & "&ID=" & rs(0) & "&RegisPerfil=" & wPerfil

      sinput = "<a  href=" & sAccion
      sinput = sinput  & " >" & simagend2 & "</a>"

   ' celdas = celdas & ofv.GenCelda("","vti ","","","","","",sInput,"si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft3o';"" onmouseout=""this.className='ft3n';"" "
    sinput   = "<a href='" & sAccion
    sinput   = sinput  & "' title='DESCRIPCIONES de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
    Celdas   = Celdas & ofv.gencelda("","fti","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Eco&volver=" & vuelta
    end if
    Formulario = ofv.genform(sForm,"",sAccion,"","post","",FilaS,"si")




    posi = posi + 1

  end if
  sHTML = sHTML & "<center>" & ofv.GenTabla("","aaa","","60%","","0","","-1","2","","",Formulario,"si")




  'Botonera
  Agregar = ""
  Buscar = ""

  sAccion = "SectoresUN.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>