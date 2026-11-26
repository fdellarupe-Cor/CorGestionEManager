<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

 ' sHTML = ofv.FormHeader(session("FormN"))
  response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
  response.write "<tr>"
  response.write "<td width=100% >"
  response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
  response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
  response.write "<tr>"
  response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
  response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
  response.write "</tr></table>"
 

  'Encabezados
  'Formulario = ofv.Encabezados("Nivel","Descripcion")
  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Nivel</TD><TD  class=tt width=30% >Descripcion</TD></tr>" & chr(13)

 ' ofv.ObtenerAtributos Session("Form"),""
 ' bAgregar = ofv.AAgregar
 ' bModificar = ofv.AModificar
 ' bSuprimir = ofv.AEliminar
 ' bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From NivelJerar Order By codnivel"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 ' if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)
    'sForm = "form" & posi

    'Eliminar
    'if bSuprimir then
    '  sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
    '  sHref = sHref & "&Tabla=Niveljerar&volver=" & vuelta & "&Eliminar=NSNC"
    '  Celda = ofv.BotonEliminar(rs(1),sHref)
    'else
      Celda = "" 'ofv.GenCelda("","","","","","","","&nbsp;","si")
    'end if

    'Nivel
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","vt","right","50","","","",sInput,"si")

    'Descripcion
    sInput =rs(1) ' ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")

    'Celda = Celda & ofv.gencelda("","","","","","","","","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    'sAccion = ""
    'if bModificar then
    '  sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
    '  sAccion = sAccion & "&tabla=Niveljerar&volver=" & vuelta
    'end if
    Formulario = Formulario & fila 'ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
 ' Agregar = ""
 ' if bAgregar then
 '   sAccion = "AgregarNiveles.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=Niveljerar"
 '   Agregar = ofv.BotonAgregar(sAccion)
 ' end if
'
 ' str1 = "Select descri From Niveljerar Order By codnivel "    ' orden pagina
 ' str2 = "Select descri From Niveljerar Order By descri "    ' orden alfabetico
 ' Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
 ' 
  'consulta
 ' sinput = "<a  href=NivelesCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
 ' Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
  
 ' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>