<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sImagenCampo = "ARCHIVO"
  CantRegMostrarx = 100
  'CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  'vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

 'sHTML = ofv.FormHeader("TIPOS DE SUBTEMA")
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
  'Formulario = ofv.Encabezados("Tipo","Descripcion","Nivel","Icono","")
  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Tipo</TD><TD  class=tt width=2% >Descripcion</TD><TD  class=tt width=2% >NIvel</TD><TD  class=tt width=2% >Icono</TD></tr>" & chr(13)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
 'ofv.ObtenerAtributos Session("Form"),""
 ' bAgregar = ofv.AAgregar
 ' bModificar = ofv.AModificar
 ' bSuprimir = ofv.AEliminar
 ' bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

	FILT="C"
	If SESSION ("BASE") ="K" Then
		FILT="K"
	Else
		FILT="C"
	End If

	Strsql = "Select * From Subtema WHERE (aplic = '" & wkaplic & "' or aplic = 'GLB') AND ICONO='"& FILT &"' Order By descri"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
   ' sForm = "form" & posi


      Celda ="" ' ofv.GenCelda("","","","","","","","&nbsp;","si")


    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vt","","30","","","",sInput,"si")

    'Descripcion
    'sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"90","dt","")
    sInput=rs(1)
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    ''Nivel
    sInput = "B"
    Celda = Celda & ofv.gencelda("","ut","center","","","","",sInput,"si")

    ''Icono
    sImagenValor = rs(15)
    sImage = ofv.GenImage("Imagen","","","","Iconos/" & sImagenValor,"","")
    Celda = Celda & ofv.gencelda("","aaa  ","center","","","","",sImage,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si") & chr(13)




    Formulario = Formulario & fila

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","1","","",Formulario,"si")

  'Botonera
  'Agregar = ""
  'if bAgregar then
  '  sAccion = "AgregarTiposdeSubtema.asp?Regisvuelta=" & CantRegMoverx
  '  sAccion = sAccion & "&Opcion=si&tabla=Subtema"
  '  Agregar = ofv.BotonAgregar(sAccion)
  'end if

  str1 = "Select descri From Subtema Order By descri "    ' orden pagina
  str2 = "Select descri From Subtema Order By descri "    ' orden alfabetico
  'Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  '  sinput = "<a  href=TiposdeSubtemaProp.asp?wperfil=" & "&CantRegMover=" & CantRegMoverx
  '  sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & " ><b>consulta</b></a>"
  '  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  'sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>