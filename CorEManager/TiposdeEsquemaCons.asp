<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	ReDim nivel(9)
	nivel(0) = "C"
	nivel(1) = "D"
	nivel(2) = "E"
	nivel(3) = "F"
	nivel(4) = "G"
	nivel(5) = "H"
	nivel(6) = "I"
	nivel(7) = "J"
	nivel(8) = "K"
	nivel(9) = "L"
	sImagenCampo = "ARCHIVO"
	CantRegMostrarx = 8
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	setproc = request.querystring("SETPROC")
	'sHTML = ofv.FormHeader(session("FormN"))
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
	'Formulario = ofv.Encabezados("Tipo","Descripcion","Perfil","Nivel","Icono"," ")
	If session("base")="K" Then
		Formulario=Chr(13)&" <TR ><TD  class=tt width=20% >Descripcion</TD><TD  class=tt width=2% >Nivel</TD></tr>" & Chr(13)
	Else
		Formulario=Chr(13)&" <TR ><TD  class=tt width=20% >Tipo</TD><TD  class=tt width=20% >Descripcion</TD><TD  class=tt width=2% >Perfil</TD><TD  class=tt width=2% >Nivel</TD></tr>" & Chr(13)
	End If
	'  ofv.ObtenerAtributos Session("Form"),""
	'  bAgregar = ofv.AAgregar
	' bModificar = ofv.AModificar
	' bSuprimir = ofv.AEliminar
	' bConsultar = ofv.AConsultar
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	' Strsql = "Select * From Tipobjeto Order By descri"
	If setproc="" Then setproc="ALL"
	If SESSION("BASE")="K" Then
		Strsql = "Select * From Tipobjeto WHERE  TIPOCPT='K' "
	Else
		Strsql = "Select * From Tipobjeto WHERE  TIPOCPT <> 'K' "
	End If
	If setproc <> "ALL" Then
		Strsql = Strsql & " and aplic = '" & setproc & "' "
	Else
		Strsql = Strsql & " and (aplic = '" & wkaplic & "' or aplic = 'GLB') "
	End If
	strsql = strsql & " Order By descri"
	' response.write setproc
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	' if not rs.eof then rs.move(CantRegMoverx)
	posI = 0
	Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)
		sForm = "form" & posi
		
		'Eliminar
		'if bSuprimir then
		'  sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
		'  sHref = sHref & "&Tabla=Tipobjeto&volver=" & vuelta & "&Eliminar=NSNC"
		'  Celda = ofv.BotonEliminar(rs(1),sHref)
		'else
		Celda ="" 'ofv.GenCelda("","","","","","","","&nbsp;","si")
		'end if
		
		'Tipo
		sInput = rs(0)
		Celda = Celda & ofv.gencelda("","vt","","50","","","",sInput,"si")
		
		'Descripcion
		sInput =rs(1) ' ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt","")
		Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")
		
		'Perfil
		' sID = "document." & sForm
		' Sql1 = "Select Tipoperfil From Objperfil Order By Tipoperfil"
		If session("base")="K" Then
		Else
			sInput =rs(3) ' ofv.GenerarCombo(sID,rs(3).name,rs(3).value,Sql1,cn,0,0,"")
			Celda = Celda & ofv.gencelda("","ut ","","30%","","","",sInput,"si")
		End If
		
		'Nivel
		' sInput = "<font class=at><select onchange=document." & sForm & ".submit() name="
		' sInput = sInput & rs(4).name & ">"
		' For i=0 to 9
		'   if nivel(i) = rs(4) then
		'     sInput = sInput & "<option value=" & rs(4) & " selected>" & rs(4)
		'   else
		'     sInput = sInput & "<option value=" & nivel(i) & ">" & nivel(i)
		'   end if
		' next
		sInput = rs(4) 'sInput & "</select></font>"
		Celda = Celda & ofv.gencelda("","ut align=center","","2%","","","",sInput,"si")
		
		'Icono
		'  sImagenValor = rs(15)
		' sImage = ofv.GenImage("Imagen","","","","Iconos/" & sImagenValor,"","")
		' Celda = Celda & ofv.gencelda("","aaa  ","center","","","","",sImage,"si")
		
		
		' sinput = "<a  href=TiposdeEsquemaProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
		' sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & " ><b>Propiedades</b></a>"
		' Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
		
		Fila = ofv.genrow("","","","","","","",Celda,"si")
		
		' sAccion = ""
		' if bModificar then
		'   sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
		'   sAccion = sAccion & "&tabla=Tipobjeto&volver=" & vuelta
		' end if
		Formulario = Formulario & fila 'ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
		
		posi = posi + 1
		rs.Movenext
	Loop
	sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")
	
	'Botonera
	'Agregar = ""
	'if bAgregar then
	'  sAccion = "AgregarTipodeEsquema.asp?Regisvuelta=" & CantRegMoverx
	'  sAccion = sAccion & "&Opcion=si&tabla=Tipobjeto"
	'  Agregar = ofv.BotonAgregar(sAccion)
	'end if
	
	'str1 = "Select descri From Tipobjeto Order By descri "    ' orden pagina
	'str2 = "Select descri From Tipobjeto Order By descri "    ' orden alfabetico
	'Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
	
	'consulta
	'sinput = "<a  href=TiposdeEsquemaCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
	'Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
	
	'sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>