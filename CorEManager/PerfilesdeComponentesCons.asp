<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	
	CantRegMostrarx = 2
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	
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
	'Formulario = ofv.Encabezados("Perfil","Descripcion",4)
	Formulario=Chr(13)&" <TR ><TD  class=tt width=2% >Perfil</TD><TD  class=tt  colspan=5 >Descripcion</TD></tr>" & Chr(13)
	
	'  if request.querystring("seg") <> "" then
	'    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
	'  end if
	
	'ofv.ObtenerAtributos Session("Form"),""
	'bAgregar = ofv.AAgregar
	'bModificar = ofv.AModificar
	'bSuprimir = ofv.AEliminar
	'bConsultar = ofv.AConsultar
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	Strsql = "Select * From ComPerfil Order By descri"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	'if not rs.eof then rs.move(CantRegMoverx)
	posi = 0
	Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)
		'sForm = "form" & posi
		
		'Eliminar
		'if bSuprimir then
		'  sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
		'  sHref = sHref & "&Tabla=Comperfil&volver=" & vuelta & "&Eliminar=NSNC"
		'  Celda = ofv.BotonEliminar(rs(1),sHref)
		'else
		Celda =""' ofv.GenCelda("","","","","","","","&nbsp;","si")
		'end if
		
		'Perfil
		sInput = rs(0)
		Celda = Celda & ofv.gencelda("","vt","","50","","","",sInput,"si")
		
		'Descripcion
		sInput =rs(1) ' ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"120","dt","")
		Celda = Celda & ofv.gencelda("","ut colspan=4","","","","","",sInput,"si")
		
		Celda = Celda & ofv.gencelda("","","","","","","","","si")
		Fila = ofv.genrow("","","","","","","",Celda,"si")
		
		'Perfiles
		Celda = ofv.gencelda("","abc colspan=1","","","","","","","si")
		Celda = Celda & ofv.gencelda("","tu","","2%","","","","Propiedades","si")
		If rs(2) Then 
			sInput = "Si"
		Else
			sInput = "No"    
		End If 'ofv.combosino(sForm,rs(2),rs(2).name)
		Celda = Celda & ofv.gencelda("","ut","","2%","","","",sInput,"si")
		Celda = Celda & ofv.gencelda("","tu","","2%","","","","Estructura","si")
		'sInput =rs(3) ' ofv.combosino(sForm,rs(3),rs(3).name)
		If rs(3) Then 
			sInput = "Si"
		Else
			sInput = "No"    
		End If
		
		Celda = Celda & ofv.gencelda("","ut","","2%","","","",sInput,"si")
		Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
		
		Celda = ofv.gencelda("","abc colspan=1","","","","","","","si")
		Celda = Celda & ofv.gencelda("","tu","","2%","","","","Detalle de la estructura","si")
		'sInput = rs(4) 'ofv.combosino(sForm,rs(4),rs(4).name)
		If rs(4) Then 
			sInput = "Si"
		Else
			sInput = "No"    
		End If    
		Celda = Celda & ofv.gencelda("","ut","","2%","","","",sInput,"si")
		Celda = Celda & ofv.gencelda("","tu","","2%","","","","Imagen de Complemento","si")
		'sInput =rs(5) ' ofv.combosino(sForm,rs(5),rs(5).name)
		
		If rs(5) Then 
			sInput = "Si"
		Else
			sInput = "No"    
		End If
		
		Celda = Celda & ofv.gencelda("","ut","","2%","","","",sInput,"si")
		Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
		
		Celda = ofv.gencelda("","abc colspan=1","","","","","","","si")
		Celda = Celda & ofv.gencelda("","tu","","2%","","","","Relacion Compuesta","si")
		'sInput =rs(6) ' ofv.combosino(sForm,rs(6),rs(6).name)
		If rs(6) Then 
			sInput = "Si"
		Else
			sInput = "No"    
		End If
		
		Celda = Celda & ofv.gencelda("","ut","","2%","","","",sInput,"si")
		Celda = Celda & ofv.gencelda("","abc colspan=2","","2%","","","","&nbsp;","si")
		Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
		
		Celda = ofv.gencelda("","aaa colspan=5","","","","","","&nbsp;","si")
		Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
		' sAccion = ""
		' if bModificar then
		'   sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
		'   sAccion = sAccion & "&tabla=Comperfil&volver=" & vuelta
		' end if
		Formulario = Formulario & fila 'ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
		
		posi = posi + 1
		rs.Movenext
	Loop
	sHTML = sHTML & ofv.GenTabla("","","","80%  ","","0","","0","0","","",Formulario,"si")
	
	'Botonera
	' Agregar = ""
	' if bAgregar then
	'   sAccion = "AgregarPerfilesdeComponentes.asp?Regisvuelta=" & CantRegMoverx
	'   sAccion = sAccion & "&Opcion=si&tabla=Comperfil"
	'   Agregar = ofv.BotonAgregar(sAccion)
	' end if
	
	' str1 = "Select descri From Comperfil Order By descri "    ' orden pagina
	' str2 = "Select descri From Comperfil Order By descri "    ' orden alfabetico
	' Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
	
	'consulta
	' sinput = "<a  href=PerfilesdeComponentesCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
	' Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
	
	' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>