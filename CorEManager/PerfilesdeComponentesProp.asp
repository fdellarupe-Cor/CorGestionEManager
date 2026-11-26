<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	CantRegMostrarx = request.querystring("CantRegMostrar")
	CantRegMoverx = request.querystring("CantRegMover")


	wperfil = request.querystring("wperfil")
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wperfil=" & wperfil

	sHTML = ofv.FormHeader(session("FormN"))

	'Encabezados
	Formulario = ""


	Dim stabla(12)


	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	Set cn = ofv.conectar(ofv.strconn2)

	posi = 0




	Strsql = "Select * From comperfil where tipoperfil = '" & wperfil & "' "
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then

		fila = ""


		Celda = ofv.GenCelda("","ttt","","15%","","","","Perfil: ","")
		sInput = rs(0) & "&nbsp;-&nbsp;" & rs(1)

		Celda = celda & ofv.GenCelda("","vti","","90%","","","",sInput,"si")

		sinput = "<a  href=PerfilesdeEsquemaNormativoOrd.asp?CantRegMover=" & CantRegMoverx
		sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & "&Wperfil=" & Wperfil  & "&Welem=0 ><b>Orden</b></a>"
		'    Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

		sinput = "<a  href=PerfilesdeEsquemaNormativo.asp?CantRegMover=" & CantRegMoverx
		sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx & " ><b>Volver</b></a>"
		'    Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
		Fila = ofv.genrow("","","","","","","",Celda,"si")



		Formulario = Fila

		sHTML = sHTML & ofv.GenTabla("","aaa","","80% style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","0","","0","1","","",Formulario,"si") & ""



		fila = ""

		Celda = ofv.GenCelda("","tt ","","","","","","solapa","si")
		Celda = celda & ofv.GenCelda("","tt","","","","","","titulo","si")
		Celda = celda & ofv.GenCelda("","tt","","","","","","activa","si")
		Celda = celda & ofv.GenCelda("","tt","","","","","","descripcion","si")
		Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
		Formulario = Fila

		Formulario = ofv.Encabezados("solapa","titulo","activa","descripcion")

		Dim atab(15)

		For xx = 1 To 15
			atab(xx) = 99
		Next

		For xx = 2 To 16


			If CLng( rs(xx + 30) ) < 99 Then

				atab(rs(xx + 30)) = xx


			End If

		Next

		For yy = 1 To 15

			If CLng(atab(yy)) = 99 Then Exit For
			xx = atab(yy)
			fila = ""
			sForm = "form" & posi


			Celda = ""
			sInput = rs(xx + 30)
			Celda = celda & ofv.GenCelda("","CC colspan=2 ","right","1%","","","",sInput,"si")

			If rs(xx)  Then
				clase = "ut style='background-color:color-mix(in srgb, #00bb00, transparent 60%);' "
			Else
				clase = "ut"
			End If


			sInput = ofv.GenerarInput(rs(xx + 15).name,rs(xx + 15),"text",sForm,"60","dt","")
			Celda = celda & ofv.GenCelda("",clase,"","40%","","","",sInput,"si")

			If yy = 1 Then
				sInput = "<b>SI</b>"
			Else
				sInput = ofv.combosino(sForm,rs(xx),rs(xx).name)
			End If




			Celda = celda & ofv.GenCelda("",clase,"center","5%","","","",sInput,"si")



			sInput = "&nbsp;"

			Strsql = "Select descri From comperfildsc where dato = '" & rs(xx).name & "' "
			Set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs2.eof Then
				sInput = rs2(0)
			End If
			Call ofv.cerrarconsulta(rs2)

			Celda = celda & ofv.GenCelda("",clase,"","40%","","","","<i>" & sInput & "</i>","si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

			sAccion = ""
			If bModificar Then
				sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
				sAccion = sAccion & "&tabla=comperfil&volver=" & vuelta
			End If
			Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

			posi = posi + 1



		Next

	End If
	sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","80% ","","0","","0","1","","",Formulario,"si")




	Agregar = ""
	Buscar  = ""

	sAccion = "Perfilesdecomponentes.asp?CantRegMover=" & CantRegMoverx
	Celda   = ofv.BotonVolver(sAccion)

	xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Orden", "_self", "Orden")
	sAccion = "PerfilesdecomponentesOrd.asp?CantRegMover=" & CantRegMoverx & "&Wperfil=" & Wperfil  & "&Welem=0"
	Listar  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")


	sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar & Listar,celda)





	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

%>