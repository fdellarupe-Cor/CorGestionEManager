<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	Tabla = request.querystring("Tabla")
	vuelta = request.querystring("vuelta")
	Regisvuelta = request.querystring("Regisvuelta")

	Set cn = ofv.conectar(ofv.strconn2)

	strsql = "Select * From " & tabla
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

	sHTML = ofv.FormHeader("AGREGAR PERFILES DE ESQUEMA NORMATIVO")

	sForm = "Agregarb"

	sInput = "AGREGAR " & session("formn")
	Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
	Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
	perfil =  CLng(Date()) &  Time()

	perfil=  retnum (perfil)
	'  response.write perfil
	If SESSION("BASE") ="K" Then
		perfil="FRMPF" & PERFIL


		Celda  = ofv.gencelda("","tt","","20%","","","","Perfil","si")
		' sInputp = ofv.GenerarInput(rs(0).name,"","agregar","","20","dt","")
		sinputpp =ofv.GenerarInput(rs(0).name,perfil,"HIDDEN","","120","dt","")
		'' Celda  = Celda & ofv.gencelda("","vti","","80%","","","",perfil,"si")
		'Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	Else
		Celda  = ofv.gencelda("","tt","","20%","","","","Perfil","si")
		' sInputp = ofv.GenerarInput(rs(0).name,"","agregar","","20","dt","")
		sinputP =ofv.GenerarInput(rs(0).name,perfil,"HIDDEN","","120","dt","")
		Celda  = Celda & ofv.gencelda("","vti","","80%","","","",perfil,"si")
		Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")
	End If

	Celda  = ofv.gencelda("","tt","","","","","","Descripcion","si")
	sInput = ofv.GenerarInput(rs(1).name,"","agregar","","90","dt","")
	Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
	Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

	'Perfiles

	sInput1 = sinputP & sinputpp & ofv.GenerarInput(rs(2).name,"1","hidden","","1","dt","")
	sInput2 = ofv.GenerarInput(rs(4).name,"1","hidden","","1","dt","")
	sInput3 = ofv.GenerarInput(rs(7).name,"1","hidden","","1","dt","")
	sInput4 = ofv.GenerarInput(rs(16).name,"1","hidden","","1","dt","")
	sInput5 = ofv.GenerarInput(rs(32).name,"9999","hidden","","1","dt","")
	sInput5 = sInput5 & ofv.GenerarInput(rs(3).name,"0","hidden","","1","dt","")
	sInput5 = sInput5 & ofv.GenerarInput(rs(5).name,"0","hidden","","1","dt","")
	sInput5 = sInput5 & ofv.GenerarInput(rs(6).name,"0","hidden","","1","dt","")

	For ii = 8 To 15
		If SESSION("BASE") ="K"  And  ii=11 Then
			sInput5 = sInput5 & ofv.GenerarInput(rs(ii).name,"1","hidden","","1","dt","")
		Else
			sInput5 = sInput5 & ofv.GenerarInput(rs(ii).name,"0","hidden","","1","dt","")
		End If

	Next

	For ii = 17 To 31
		sInput5 = sInput5 & ofv.GenerarInput(rs(ii).name,"SD","hidden","","1","dt","")

	Next

	For ii = 33 To 46

		sInput5 = sInput5 & ofv.GenerarInput(rs(ii).name,"99","hidden","","1","dt","")
	Next
	Celda = ofv.gencelda("","vt colspan=2 ","","","","","",sInput1 & sInput2 & sInput3 & sInput4 & sInput5,"si")
	Fila = Fila & ofv.genrow("","","","","","","",celda,"si")



	sFuncion =  Chr(34) & " var fr=document.Agregarb; if (fr." & rs(0).name
	sFuncion = sFuncion & ".value.length==0 || fr." & rs(1).name & ".value.length==0 "
	sFuncion = sFuncion & "|| fr." & rs(0).name & ".value==' ') "
	sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;} else {document.Agregarb.submit();}" & Chr(34)

	'sFuncion = "onsubmit=" & chr(34) & "if (this." & rs(0).name
	'sFuncion = sFuncion & ".value.length==0 || this." & rs(1).name & ".value.length==0 "
	'sFuncion = sFuncion & "|| this." & rs(0).name & ".value==' ') "
	'sFuncion = sFuncion & "{alert('Los campos no deben ser nulos');return false;}" & chr(34)
	sAccion = "asp/Agregar.asp?Tabla=" & Tabla
	sAccion = sAccion & "&volver=../PerfilesdeEsquemaNormativo.asp?CantRegMover=" & Regisvuelta   &  "_perfil=" & perfil
	Formulario = Formulario & ofv.genform("Agregarb","",sAccion,"","post","",Fila,"no")

	sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

	'Aceptar y Cancelar
	sAccion = "PerfilesdeEsquemaNormativo.asp?CantRegMover=" & Regisvuelta  &  "&perfil=" & perfil
	zboton=ofv.AgregarBoton("Agregarb",sAccion,"") & Chr(13)
	zboton=Replace(zboton,"'document.Agregarb.submit();'",sFuncion)
	sHTML = sHTML & zboton
	'sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")


	Agregar = ""
	Buscar  = ""
	Celda   = ""

	sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

%>