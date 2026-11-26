	<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")
	sID = request.querystring("ID")
	sIDx = request.querystring("IDx")
	sIDt = request.querystring("IDt")
	Tabla = request.querystring("Tabla")
	Volver = request.querystring("Volver")
	
	sxopc1 = ""
	
	If sidt = "X" Then
		sxopc1 = sxopc1 & "<option value=A>Titulo</option> "
	End If
	
	If sidt = "A" Then
		sxopc1 = sxopc1 & "<option value=B>Subtitulo&nbsp;1</option> "
	End If 
	
	If sidt = "B" Then
		sxopc1 = sxopc1 & "<option value=C>Subtitulo&nbsp;2</option> "
	End If 
	If sidt = "C" Then
		sxopc1 = sxopc1 & "<option value=D>Subtitulo&nbsp;3</option> "
	End If 
	If sidt = "D" Then
		sxopc1 = sxopc1 & "<option value=E>Subtitulo&nbsp;4</option> "
	End If 
	If sidt = "E" Then
		sxopc1 = sxopc1 & "<option value=F>Subtitulo&nbsp;5</option> "
	End If 
	
	If sidt < "G" Then
		sxopc1 = sxopc1 & "<option value=G>Texto</option> "
		sxopc1 = sxopc1 & "<option value=H>Imagen</option> "
		'     sxopc1 = sxopc1 & "<option value=I>Tabla</option> "
		If session("base")="K" Then
			sxopc1 = sxopc1 & "<option value=J>Formulario</option> "
		Else
			sxopc1 = sxopc1 & "<option value=J>Componente</option> "
		End If 
	End If
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	sIDw = sIDx
	
	
	strsql = "Select * From " & tabla & " Where detdepen = " & sIDw
	strsql = strsql & " and detcon =" & sID
	strsql = strsql & " Order By detsecuen desc"
	
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then
		nSecuencia = CInt(rs("detsecuen")) + 10
		nTipo = "2"
	Else
		nSecuencia = 10
		nTipo = "1"
	End If
	
	sHTML = ofv.FormHeader(session("FormN"))
	
	'----------------------------------------------
	
	shtml = shtml & CorPnlFnc("N","")
	
	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
	
	'-------------------------------------------------
	
	f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
	f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
	f1a = f1a & "&Tipobjeto=" & sTipo 
	sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")
	
	sHTML = sHTML & ofv.GenTabla("panel","aaa","","80%","",0,0,0,0,"","","","no")
	
	'Encabezados Descripcion y Secuencia
	h1 = ofv.GenCelda("","tt","","2%","","","","Orden","si")
	h1 = h1 & ofv.GenCelda("","tt","","2%","","","","Descripcion","si")
	h1 = h1 & ofv.GenCelda("","tt","","14%","","","","Tipo","si")
	sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
	
	'Texto Secuencia
	t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(4).name,nSecuencia,"Agregar","","10","dtn",""),"si")
	'Texto Descripcion
	t2 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(7).name,"Descripcion","Agregar","","100","dt",""),"si")
	
	t21 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(5).name & " >" & sxopc1 & "</select>","si")
	
	t3 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(3).name,sID,"hidden","","","",""),"si")
	
	
	t4 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(6).name,sIDx,"hidden","","","","") & ofv.GenerarInput(rs(8).name," ","hidden","","","",""),"si")
	
	t4 = t4 & ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(10).name,GTkey(nCodigo),"hidden","","","",""),"si")
	
	
	
	sInput1 = ofv.GenerarInput("Modi","1","hidden","","10","","")
	sInput2 = ofv.GenerarInput("TipModi","A","hidden","","10","","")
	
	t5 = ofv.GenCelda("","","","","","","",sInput1 & sInput2,"si")
	
	sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2 & t21 & t3 & t4 & t5,"si")
	sHTML = sHTML & ofv.GenFTabla()
	'Botones Aceptar y Cancelar
	' sAccion = "edestrdet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & sid
	
	saccion = Replace(volver,"_","&")
	saccion = Replace(saccion,"../","") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")
	
	
	
	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</span></div>"
	
	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
	
	'Estructura del Indice
	
	sHTML = sHTML & CorPnlIndx("N","")
	
	'--------------------------------------------------------------------------------
	
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	
	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

%>