<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	ofv.ObtenerAtributos Session("Form"),""
	bModificar = ofv.AModificar

	CantRegAMover = request.querystring("CantRegAMover")
	nblk = request.querystring("nblk")
	If  Len(nblk) = 0 Then nblk = "D"

	ord  = request.querystring("ord")
	If  Len(ord) = 0 Then ord = "T"

	newdir = request.querystring("newdir") & ""

	If  Len(newdir) = 0 Then newdir = "0"

	opc  = request.querystring("opc")
	cid  = request.querystring("cid")
	xp=request.querystring("xp")


	nCodigo = request.querystring("Codobjeto")
	sTipo   = request.querystring("Tipobjeto")
	sfamily = request.querystring("family")
	idx   = request.querystring("idx")
	vueltaC = request.querystring("volver")
	volver2 = ofv.vueltaasp2("../")

	If IDX <> "" Then
		session("idx")=idx
	Else
		idx=session("idx")
	End If

'response.write "2222222222"
	If session("sfamily")="" Then
		session("sfamily")=sfamily
	Else
		sfamily=session("sfamily")
	End If

	esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

	ccomp  = 0
	tcomp  =  sTipo

	Set cn = ofv.conectar(ofv.strconn2)
	xentro="N"
	Strsql = "Select * From objinfoIMG "
	If opc = "NO" Then
		strsql = strsql & " Where  objid =" & IDX
	End If
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  Not rs.eof Then
		If  opc = "NO" Then
			ccomp = rs(6)
			ccomp2 = rs(6)
			XENTRO="S"
		End If
	End If




	If  nblk = "S" Then ccompS = request.form("comp1")
	If  nblk = "S" Then newdir = ccompS
	If  nblk = "T" Then ccompT = request.form("comp2")
	If  nblk = "T" Then ccomp2 = ccompT

	If  nblk = "D" Then ccomp2 = ccomp

	If  nblk = "D" And ord = "F" Then
		session("comp31") = request.form("comp31")
		session("comp32") = request.form("comp32")
	End If

	If  session("comp31") = "" Then
		session("comp31") = "TODOS"
	End If


	If  session("comp32") = "" Then
		session("comp32") = "TODOS"
	End If


	comp31 = session("comp31")
	comp32 = session("comp32")

	sHTML = ofv.FormHeader(session("FormN"))

	'----------------------------------------------

	shtml = shtml & CorPnlFnc("N","")

	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

	'-------------------------------------------------



	Set fso = CreateObject("Scripting.FileSystemObject")

	sHTML = sHTML & "<center><br>"



	sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","80%","",0,0,0,0,"","","","no")

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=50%  class=tt align=center  >"
	sHTML = sHTML & "Directorios&nbsp;disponibles"
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "<td width=50%  class=tt align=center  >"
	sHTML = sHTML & "Imagenes&nbsp;disponibles"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=50% align=center  >"

	'  response.write direct & " 0<br>"


	sPath = Request.ServerVariables("PATH_INFO")
	'  response.write spath & " 1<br>"
	Xruta = Request.ServerVariables("SERVER_NAME") & Left(sPath,InStrRev(sPath,"/",-1) - 1)
	' response.write Xruta & " 2<br>"
	sPath = server.mappath(sPath)
	'  response.write spath & " 3<br>"
	sPath1 = Replace(sPath,"EdimagenCMPex.asp","",1,-1,1) & "archivos"
	'  response.write spath & " 4<br>"
	'  Set Folder = fso.getfolder(sPath1)
	'  set sfolder = Folder.SubFolders
	'  Set Files = Folder.Files

	If bModificar Then
		sAccion = "EdimagenCMPex.asp?"
		sAccion = sAccion & "nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid & "&newdir=" & newdir
	Else
		sAccion = ""
	End If
	sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action="  & Chr(34) &  sAccion  & Chr(34) &  "  method=post >" & Chr(13)
	sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
	sHTML = sHTML & " onchange=" & Chr(34) & "fcomp1.submit();" & Chr(34)

	sHTML = sHTML & " >"

	sHTML = sHTML & "<option value=" & Chr(34) & "0" & Chr(34)
	If newdir = "0" Then
		sHTML = sHTML & " selected "
	End If
if session("EMPMUN") = "S" then
	sHTML = sHTML & " >Emp" & session("UN") &"/archivos/imagenes</option>"
else
	sHTML = sHTML & " >archivos/imagenes</option>"
end if
	sfolder=Split(getfolders("archivos/IMAGENES",""),"|")
	'  For Each F in sfolder
	For f=0 To UBound(sfolder)-1

		sHTML = sHTML & "<option value=" & Chr(34) & "/" & sfolder(f) & Chr(34)
		If newdir = ("/" & sfolder(f)) Then   sHTML = sHTML & " selected "
		sHTML = sHTML & " >archivos/IMAGENES/" & UCase(sfolder(f)) & "</option>"


	Next




	sHTML = sHTML & "</select></font>"
	sHTML = sHTML & "</FORM>" & Chr(13)

	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td width=50% align=center  >"
	'-------------
	If bModificar Then
		sAccion = "EdimagenCMPex.asp?"
		sAccion = sAccion & "nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid & "&newdir=" & newdir
	Else
		sAccion = ""
	End If

	If newdir = "0" Then
	  if session("EMPMUN") = "S" then
		newdirn = "Emp" & session("UN") &"/archivos/imagenes"
	  else
	  	newdirn = "archivos/imagenes"
	  end if
	Else
	  if session("EMPMUN") = "S" then
		newdirn = "Emp" & session("UN") &"/archivos/imagenes" & newdir
	  else
	    newdirn = "archivos/imagenes" & newdir
	  end if
	End If

	sPath2 = Replace(sPath,"EdimagenCMPex.asp","",1,-1,1) & newdirn
	'  response.write spath2 & " 4<br>"
	'Set Folder = fso.getfolder(sPath2)
	' Set Files = Folder.Files
	'response.write newdirn
	sfilesn=Split(getfilesinfolder(newdirn,""),"|")
	sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action="  & Chr(34) &  sAccion  & Chr(34) &  "  method=post >" & Chr(13)
	sHTML = sHTML & "<font class=at><select size=10  name=comp2 style={border-style:outset;border-width:4;} "
	sHTML = sHTML & " onchange=" & Chr(34) & "fcomp2.submit();" & Chr(34)
	sHTML = sHTML & " >"



	sfcount = 0

	If comp32 = "TODOS" Then
		If comp31 = "TODOS" Then
			Sfiles = "SI"
		Else
			Sfiles = "NO"
		End If
	Else
		Sfiles = "NO"
	End If

	'For Each F in Files
	For f=0 To UBound(sfilesn)-1
		'''  response.write sfilesn(f) & " 1<br>"

		file=Split(sfilesn(f),"###")
		If Sfiles = "NO" Then
			If comp32 = "TODOS" Then
				sfilesva1 = "SI"
			Else
				sfext = InStr(file(0),".")
				If sfext = 0 Then
					sfext = "NO"
				Else
					sfext = UCase(Right(file(0),Len(file(0))-sfext))
				End If
				If sfext = comp32 Then
					sfilesva1 = "SI"
				Else
					sfilesva1 = "NO"
				End If
			End If

			If comp31 = "TODOS" Then
				sfilesva2 = "SI"
			Else
				sfextn = Len(comp31)
				sfext = UCase(Left(file(0),sfextn))
				If sfext = UCase(comp31) Then
					sfilesva2 = "SI"
				Else
					sfilesva2 = "NO"
				End If
			End If

			''''' response.write "sfiles " & sfiles & " comp32 " & comp32 & " comp31 " & comp31 & " va1 " & sfilesva1 & " va2 " & sfilesva2  & " <br>"

			If sfilesva1 = "SI" And sfilesva2 = "SI" Then
				'           response.write " entro " & f.name & " <br>"
				sfilesva = "SI"
			Else
				sfilesva = "NO"
			End If
		Else
			sfilesva = "SI"
		End If



		If sfilesva = "SI" Then
			sfcount = sfcount + 1
			If newdir = "0" Then
				sHTML = sHTML & "<option value=" & Chr(34) & file(0) & Chr(34)
			Else
				sHTML = sHTML & "<option value=" & Chr(34) & Replace(newdir,"/","",1) & "/" & file(0) & Chr(34)
			End If

			If ccomp2 = (newdir & "/" & file(0)) Then   sHTML = sHTML & " selected "
			sHTML = sHTML & " >" & UCase(newdirn & "/" & file(0)) & "</option>"

		End If

	Next

	If sfcount = 0 Then
		sHTML = sHTML & "<option value=" & Chr(34) & ccomp & Chr(34)
		sHTML = sHTML & " selected "
		sHTML = sHTML & " >No hay imagenes disponibles</option>"

	End If


	sHTML = sHTML & "</select></font>"
	sHTML = sHTML & "</FORM>" & Chr(13)
	'--------
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"


	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=50% align=center class=tt   >"
	sHTML = sHTML & "&nbsp;"
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "<td width=50% align=center class=tt   >"
	sHTML = sHTML & "FILTRO: "
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"


	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=50% align=center  >"
	sHTML = sHTML & "&nbsp;"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td width=50% align=center  >"
	'-------------
	'  sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","100%","",0,0,0,0,"","","","no")
	' sHTML = sHTML & "<tr>"

	If bModificar Then
		sAccion = "EdimagenCMPex.asp?"
		sAccion = sAccion & "nblk=D&ord=F&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&cid=" & cid & "&newdir=" & newdir

	Else
		sAccion = ""
	End If
	sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & Chr(34) &  sAccion  & Chr(34) &  "  method=post >" & Chr(13)

	sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","100%","",0,0,0,0,"","","","no")
	sHTML = sHTML & "<tr>"

	sHTML = sHTML & "<td width=35% align=center  >"
	sHTML = sHTML & ofv.GenerarInput("comp31",comp31,"agregar","","80","dt style='display:none'","")
	sHTML = sHTML & "</td>"


	sHTML = sHTML & "<td width=35% align=center  >"

	sHTML = sHTML & "<font class=at><select  name=comp32  "
	' sHTML = sHTML & " onchange=" & chr(34) & "fcomp3.submit();" & chr(34)

	sHTML = sHTML & " >"

	sHTML = sHTML & "<option value=" & comp32
	sHTML = sHTML & " selected "
	sHTML = sHTML & " >" & comp32 & "</option>"

	sHTML = sHTML & "<option value=TODOS "
	sHTML = sHTML & " >TODOS</option>"
	sHTML = sHTML & "<option value=JPG "
	sHTML = sHTML & " >JPG</option>"
	sHTML = sHTML & "<option value=GIF "
	sHTML = sHTML & " >GIF</option>"
	sHTML = sHTML & "<option value=BMP "
	sHTML = sHTML & " >BMP</option>"




	sHTML = sHTML & "</select></font>"
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "<td width=30% align=center  >"
	sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='FILTRAR' class=bt >"
	sHTML = sHTML & "</td>"

	' sHTML = sHTML & "</FORM>" & chr(13)

	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</FORM>" & Chr(13)
	'--------
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"


	sHTML = sHTML & "</table>"

	sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3 style={border-style:solid;border-width:1;border-color:#000000;} ","","100%","",0,0,0,0,"","","","no")
	sHTML = sHTML & "<tr>"

	If opc = "NO" Then
		sHTML = sHTML & "<td valign=top cellpadding=0>"
		sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")



		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  class=tt valign=top>"
		sHTML = sHTML & "Imagen&nbsp;anterior"
		sHTML = sHTML & "</td>"



		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td   class=cc   >"



		ncomp = "" & ccomp

		sHTML = sHTML & "<font color=#000080 ><B>" & ncomp & "</b></font>"
		sHTML = sHTML & "</td>"


		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"

		sImage = ofv.GenImage("Imagen","aaa width=100 border=1","","","" & ccomp,"","")
		sHTML = sHTML & ofv.GenCelda("aaa","","center","","","","",sImage,"si")




		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"

		xcolsp = ""

	Else
		xcolsp = " colspan=2 "

	End If
	sHTML = sHTML & "<td valign=top" & xcolsp & "   >"
	sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td  class=tt align=center>"
	sHTML = sHTML & "Imagen&nbsp;nueva"
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "</tr>"

	'  if nblk = "D" then ccomp2 = ccomp
	'  if nblk = "S" then ccomp2 = ccompS
	'  if nblk = "T" then ccomp2 = ccompT

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td  class=cc  valign=top>"


	If xp="N" Then
		ncomp = ccomp2
	Else

		ncomp = NEWDIRN &"/" & ccomp2
	End If


	sHTML = sHTML & "<font color=#993333 ><B>" & ncomp & "</b></font>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "<tr>"
	sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","" & ncomp,"","")
	sHTML = sHTML & ofv.GenCelda("aaa valign=top","","center","","","","",sImage,"si") & ofv.GenerarInput(rs(6).name,newdirn &"/" & ccomp2,"hidden","","10","","")
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</td>"

	'''''''''''''''''''''



	If bModificar Then

		If opc = "NO" Then
			saccion ="edgeneralinfoimg.asp?"
			sAccion = sAccion & "Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&family=" & sfamily & "&idx=" & idx & "&opcion=UPD"
			'  response.write saccion
		Else

			sAccion = "edgeneralinfoimg.asp?"
			sAccion = sAccion & "Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&family=" & sfamily & "&idx=" & idx & "&opcion=UPD"

		End If


	Else
		sAccion = ""
	End If


	sAccionac = sAccion



	sHTML = sHTML & ofv.GenTabla("panelxx","aaa cellpadding=3  ","","80%","",0,0,0,0,"","","","no")

	sHTML = sHTML & "<tr>"
	'  sImage = ofv.GenerarInput(rs(6).name,newdirn & "/" & ccomp2,"hidden","","10","","")
	sImage = "&nbsp;"
	sHTML = sHTML & ofv.GenCelda("aaa","","center","","","","",sImage,"si")
	sHTML = sHTML & "</tr>"



	'  sInput  = ofv.GenerarInput("","Aceptar","submit","","10","bt","") & ofv.GenerarInput(rs(3).name,ccomp2,"hidden","","10","","")
	'  Celda   = ofv.GenCelda("","","","","","","",sInput,"si")
	'  Agregar = ofv.GenForm("Agregarb","",sAccion,"","post","",Celda,"si")

	'sAccion = "edImagen2.asp?Codcomp=" & nCodigo & "&Opcion=no&CantRegAMover=" & CantRegAMover
	'sAccion = sAccion & "&Tipocomp=" & sTipo

	'  sInput   = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
	'  Celda    = ofv.GenCelda("","","","","","","",sInput,"si")
	'  Cancelar = ofv.GenForm("AgregarC","",sAccion,"","post","",Celda,"si")



	sHTML = sHTML & "</table>"

	sHTML = sHTML & ofv.GenForm("AgregarControlb","",sAccionac,"","post","","","no")

	'Botones Aceptar y Cancelar
	' sAccion = "edgeneralinfoimg.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

	sImage = ofv.GenerarInput(rs(6).name,newdirn & "/" & ccomp2,"hidden","","10","","")

	saccion ="edgeneralinfoimg.asp?"
	sAccion = sAccion & "Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&family=" & sfamily & "&idx=" & idx & "&opcion=NO"

	sHTML = sHTML &   sImage  & ofv.AgregarBoton("AgregarControlb",sAccion,"_self")


	sHTML = sHTML & "</table>"


	sHTML = sHTML & "</center>"

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)


	' Buscar  = Cancelar
	' Celda   = ""

	' sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</div>"

	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	'--------------------------------------------------------------------------------

	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

%>