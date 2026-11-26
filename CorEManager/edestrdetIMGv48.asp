<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False

	IDTXT = request.querystring("IDTXT")
	vuelvo = request.querystring("vuelve")
	vuelvo = Replace(vuelvo,"codobjeto","Codobjeto")
	vuelve = Replace(vuelvo,"***","_")
	volver = "_Vuelve=" & Replace(vuelve,"_","***")
	addqry = Replace(vuelve,"_","&")

	nblk = request.querystring("nblk")
	If Len(nblk) = 0 Then nblk = "D"

	ord = request.querystring("ord")
	If Len(ord) = 0 Then ord = "T"

    nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")

	if session("EMPMUN") = "S" then
    Draiz="Emp" & SESSION("D" & nCodigo & "-0" ) & "/ESTRUCTURA"
    else
     Draiz= "ESTRUCTURA"
    end if


	opc = request.querystring("opc")


	'response.write request.form




	esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

	ccomp = 0
	tcomp =  sTipo

	Set cn = ofv.conectar(ofv.strconn2)
	opc = "NO"

	strsql = "SELECT * from Objestdet where objid = " & IDTXT
	response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		If opc = "NO" Then
			original=rs(9)
			ccomp = rs(9)&""
			'response.write ccomp & " -----"
			ccomp2 = rs(9)
			IF CCOMP="" THEN
			XDIRP="/"
			ELSE
			xdirp= Left(ccomp,InStrRev(ccomp,"/",-1) - 1)
			'response.write  ccomp & "<br>"
			END IF
		End If
	ELSE

	End If
'RESPONSE.WRITE xdirp

	'  RESPONSE.WRITE "ccomp " & ccomp & " ccomp2 " & ccomp2 & " <br>"

	If nblk = "S" Then ccompS = request.form("comp1")
	If nblk = "S" Then newdir = ccompS
	If nblk = "T" Then ccompT = request.form("comp2")
	If nblk = "T" Then ccomp2 = ccompT

	If nblk = "D" Then ccomp2 = ccomp

	If nblk = "D" And ord = "F" Then
		session("comp31") = request.form("comp31")
		session("comp32") = request.form("comp32")
	End If

	If session("comp31") = "" Then
		session("comp31") = "TODOS"
	End If


	If session("comp32") = "" Then
		session("comp32") = "JPG"
	End If


	comp31 = session("comp31")
	comp32 = session("comp32")

   xdir=request.form("comp1")
RESPONSE.WRITE "<br>" & XDIR & "<br>"
'   if xdir <> "" then
     IF XDIR <> "RAIZ" AND XDIR <> "" THEN
     xdirp=xdir
     ELSE
     XDIRP="/"
     END IF
  ' end if

RESPONSE.WRITE XDIRp
	Set fso = CreateObject("Scripting.FileSystemObject")

	'''' sHTML = ofv.menuheader("","")


	sHTML = ofv.FormHeader(session("FormN"))

	'----------------------------------------------

	shtml = shtml & CorPnlFnc("N","")

	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

	'-------------------------------------------------



	sHTML = sHTML & "<center><br>"

	sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","80%","",0,0,0,0,"","","","no")

	sicK = "<font face='corflaticon' size=4 color=#993333 >&#0163;4</font>&nbsp;"
	Celda =  sicK & "<b>Modificacion de Imagen de contenido</b>"

	'  sHTML = sHTML & "<tr>"
	'  sHTML = sHTML & "<td width=50%  colspan=2 class=tt   >"
	'  sHTML = sHTML & Celda
	'  sHTML = sHTML & "</td>"
	'  sHTML = sHTML & "</tr>"

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=50%  class=ttt align=center  >"
	sHTML = sHTML & "Directorios&nbsp;disponibles"
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "<td width=50%  class=ttt align=center  >"
	sHTML = sHTML & "Imagenes&nbsp;disponibles"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td width=50% align=center >"




	sPath = Request.ServerVariables("PATH_INFO")

	Xruta = Request.ServerVariables("SERVER_NAME") & Left(sPath,InStrRev(sPath,"/",-1) - 1)
	' response.write Xruta & " 2<br>"
	sPath = server.mappath(sPath)
	'  response.write spath & " 3<br>"
	sPath1 = Replace(sPath,"edestrdetIMG.asp","",1,-1,1) & "ESTRUCTURA"
	'  response.write spath & " 4<br>"
	'  Set Folder = fso.getfolder(sPath1)
	' set sfolder = Folder.SubFolders
	'  Set Files = Folder.Files

	If bModificar Then
		sAccion = "edestrdetIMG.asp?"
		sAccion = sAccion & "nblk=S&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&IDTXT=" & IDTXT & "&newdir=" & newdir & "&vuelve=" & vuelvo
	Else
		sAccion = ""
	End If
	sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action="  & Chr(34) &  sAccion  & Chr(34) &  "  method=post >" & Chr(13)
	sHTML = sHTML & "<font class=at><select size=10  name=comp1 style={border-style:outset;border-width:4;} "
	sHTML = sHTML & " onchange=" & Chr(34) & "fcomp1.submit();" & Chr(34)

	sHTML = sHTML & " >"

    if Left(xdirp,InStrRev(xdirp,"/",-1) - 1)= ""  then

     sHTML = sHTML & "<option STYLE='background-color:#F7AA22;font-weight:bold;' value=" & Chr(34) & "RAIZ" & Chr(34)
     sHTML = sHTML & "  "
	 sHTML = sHTML & " >" & draiz & XDIRP &"</option>"
    else
	 sHTML = sHTML & "<option STYLE='background-color:#F7AA22;font-weight:bold;' value=" & Chr(34) & Left(xdirp,InStrRev(xdirp,"/",-1) - 1) & Chr(34)
	 sHTML = sHTML & "  "
	 sHTML = sHTML & " >" & draiz & xdirp & "</option>"
    end if



	If session("EMPMUN") = "S" Then

  '   RESPONSE.WRITE xdirp

		sfolder=Split(getfolders(draiz & xdirp,""),"|")

	'	ruta=server.mappath(request.servervariables("script_name"))

		'pos=InStrRev(ruta,"\")
		'ruta=Mid(ruta,1,pos)
	'	response.write ruta
	Else
		sHTML = sHTML & "ESTRUCTURA</option>"
		sfolder=Split(getfolders("ESTRUCTURA",""),"|")
	End If

'response.write "0000" & xdirp & "000"

	'  For Each F in sfolder
	For f=0 To UBound(sfolder)-1
	    IF XDIRP="/" THEN XDIRP=""
		sHTML = sHTML & "<option value=" & Chr(34) & wxdirp &"/" & sfolder(f) & Chr(34)
		If newdir = ("/" & sfolder(f)) Then   sHTML = sHTML & " selected "
		If session("EMPMUN") = "S" Then
			sHTML = sHTML & " >" & draiz & xdirp &"/" & UCase(sfolder(f)) & "</option>"
		Else
			sHTML = sHTML & " >ESTRUCTURA/" & UCase(sfolder(f)) & "</option>"
		End If





	Next




	sHTML = sHTML & "</select></font>"
	sHTML = sHTML & "</FORM>" & Chr(13)

	sHTML = sHTML & "</td>"
	sHTML = sHTML & "<td width=50% align=center  >"
	'-------------
	If bModificar Then
		sAccion = "edestrdetIMG.asp?"
		sAccion = sAccion & "nblk=T&ord=" & ord & "&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&IDTXT=" & IDTXT & "&newdir=" & newdir & "&vuelve=" & vuelvo
	Else
		sAccion = ""
	End If

	If newdir = "0" Then
		If session("EMPMUN") = "S" Then
			newdirn = "Emp" & SESSION("D" & nCodigo & "-0" ) & "/ESTRUCTURA"
		Else
			newdirn = "ESTRUCTURA"
		End If
	Else
		If session("EMPMUN") = "S" Then
			newdirn = newdir ' "Emp" & SESSION("D" & nCodigo & "-0" ) & "/ESTRUCTURA" & newdir
		Else
			newdirn = "ESTRUCTURA" & newdir
		End If
	End If




	sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action="  & Chr(34) &  sAccion  & Chr(34) &  "  method=post >" & Chr(13)
	sHTML = sHTML & "<font class=at><select size=10  name=comp2  "
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


	sfilesn=Split(getfilesinfolder(DRAIZ & xdirp,""),"|")
	'For Each F in Files
	For f=0 To UBound(sfilesn)-1

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

			'      response.write "sfiles " & sfiles & " comp32 " & comp32 & " comp31 " & comp31 & " va1 " & sfilesva1 & " va2 " & sfilesva2 & " name " & f.name & " <br>"

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
				sHTML = sHTML & " title=" & Chr(34) & file(0) & Chr(34)
			Else
				' response.write replace(newdir,"/","",1) & "/" & file(0)
				sHTML = sHTML & "<option value=" & Chr(34) & file(0) & Chr(34)
				sHTML = sHTML & " title=" & Chr(34) & Replace(newdir,"","",1) & "/" & file(0) & Chr(34)

			End If

			If ccomp2 = (newdir & "/" & file(0)) Then   sHTML = sHTML & " selected "
			sHTML = sHTML & " >" & UCase( file(0)) & "</option>"

		End If

	Next

	If sfcount = 0 Then
		sHTML = sHTML & "<option value=" & Chr(34) & ccomp & Chr(34)
		sHTML = sHTML & " selected "
		sHTML = sHTML & " >No hay imagenes disponibles</option>"

	End If


	sHTML = sHTML & "</select></font>"
	sHTML = sHTML & "<INPUT TYPE=HIDDEN NAME=comp1 VALUE=" & CHR(34)& XDIRP & CHR(34) & "></INPUT>"
	sHTML = sHTML & "</FORM>" & Chr(13)
	'--------
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"


	sHTML = sHTML & "<tr>"

	sHTML = sHTML & "<td colspan=2 align=center  >"
	'-------------
	If bModificar Then
		sAccion = "edestrdetIMG.asp?"
		sAccion = sAccion & "nblk=D&ord=F&Codobjeto=" & nCodigo & "&CantRegAMover=" & CantRegAMover
		sAccion = sAccion & "&Tipobjeto=" & sTipo & "&opc=" & opc & "&IDTXT=" & IDTXT & "&newdir=" & newdir & "&vuelve=" & vuelvo

	Else
		sAccion = ""
	End If
	sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action="  & Chr(34) &  sAccion  & Chr(34) &  "  method=post >" & Chr(13)

	sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")
	sHTML = sHTML & "<tr>"

	' sHTML = sHTML & "<td width=15% class=ttt   >"
	' sHTML = sHTML & "FILTRO: "
	' sHTML = sHTML & "</td>"


	sHTML = sHTML & "<td width=35% align=center  >"
	sHTML = sHTML & ofv.GenerarInput("comp31",comp31,"agregar","","80","dt style='display:none'","")
	sHTML = sHTML & "</td>"


	sHTML = sHTML & "<td width=35% align=center  >"

	sHTML = sHTML & "<font class=at><select  name=comp32  "
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

	sHTML = sHTML & "<td width=30% align=center >"
	sHTML = sHTML & "<INPUT TYPE=HIDDEN NAME=comp1 VALUE=" & CHR(34)& XDIRP & CHR(34) & "></INPUT>"
	sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='FILTRAR' class=bt >"
	sHTML = sHTML & "</td>"



	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</FORM>" & Chr(13)
	'--------



	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"


	sHTML = sHTML & "</table>"

	sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=2 style='border-style:solid;border-width:1;border-color:" & session("cmfmenu") & ";' ","","80%","",0,0,0,0,"","","","no")
	sHTML = sHTML & "<tr>"

	opc = "NO"
	If opc = "NO" Then
		sHTML = sHTML & "<td align=center valign=top style='border-style:solid;border-width:1;border-color:#dddddd;' >"
		sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")



		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  class=ttt align=center  >"
		sHTML = sHTML & "Documento&nbsp;anterior"
		sHTML = sHTML & "</td>"



		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td   class=ut   >"



		ncomp = newdirn & "/" & ccomp
        if session("EMPMUN") = "S" then
		 sHTML = sHTML & "<font color=#000080 ><B>Emp" & SESSION("D" & nCodigo & "-0" )  & "/ESTRUCTURA" & original & "</b></font>"
		else
		sHTML = sHTML & "<font color=#000080 ><B>ESTRUCTURA" & original & "</b></font>"
		end if

		sHTML = sHTML & "</td>"


		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		If session("EMPMUN") = "S" Then
			sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","Emp" & SESSION("D" & nCodigo & "-0" )  & "/ESTRUCTURA/" & ccomp,"","")
		Else
			sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","ESTRUCTURA/" & ccomp,"","")
		End If
		sHTML = sHTML & ofv.GenCelda("aaa ","","center","","","","",sImage,"si")
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</td>"

		xcolsp = ""

	Else
		xcolsp = " colspan=2 "

	End If
	sHTML = sHTML & "<td " & xcolsp & " align=center valign=top style='border-style:solid;border-width:1;border-color:#dddddd;' >"
	sHTML = sHTML & ofv.GenTabla("panel2","aaa cellpadding=3  ","","100%","",0,0,0,0,"","","","no")

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td  class=ttt align=center  >"
	sHTML = sHTML & "Documento&nbsp;nuevo"
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "</tr>"

	'  if nblk = "D" then ccomp2 = ccomp
	'  if nblk = "S" then ccomp2 = ccompS
	'  if nblk = "T" then ccomp2 = ccompT

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td  class=ut   >"


	ncomp = newdirn & "/" & ccomp2

	sHTML = sHTML & "<font color=#993333 ><B>" & replace(Draiz & XDIRP  & "/" &ccomp2,"//","/") & "</b></font>"
	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "<tr>"

	' response.write request.form & "<br>"
	' response.write XDIRP
	If session("EMPMUN") = "S" Then
		sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","", replace(Draiz & XDIRP  & "/" &ccomp2,"//","/"),"","")
	Else
	'	sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","","ESTRUCTURA/" & ccomp2,"","")
        sImage = ofv.GenImage("Imagen","aaa width=100 border=1 ","","",replace(Draiz & XDIRP  & "/" &ccomp2,"//","/"),"","")
	End If
	sHTML = sHTML & ofv.GenCelda("aaa","","center","","","","",sImage,"si")
	sHTML = sHTML & "</tr>"

	sHTML = sHTML & "</table>"
	sHTML = sHTML & "</td>"

	'''''''''''''''''''''

	sHTML = sHTML & "<tr>"
	sHTML = sHTML & "<td align=center  >"
	If bModificar Then

		sAccion = Chr(34) & "asp/ActualizarEN.asp?ID=" & idtxt & "&Tabla=Objestdet"
		sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&volver=../edestrdet.asp?" & vuelvo & Chr(34)



	Else
		sAccion = ""
	End If

'	RESPONSE.WRITE  ccomp2
	sHTML = sHTML & "<FORM  name=acep ID=acep action=" & sAccion & "  method=post >" & Chr(13)
	sHTML = sHTML & ofv.GenerarInput(rs(9).name,XDIRP  & "/"  & ccomp2,"hidden","","10","","")
	sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='ACEPTAR' class=bt >"
	sHTML = sHTML & "</FORM>" & Chr(13)
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "<td  align=center   >"

	sAccion = "edestrdet.asp?" & addqry
	sHTML = sHTML & "<FORM  name=Cancel ID=Cancel action=" & sAccion & "  method=post >" & Chr(13)
	sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='CANCELAR' class=bt >"
	sHTML = sHTML & "</FORM>" & Chr(13)
	sHTML = sHTML & "</td>"

	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & "</center>"


	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</div>"

	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	'--------------------------------------------------------------------------------

	sHTML = sHTML & "</body></html>"

End If

response.write sHTML

%>