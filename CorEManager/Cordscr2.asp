<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->


<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	shtml = ofv.formheader(session("FormN"))

	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	direct = request.querystring("direct")
	nomdir = request.querystring("nomdir")
	rutdir = request.querystring("rutdir")
	OPC = request.querystring("OPC")

	'response.write request.querystring

	If opc = "" Then opc = "NO"
	If newdir = "" Then newdir = "NO"


	If OPC="DEL" Then
		arch=request.FORM("ARCH")
		delfile(rutdir  & "\" & arch)
		'response.write rutdir  & "\" & arch
	End If

	sHTML = sHTML & ofv.GenTabla("","aaa","","650 ","","0","0","0","0","","","","no")
	sHTML = sHTML & "<font color=#DCDCDC face=tahoma size=1><b>"
	h1 = ofv.GenCelda("","ttt colspan=4","","","","","",rutdir,"si")
	sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
	h0 = ofv.GenCelda("","ttr","","","","","","*","si")
	h1 = ofv.GenCelda("","ttr","","","","","","Nombre","si")
	h2 = ofv.GenCelda("","ttr","","","","","","Fecha","si")
	h3 = ofv.GenCelda("","ttr","","","","","","Subir","si")
	sHTML = sHTML & ofv.GenRow("","","","","22","","",H0 & h1 & H2 & H3,"si")
	sHTML = sHTML & "</b></font>"

  ' response.write request.form("newdir")
	If  opc = "SI" Then
		newdir = request.form("newdir")
		If  newdir = "" Then newdir = "NO"
		If  newdir <> "NO" Then
			'sfolder.add(newdir)
			' RESPONSE.WRITE rutdir & newdir
			If SESSION("FILETIPO")  = "L" Then
				Call getdata( "","AF",RUTDIR,newdir)
			Else
				Call ADDFOLDER(rutdir,newdir,"")
			End If
		End If
	End If


	xTrColor = "#F9F9F9"

	If Trim(UCase(rutdir))<> Trim(UCase(direct)) Then

		celda=  ofv.GenCelda("","ut","","1%","","","","","si")
		saccion   = "<i class='las la-folder' style=font-size:12pt;color:" & Session("CMFPNTIT") & "></i>&nbsp;&nbsp;....<a class=nt href=Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir
		xsubir=Replace(Rutdir,Chr(32),"%20")
		xsubir=Left(xsubir,InStrRev(xsubir,"\",-1)-1)
		' response.write xsubir
		saccion   = saccion & "&rutdir=" & xsubir & "   ></a>"

		h3      = Celda
		h1      = ofv.GenCelda("","ut","","","","","",saccion,"si")
		h2      = ofv.GenCelda("","ut","","","","","",  "...","si")
		sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","",h3 & h1 & H2 ,"si")

		If  xTrColor = "#FFFFFF" Then
			xTrColor = "#F9F9F9"
		Else
			xTrColor = "#FFFFFF"
		End If
     else
            celda=  ofv.GenCelda("","ut","","1%","","","","","si")
	  		saccion   = "<i class='las la-folder' style=font-size:12pt;color:" & Session("CMFPNTIT") & "></i>&nbsp;&nbsp;<a class=nt xhref=Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir
	  		xsubir=Replace(Rutdir,Chr(32),"%20")
	  		xsubir=Left(xsubir,InStrRev(xsubir,"\",-1)-1)
	  		' response.write xsubir
	  		saccion   = saccion & "&rutdir=" & xsubir & "   >" & UCase(rutdir) &"</a>"

            h0=celda
	  		h1      = ofv.GenCelda("","ut","","","","","",saccion,"si")
	  		h2      = ofv.GenCelda("","ut","","","","","",  "","si")
	  		saccion = "<a href=uploadn.asp?direct=" & Replace(rutdir   ,Chr(32),"%20") & "&nomdir=" & nomdir & "  target=_new ><i class='las la-upload'></i>&nbsp;&nbsp;</a>"
            h3 = ofv.GenCelda("","ut align=center","","","","","",saccion,"si")
	  		sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","",h0 & h1 & H2 & h3,"si")

	  		If  xTrColor = "#FFFFFF" Then
	  			xTrColor = "#F9F9F9"
	  		Else
	  			xTrColor = "#FFFFFF"
		    End If
	End If


	sfolder=Split(getfolders(rutdir,""),"|")

	For f=0 To UBound(sfolder)-1
		celda=  ofv.GenCelda("","ut","","1%","","","","","si")
		saccion   = "<i class='las la-folder' style=font-size:12pt;color:" & Session("CMFPNTIT") & "></i>&nbsp;&nbsp;<a class=nt href=Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir
		saccion   = saccion & "&rutdir=" & Replace(rutdir  & "\" & sfolder(f) ,Chr(32),"%20") & "   >" & rutdir  & "\" & sfolder(f)  & "</a>"


		h1      = ofv.GenCelda("","ut","","","","","",saccion,"si")
		h2      = ofv.GenCelda("","ut","","","","","",  "","si")
		saccion = "<a href=uploadn.asp?direct=" & Replace(rutdir  & "\" & sfolder(f) ,Chr(32),"%20") & "&nomdir=" & nomdir & "  target=_new ><i class='las la-upload'></i>&nbsp;&nbsp;</a>"
		'sBrow = "<a class=nt href=CordscrBrow.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir & "&browdir=" & direct & "   ><b>+</b></a>"
		h3 = ofv.GenCelda("","ut align=center","","","","","",saccion,"si")

		sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","", CELDA & h1 & H2 & h3 ,"si")

		If  xTrColor = "#FFFFFF" Then
			xTrColor = "#F9F9F9"
		Else
			xTrColor = "#FFFFFF"
		End If

	Next

	sfolder=Split(getfilesinfolder(rutdir,""),"|")

	For f=0 To UBound(sfolder)-1
		xdato=Split(sfolder(f),"###")
		' XREF="CordscrBrow.ASP?OP=DEL"
		saccion   = "Cordscr2.asp?opc=DEL&direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & Replace(rutdir  ,Chr(32),"%20")
		If bSuprimir  Then
			celda= BotonEliminarlnk("Confirma la eliminación de " & xdato(0) & "?","",saccion,f,xdato(0))
		Else

			celda=  ofv.GenCelda("","ut colspan=1","","1%","","","","","si")
		End If

		sinput =  "<i class='las la-file-alt' style=font-size:12pt;color:" & Session("CMFPNTIT") & " ></i>"

		surlx=Replace(rutdir  & "/" & XDATO(0) ,Chr(32)," ")
		surlx=Replace(surlx ,"\","/")

        If session("URLTYP") = "V" Then
			XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
		 End If


		  If session("URLPORT") = "N" Then
			 XMESAURL = TipoSite("")  & request.servervariables("SERVER_NAME") & XMESAURL
		  Else
			 XMESAURL = TipoSite("")  & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
          End If

		If Session("EMPMUN") = "N" Then surlx= XMESAURL & "/coremanager" & surlx
		saccion   = "<a class=nt target=_new HREF='" & surlx & "'>" & XDATO(0)  & "</a>"
		h1      = ofv.GenCelda("","ut","","","","","",sinput & saccion,"si")
		h2      = ofv.GenCelda("","ut","","","","","",xdato(1),"si")
		sHTML   = sHTML & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","22","","", CELDA & h1 & H2,"si")

		If  xTrColor = "#FFFFFF" Then
			xTrColor = "#F9F9F9"
		Else
			xTrColor = "#FFFFFF"
		End If

	Next






	sHTML = sHTML & ofv.GenFTabla()


	Agregar = ""
	Buscar  = ""

	sAccion = "imgfiles.asp"
	Celda   = ofv.BotonVolver(sAccion)

	xxxx    = CorGenBotonXp(3, "NO", "NO", "NO", "NO", "Agregar Subdirectorio", "NO", "Agregar Subdirectorio")
	sAccion = "Agregarsubdir.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir
	Agregar = ofv.BotonAgregar(sAccion)

	sHTML = sHTML & "<script type='text/javascript'>"
	sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
	sHTML = sHTML & "  {"
	sHTML = sHTML & "  var r=confirm(disp_msg);"
	sHTML = sHTML & "  if (r==true)"
	sHTML = sHTML & "    {"
	sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
	'  sHTML = sHTML & "    alert(fr.name);"
	sHTML = sHTML & "    fr.submit();"
	sHTML = sHTML & "    }"
	sHTML = sHTML & "  else"
	sHTML = sHTML & "    {"
	sHTML = sHTML & "    alert('Se cancelo la accion!');"
	sHTML = sHTML & "    }"
	sHTML = sHTML & "  }</script>"

	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)


	sHTML = sHTML & "</body></html>"


	Function BotonEliminarlnk(titdelete,sNombre, sHref,CLAVE,ARCH)

		wktitdelete=titdelete '"Confirma la eliminacion de TODAS las Relaciones?"
		If sNombre="" Then
			sImageELS    = "<i class='las la-times'></i>"
			xclass="ut"
		Else
			sImageELS    =sNombre
			xclass="ttt"
		End If
		sformobj   = "FRMOBJ" & CLAVE
		sAccion    = sHref

		SINPUT="<INPUT TYPE=HIDDEN NAME=ARCH VALUE=""" & ARCH & """>"

		formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","",SINPUT,"si")
		sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

		BotonEliminarlnk=ofv.GenCelda("",xclass & " style='cursor:pointer;'  title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs & formobj ,"si")



	End Function

End If

Response.Write sHTML



%>