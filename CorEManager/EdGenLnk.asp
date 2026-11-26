<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	APLCHECK()
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False

	nblk = request.querystring("nblk")
	If Len(nblk) = 0 Then
		nblk = "S"
		vueltaC  = request.querystring("volver")
		session("volverproc") = ""
	Else
		vueltaC  = ""
	End If

	nblkP = request.querystring("nblkP")

	nTabla   = request.querystring("tabla")
	nid      = request.querystring("id")
	nCodigo  = request.querystring("Codobjeto")
	sTipo    = request.querystring("Tipobjeto")

	'response.write nid & " nid " & "<br>"
	'response.write vueltaC & " vueltaC " & "<br>"

'response.write request.querystring
	If session("volverproc") = "" Then session("volverproc") = vueltaC
	If vueltaC = "" Then vueltaC = session("volverproc")

	volver2 = ofv.vueltaaspEx("../","0","0") & "_Id=" & nId & "_tabla=" & Ntabla & "_nblk=S"

	esttxt = " STYLE={font-family:tahoma;color:#4466bb;font-size:10pt;font-weight:bold;background-color:#ffffff;} "

	Set cn = ofv.conectar(ofv.strconn2)
	Strsql = "SELECT * From " & ntabla & " where objid = " & nID

	''response.write chr(13) & strsql & chr(13)

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then

		Select Case UCase(ntabla)

			Case "OBJACCION"

			wproced = rs(3)
			wID     = rs(18)
			wlink   = rs(19)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(20)
			wDAT    = 20


			Case "OBJESQUEMA2"

			wproced = RS(3)
			wID     = nid
			wlink   = rs(9)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(10)
			wDAT    = 10


			Case "OBJCONTROL2"

			wproced = rs(3)
			wID     = nid
			wlink   = rs(8)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(9)
			wDAT    = 9

			Case "OBJESTDET"

			wproced = nid
			wID     = nid
			wlink   = rs(32)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(33)
			wDAT    = 33

			Case "OBJPORTADA"

			wproced = nid
			wID     = "PORTADA"
			wlink   = rs(32)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(33)
			wDAT    = 33


			Case "REGIMENFIRMAS"

			wproced = nid
			wID     = "REGFIRMAS"
			wlink   = rs(2)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(2)
			wDAT    = 2

			Case "OBJSECTOR"

			wproced = nid
			wID     = "SECTOR"
			wlink   = rs(11)
			If IsNull(wlink) Then wlink = 0
			wtxtl   = rs(11)
			wDAT    = 11


		End Select




	End If

	Call ofv.cerrarconsulta(rs)


	''''''''''''''''''  COMIENZO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	sHTML = ofv.FORMheader("")




	sHTML = sHTML & "<SCRIPT language='javascript'  >" & Chr(13) & Chr(13)
	sHTML = sHTML & "    var myGlos;                           " & Chr(13)
	sHTML = sHTML & "</SCRIPT>" & Chr(13) & Chr(13)

	'' sHTML = sHTML & "<SCRIPT language='javascript' for='LINKTXT' event=""onmouseout"" > alert ('3'); " & chr(13) & chr(13)

	sHTML = sHTML & "<SCRIPT language='javascript'  > var pi;var pf; var st;" & Chr(13) & Chr(13)
	sHTML = sHTML & "   function LinkX() {" & Chr(13)
	sHTML = sHTML & "     var textArea = document.getElementById('LINKTXT');" & Chr(13)
	sHTML = sHTML & "    if (document.selection) { " & Chr(13)
	sHTML = sHTML & "         var bm = document.selection.createRange().getBookmark();" & Chr(13)
	sHTML = sHTML & "         var sel = textArea.createTextRange();" & Chr(13)
	sHTML = sHTML & "         sel.moveToBookmark(bm);" & Chr(13)

	sHTML = sHTML & "         var sleft = textArea.createTextRange();" & Chr(13)
	sHTML = sHTML & "         sleft.collapse(true);" & Chr(13)
	sHTML = sHTML & "         sleft.setEndPoint('EndToStart', sel);" & Chr(13)
	sHTML = sHTML & "         textArea.selectionStart = sleft.text.length;" & Chr(13)
	sHTML = sHTML & "         textArea.selectionEnd = sleft.text.length + sel.text.length;" & Chr(13)
	sHTML = sHTML & "        textArea.selectedText = sel.text;  " & Chr(13)
	sHTML = sHTML & "     }"
	sHTML = sHTML & "     else if (textArea.selectionStart){ " & Chr(13)
	sHTML = sHTML & "        textArea.selectedText = textArea.value.substring(textArea.selectionStart,textArea.selectionEnd);" & Chr(13)
	sHTML = sHTML & "    }" & Chr(13)

	' sHTML = sHTML & "    alert('Selection Start==> ' + textArea.selectionStart + '\n' +" & chr(13)
	' sHTML = sHTML & "      'Selection End  ==> ' + textArea.selectionEnd + '\n' +" & chr(13)
	' sHTML = sHTML & "      'Selected Text  ==> ' + textArea.selectedText + '\n' +" & chr(13)
	'sHTML = sHTML & "        'TextArea Value ==> ' + textArea.value);" & chr(13)

	sHTML = sHTML & " pi=textArea.selectionStart;pf=textArea.selectionEnd;st=textArea.selectedText;}" & Chr(13)

	sHTML = sHTML & " function LinkK () {"

	sHTML = sHTML & " if (document.selection) { "

	sHTML = sHTML & " if (document.selection.type=='Text') {"
	'
	'sHTML = sHTML & " alert(document.selection);"

	' sHTML = sHTML & "     myGlos = document.selection.createRange().text;"

	'sHTML = sHTML & "     myGlos = document.selection.createRange().text;"
	sHTML = sHTML & "     f=LinkX();myGlos=st;"
	sHTML = sHTML & "     document.selection.createRange().text = '||||$' + myGlos + '$||||' ;"
	'sHTML = sHTML & "     windows.getSelection.createRange().text = '||||$' + myGlos + '$||||' ;"
	sHTML = sHTML & "     document.getElementById('TDLINKTXT').style.display = 'none' ; "
	sHTML = sHTML & "     document.getElementById('TD2LINKTXT').style.display = '' ;"
	sHTML = sHTML & "     document.getElementById('TD3LINKTXT').style.display = '' ;"
	sHTML = sHTML & "     document.getElementById('TD4LINKTXT').style.display = '' ;"
	sHTML = sHTML & "     document.getElementById('LINKTXT2').innerHTML = 'El texto asociado es: ' + myGlos;"
	sHTML = sHTML & "     document.getElementById(xnomdato).value = document.getElementById('LINKTXT').innerText;"

	'''' sHTML = sHTML & "     document.selection.createRange().text = '<@>' + myGlos + '<@>' ;"

	''' sHTML = sHTML & "     alert(document.all[xnomdato].value);"
	sHTML = sHTML & " }  "
	sHTML = sHTML & " }  "
	' sHTML = sHTML & " else {  alert('1'); }

	sHTML = sHTML & "     else { "
	'sHTML = sHTML & "    document.getSelection.createRange().text;" = '||||$' + myGlos + '$||||' ;"
	'  sHTML = sHTML & "     myGlos = window.getSelection().toString(); "
	sHTML = sHTML & "    f=LinkX();myGlos=st;"
	'' sHTML = sHTML & "     window.getSelection().text = '||||$' + myGlos + '$||||' ;"
	''  sHTML = sHTML & "    window.getSelection().getRangeAt(0)='||||$' + myGlos + '$||||' ;"
	sHTML = sHTML & "     document.getElementById('TDLINKTXT').style.display = 'none' ;"
	sHTML = sHTML & "     document.getElementById('TD2LINKTXT').style.display = '' ;"
	sHTML = sHTML & "     document.getElementById('TD3LINKTXT').style.display = '' ;"
	sHTML = sHTML & "     document.getElementById('TD4LINKTXT').style.display = '' ;"
	sHTML = sHTML & "     document.getElementById('LINKTXT2').innerText = 'El texto asociado es: ' + myGlos;"
	sHTML = sHTML & "     var str = document.getElementById('LINKTXT').innerHTML; " 'asignamos str al string con toda la cadena'

	'sHTML = sHTML & "     var position=window.getSelection().getRangeAt(0).startOffset;alert(position);" 'guardamos la primera posicion en position'

	'  sHTML = sHTML & "     var position=document.getElementById('LINKTXT').selectionStart;" 'guardamos la primera posicion en position'
	sHTML = sHTML & "     var position1=pi;"
	'  sHTML = sHTML & "     var position= str.indexOf(myGlos) ;" 'guardamos la primera posicion en position'

	'sHTML = sHTML & "     var out= [str.slice(0, position1), '||||$', str.slice(position1)].join('') ;alert(out);"'colocamos ||||$? al principio de la posicion'
	sHTML = sHTML & "     var out= insert(str, '||||$', position1);"
	' sHTML = sHTML & "     var position=document.getElementById('LINKTXT').selectionEnd + 5;" 'guardamos la primera posicion en position'
	sHTML = sHTML & "      position=pf + 5;"
	'  sHTML = sHTML & "     position=window.getSelection().getRangeAt(0).endOffset +5 ;" 'guardamos la ultima posicion en position'

	' sHTML = sHTML & "   position=position + myGlos.length  + 5 ;;" 'guardamos la ultima posicion en position'
	sHTML = sHTML & " var out=insert(out, '$||||', position);"
	' sHTML = sHTML & "     var out= [out.slice(0, position), '$||||', out.slice(position)].join('');"'colocamos $|||| al final de la seleccion'

	'sHTML = sHTML & " alert(document.getElementById(xnomdato).value); alert(out);"
	sHTML = sHTML & "     document.getElementById(xnomdato).value = out;" 'asignamos el string terminado OUT a xnomdato'

	'' sHTML = sHTML & "     document.getElementById(xnomdato).value = document.getElementById('LINKTXT').innerHTML;"
	'"
	'document.getElementById('LINKTXT').innerHTML;"

	sHTML = sHTML & " }"
	sHTML = sHTML & " }"

	sHTML = sHTML & "function insert(main_string, ins_string, pos) {"
	sHTML = sHTML &  "  if(typeof(pos) == 'undefined') {  "
	sHTML = sHTML & "    pos = 0;  "
	sHTML = sHTML & "  }  "
	sHTML = sHTML &  "  if(typeof(ins_string) == 'undefined') {  "
	sHTML = sHTML & "   ins_string = ''; "
	sHTML = sHTML & " }"
	sHTML = sHTML &  " return main_string.slice(0, pos) + ins_string + main_string.slice(pos);  "
	sHTML = sHTML &  "}  "


	sHTML = sHTML & "</SCRIPT>" & Chr(13) & Chr(13)

	''  sHTML = sHTML & "<p>ing: " & request.querystring & "<br><br></p>"

	sHTML = sHTML & "<center><br>"


	If UCase(ntabla) <> "REGIMENFIRMAS" Then
		'----------------------------------------------

		shtml = shtml & CorPnlFnc("N","")

		sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

		'-------------------------------------------------
	End If


	''''''' inserta link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	If nblk = "IL" Then


		wlink   = request.form("comp1") & ""
		wlinkT  = request.form("comp2") & ""
		wlinkN1 = request.form("compN1") & ""
		wlinkN2 = request.form("compN2") & ""
		wlinkN3 = request.form("compN3") & ""
		wlinkN4 = request.form("compN4") & ""
		wlinkN5 = request.form("compN5") & ""
		wlinkN6 = request.form("compN6") & ""



		If wlink > 0 Then

			strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
			strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & nCodigo & ",'" & sTipo & "',"
			strsql = strsql & "'" & wproced & "','" & wId & "',"
			If request.querystring("cmp")="K" Then
				strsql = strsql & wlink & ",'" & wlinkT & "','','','KF','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
			Else
				strsql = strsql & wlink & ",'" & wlinkT & "','','','L','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
			End If
			strsql = strsql & ")"
			Call ofv.crearconsultaEx(StrSql,cn,1,volver)
			If request.querystring("cmp")="K" Then
				Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & wId & "' AND LTIPO = 'KF'  ORDER BY OBJID DESC "
			Else

				Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & wId & "' AND LTIPO = 'L'  ORDER BY OBJID DESC "

			End If

			'	response.write chr(13) & request.querystring & chr(13)

			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs.eof Then
				wlnkid = rs(0)
			Else
				wlnkid = 0
				wlink = 0
			End If
			Call ofv.cerrarconsulta(rs)

			If CLng(wlnkid) > 0 Then
				strsql = "update " & ntabla & " set rellink = " & wlnkid
				strsql = strsql & " where  objid = " & nID
				Call ofv.crearconsultaEx(StrSql,cn,1,volver)
			End If
		End If

		nblk = "S"
	End If

	''''''' FIN inserta link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' inserta link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	If nblk = "II" Then

		wlink = request.form("comp1")


		strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
		strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',"
		strsql = strsql & "'" & wproced & "','" & wId & "',"
		strsql = strsql & "0,'','" & wlink & "','" & wlink & "','I'"
		strsql = strsql & ")"
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		Strsql = "SELECT * From OBJLNK where iden1 = '" & wproced & "' and iden2 = '" & wId & "' AND LTIPO = 'I' ORDER BY OBJID DESC"

		'response.write chr(13) & strsql & chr(13)

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs.eof Then
			wlnkid = rs(0)
		Else
			wlnkid = 0
			wlink = 0
		End If
		Call ofv.cerrarconsulta(rs)

		If CLng(wlnkid) > 0 Then
			strsql = "update " & ntabla & " set rellink = " & wlnkid
			strsql = strsql & " where  objid = " & nID
			Call ofv.crearconsultaEx(StrSql,cn,1,volver)
		End If

		wlink = wlnkid
		nblk = "S"
	End If

	''''''' FIN inserta link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' agrega link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	If nblk = "AL" Then

		wtipoL   = 0
		wtipOBJ  = 0
		wUrl     = "NO"
		wUrlDesc = "NO"
		wtipo    = "NO"

		wNV1    = "0"
		wNV2    = "0"
		wNV3    = "0"

		WWWAPUNTA = ""
		If UCase(ntabla) = "OBJPORTADA" Then
			fdato = request.form("fdatx") & ""
			If fdato = "S" Then
				WWWAPUNTA = " <script language=javascript >  clactfrm('11','" & nCodigo & "','" & sTipo & "','0','0','0','0','0','11','" & nCodigo & "'); document.XGUFRM.submit();</script>"
			End If
		End If


		If UCase(ntabla) = "REGIMENFIRMAS" Then
			fdato = request.form("fdatx") & ""
			If fdato = "S" Then
				WWWAPUNTA = " <script language=javascript >  clactfrm('11','" & nCodigo & "','" & sTipo & "','0','0','0','0','0','11','" & nCodigo & "'); document.XGUFRM.submit();</script>"
			End If
		End If

		'alert('estoy');


		sHTML = sHTML & GenBuscaObj(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0")


		sAccion =  Replace(vueltaC,"../","")
		sAccion =  Replace(sAccion,"_","&")

		'  dim sarasa
		'    sAccion = sarasa


		sHTML = sHTML & "<FORM  accept-charset=""utf-8"" name=""fcomp2x"" ID=""fcomp2"" action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
		sHTML = sHTML & "</FORM>" & Chr(13)



		If bModificar Then
			'sAccion = "EdGenLnk.asp?"
			If request.querystring("cmp")="K" Then
				sAccion = "EdGenLnk.asp?base=K&cmp=K&"
			Else
				sAccion = "EdGenLnk.asp?"
			End If
			sAccion = sAccion & "Id=" & nId  & "&nblk=IL&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
		Else
			sAccion = ""
		End If
		sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
		sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3 ","","60%","",0,0,0,0,"","","","no") 'Movimos la tabla de lugar porque no funcionaba el form en ffox'

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  colspan=2 align=center  >"

		sHTML = sHTML & "<input name='comp1'   type='hidden'  size=1  value='" & XGUFCOBJ & "' class=bt >"
		sHTML = sHTML & "<input name='comp2'   type='hidden'  size=1  value='" & XGUFTOBJ & "' class=bt >"
		sHTML = sHTML & "<input name='compn1'  type='hidden'  size=1  value='" & XGUFNIV1 & "' class=bt >"
		sHTML = sHTML & "<input name='compn2'  type='hidden'  size=1  value='" & XGUFNIV2 & "' class=bt >"
		sHTML = sHTML & "<input name='compn3'  type='hidden'  size=1  value='" & XGUFNIV3 & "' class=bt >"
		sHTML = sHTML & "<input name='compn4'  type='hidden'  size=1  value='" & XGUFNIV4 & "' class=bt >"
		sHTML = sHTML & "<input name='compn5'  type='hidden'  size=1  value='" & XGUFNIV5 & "' class=bt >"
		sHTML = sHTML & "<input name='compn6'  type='hidden'  size=1  value='" & XGUFNIV6 & "' class=bt >"

		''response.write XGUFNIV2 & " XGUFNIV2 linea 335 " & "<br>"
		''response.write XGUFCOBJ & " XGUFCOBJ linea 336 " & "<br>"


		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  align=center   >"

		sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"

		sHTML = sHTML & "</td>"

		sHTML = sHTML & "<td  align=center   >"
		sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
		sHTML = sHTML & " onclick=" & Chr(34) & "document.getElementById('fcomp2').submit();" & Chr(34) &  " >"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</table><br><br>"
		sHTML = sHTML & "</FORM>" & Chr(13)



		nblk = "NO"

	End If
	''''''' FIN agrega link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' agrega link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	If nblk = "AI" Then

		sAccion =  Replace(vueltaC,"../","")
		sAccion =  Replace(sAccion,"_","&")
		sHTML = sHTML & "<FORM  name=fcomp2x ID=fcomp2 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
		sHTML = sHTML & "</FORM>" & Chr(13)


		If bModificar Then
			sAccion = "EdGenLnk.asp?"
			sAccion = sAccion & "Id=" & nId  & "&nblk=II&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
		Else
			sAccion = ""
		End If
		sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
		sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  colspan=2 class=ttt align=center  >"
		sHTML = sHTML & "Agregar&nbsp;Link&nbsp;a&nbsp;Internet "
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"


		sHTML = sHTML & "<tr height=50 >"
		sHTML = sHTML & "<td  colspan=2 align=center  >"
		sHTML = sHTML & ofv.GenerarInput("comp1","www.","agregar","","100","dt","sb")
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  align=center   >"
		sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='Aceptar' class=bt >"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td  align=center   >"
		sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
		sHTML = sHTML & " onclick=" & Chr(34) & "fcomp2.submit();" & Chr(34) &  " >"
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "</table><br><br>"
		sHTML = sHTML & "</FORM>" & Chr(13)

		nblk = "NO"

	End If

	''''''' FIN agrega link internet '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' Alta 1era vez ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If CLng(wlink) = 0 And nblk = "S" Then

		''''Para obtener la descripcion de la Accion
		''Primero obtiene el tipo de accion de la tabla objaccion
		Strsql = "SELECT * From OBJACCION "
		Strsql = Strsql & " WHERE codproced = '" & wProced & "' "
		Strsql = Strsql & " AND objid = '" & nID & "' "
		'response.write Strsql & " Strsql -1-"
		Set rsz = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsz.eof Then
			nTAccion  = rsz(4)
		End If
		Call ofv.cerrarconsulta(rsz)

		''Despues obtiene la descri de la accion de la tabla Tipoaccalt
		Strsql = "Select descri From Tipoaccalt "
		Strsql = Strsql & " Where tipoaccionalt = '" & nTAccion & "' "
		'response.write Strsql & " Strsql -2-"
		Set rsy = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsy.eof Then
			nTdescri  = rsy(0)
			'''response.write nTdescri & " nTdescri " & "<br>"
		End If
		Call ofv.cerrarconsulta(rsy)

		sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=5  ","","60%","",0,0,0,0,"","","","no")
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  colspan=4 class=ttt align=center >"
		sHTML = sHTML & "Agregar&nbsp;Link - "
		sHTML = sHTML & nTdescri
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  COLSPAN=4 align=center   >&nbsp;</TD>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "<tr>"


		sHTML = sHTML & "<td  align=center   >"

		If UCase(ntabla) = "OBJPORTADA" Then

			If bModificar Then
				sAccion = "EdGenLnk.asp?"
				sAccion = sAccion & "Id=" & nId  & "&nblk=AL&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
				sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
			Else
				sAccion = ""
			End If

			sHTML = sHTML & "<FORM  name=fcomp1x ID=fcomp1x action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
			sHTML = sHTML & "<input name=fdatx ID=fdatx type=hidden  size=1 value='S' class=bt >"
			sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='LINK DENTRO DEL DOCUMENTO' class=bt >"
			sHTML = sHTML & "</FORM>" & Chr(13)

		Else
			sHTML = sHTML & "&nbsp;"
		End If

		sHTML = sHTML & "</td>"


		sHTML = sHTML & "<td  align=center   >"

		If bModificar Then
			If request.querystring("base")="K" Then
				sAccion = "EdGenLnk.asp?cmp=K&"
			Else
				sAccion = "EdGenLnk.asp?"
			End If
			sAccion = sAccion & "Id=" & nId  & "&nblk=AL&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
		Else
			sAccion = ""
		End If
		'response.write ("aca esta el accion:" & sAccion)'sarasa)

		sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
		If  request.querystring("base")="K" Then
			sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='ASOCIAR FORMULARIO' class=bt >"
		Else
			sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='LINK A OTRO DOCUMENTO' class=bt >"
		End If
		sHTML = sHTML & "</FORM>" & Chr(13)
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "<td  align=center   >"

		If bModificar Then
			If request.querystring("base")="K" Then
				sAccion = "EdGenLnk.asp?cmp=K&"
			Else
				sAccion = "EdGenLnk.asp?"
			End If
			sAccion = sAccion & "Id=" & nId  & "&nblk=AI&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
			sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
		Else
			sAccion = ""
		End If
		If  request.querystring("base")="K" Then
		Else

			sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
			sHTML = sHTML & "<input  type=SUBMIT  size=10%  value='LINK A INTERNET' class=bt >"
			sHTML = sHTML & "</FORM>" & Chr(13)
			sHTML = sHTML & "</td>"
		End If

		sHTML = sHTML & "<td  align=center   >"

		sAccion =  Replace(vueltaC,"../","")
		sAccion =  Replace(sAccion,"_","&")
		sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
		sHTML = sHTML & "<input  type=submit  size=10%  value='CANCELAR' class=bt >"
		sHTML = sHTML & "</FORM>" & Chr(13)
		sHTML = sHTML & "</td>"
		sHTML = sHTML & "</tr>"

		sHTML = sHTML & "<tr>"
		sHTML = sHTML & "<td  COLSPAN=4 align=center   >&nbsp;</TD>"
		sHTML = sHTML & "</tr>"
		sHTML = sHTML & "</table><br><br>"


		nblk = "NO"


	End If

	''''''' FIN Alta 1era vez ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' modificar link (OBTIENE CLAVE) '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "ML" Or nblk = "MLD" Or nblk = "MI" Or nblk = "MID" Or nblk = "A"  Then


		Strsql = "SELECT * From " & ntabla & " where objid = " & nID

		'response.write chr(13) & strsql & chr(13)

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs.eof Then

			Select Case UCase(ntabla)

				Case "OBJACCION"


				wlinkEY   = rs(19)



				Case "OBJESQUEMA2"


				wlinkEY   = rs(9)



				Case "OBJCONTROL2"


				wlinkEY   = rs(8)


				Case "OBJESTDET"


				wlinkEY   = rs(32)

				Case "OBJPORTADA"


				wlinkEY   = rs(32)

				Case "REGIMENFIRMAS"


				wlinkEY   = rs(2)

				Case "OBJSECTOR"


				wlinkEY   = rs(11)

			End Select

		End If


	End If


	''''''' modificar link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "ML" Then

		wlinkD = request.form("comp2") & ""

		strsql = "update OBJLNK set LurlDesc = '" & wlinkD & "' "
		strsql = strsql & "where objid = " & wlinkEY
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		nblk = "S"

	End If

	''''''' FIN modificar link local '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' modificar link local D '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "MLD" Then

		wlink   = request.form("comp1") & ""
		wlinkT  = request.form("comp2") & ""
		wlinkN1 = request.form("compN1") & ""
		wlinkN2 = request.form("compN2") & ""
		wlinkN3 = request.form("compN3") & ""
		wlinkN4 = request.form("compN4") & ""
		wlinkN5 = request.form("compN5") & ""
		wlinkN6 = request.form("compN6") & ""

		If wlink > "0" Then

			strsql = "update OBJLNK set codobjetoL = " & wlink & " "
			strsql = strsql & ", tipobjetoL = '" & wlinkT & "' "
			strsql = strsql & ", IDEN3L = '" & wlinkN1 & "' "
			strsql = strsql & ", IDEN4  = '" & wlinkN2 & "' "
			strsql = strsql & ", IDEN4L = '" & wlinkN3 & "' "
			strsql = strsql & "where objid = " & wlinkEY
			Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		End If

		nblk = "S"
	End If


	''''''' FIN modificar link local D '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''' modificar link internet ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "MI" Then

		wlink = request.form("comp1")

		'''  wlinkD = request.form("comp2")



		strsql = "update OBJLNK set Lurl = '" & wlink & "' "
		strsql = strsql & "where objid = " & wlinkEY
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)



		nblk = "S"
	End If
	''''''' FIN modificar link internet ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''' modificar link internet D''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "MID" Then

		'''  wlink = request.form("comp1")

		wlinkD = request.form("comp2")



		strsql = "update OBJLNK set LurlDesc = '" & wlinkD & "' "
		strsql = strsql & "where objid = " & wlinkEY
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)



		nblk = "S"
	End If
	''''''' FIN modificar link internet D ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	''''''' anular link ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "A" Then

		'    response.write "anular link: " & "<br>"

		sAccion =  Replace(vueltaC,"../","")
		sAccion =  Replace(sAccion,"_","&")

		strsql = "delete from OBJLNK "
		strsql = strsql & "where objid = " & wlinkEY
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		' response.write "anular link: " & strsql & "<br>"

		strsql = "update " & ntabla & " set rellink = 0, reflink = '' "
		strsql = strsql & " where  objid = " & nID
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		response.redirect sAccion

		nblk = "NO"

	End If
	''''''' FIN anular link ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



	''''''' Mostrar link '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	If nblk = "S" Or nblk = "SD" Then

		Strsql = "SELECT * From " & ntabla & " where objid = " & nID

		'response.write chr(13) & strsql & chr(13)

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs.eof Then

			Select Case UCase(ntabla)

				Case "OBJACCION"

				wproced = rs(3)
				wID     = rs(18)
				wlink   = rs(19)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(20)
				wDAT    = 20


				Case "OBJESQUEMA2"

				wproced = RS(3)
				wID     = nAccion
				wlink   = rs(9)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(10)
				wDAT    = 10


				Case "OBJCONTROL2"

				wproced = rs(3)
				wID     = nAccion
				wlink   = rs(8)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(9)
				wDAT    = 9

				Case "OBJESTDET"

				wproced = nid
				wID     = nid
				wlink   = rs(32)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(33)
				wDAT    = 33


				Case "OBJPORTADA"

				wproced = nid
				wID     = "PORTADA"
				wlink   = rs(32)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(33)
				wDAT    = 33

				Case "REGIMENFIRMAS"

				wproced = nid
				wID     = "REGFIRMAS"
				wlink   = rs(2)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(2)
				wDAT    = 2

				Case "OBJSECTOR"

				wproced = nid
				wID     = "SECTOR"
				wlink   = rs(11)
				If IsNull(wlink) Then wlink = 0
				wtxtl   = rs(11)
				wDAT    = 11


			End Select


			wtipoL   = 0
			wtipOBJ  = 0
			wUrl     = "NO"
			wUrlDesc = "NO"
			wtipo    = "NO"

			wNV1    = "0"
			wNV2    = "0"
			wNV3    = "0"

			Strsql = "SELECT * From objlnk where objid = " & wlink

			'response.write chr(13) & strsql & chr(13)

			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs1.eof Then
				wtipoL   = CLng(rs1(6))
				wtipOBJ  = rs1(7)
				wUrl     = rs1(8)
				wUrlDesc = rs1(9)
				wtipo    = rs1(10)

				wNV1    = rs1(11) & ""
				wNV2    = rs1(12) & ""
				wNV3    = rs1(13) & ""

				If  WNV1 = "" Then WNV1 = "0"
				If  WNV2 = "" Then WNV2 = "0"
				If  WNV3 = "" Then WNV3 = "0"


				wasotxt     = rs1(14) & ""
				wasotxtstr  = rs1(15)
				wasotxtcan  = rs1(16)
				wasotxtdat  = rs1(17) & ""

			End If
			Call ofv.cerrarconsulta(rs1)

			' response.write  wtipoL &  wUrl & wUrlDesc & wtipo & "<br>"

			'''----'''''' LINK A INTERNET O EXTERNO '''''''''''''''''''''''''''''''''''''''


			If wtipo = "I" Then

				'-------- muestra link Internet

				sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

				sHTML = sHTML & "<tr>"
				sHTML = sHTML & "<td  colspan=2 class=ttt >"
				sHTML = sHTML & "Link&nbsp;a&nbsp;direccion"
				sHTML = sHTML & "</td>"
				sHTML = sHTML & "</tr>"

				If bModificar Then
					sAccion = "EdGenLnk.asp?"
					sAccion = sAccion & "Id=" & nId  & "&nblk=MI&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
					sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
				Else
					sAccion = ""
				End If



				sHTML = sHTML & "<tr height=40 >"
				sHTML = sHTML & "<td  colspan=2 class=uti  >"
				sHTML = sHTML & "<FORM  name=fcomp1 ID=fcomp1 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
				sHTML = sHTML & ofv.GenerarInput("comp1",WUrl,"textsb","fcomp1","100","dt","sb")
				sHTML = sHTML & "</FORM>" & Chr(13)
				sHTML = sHTML & "</td>"
				sHTML = sHTML & "</tr>"







				'-------- FIN muestra link Internet


			Else

				'-------- muestra link local

				''''Para obtener la descripcion de la Accion
				''Primero obtiene el tipo de accion de la tabla objaccion
				Strsql = "SELECT * From OBJACCION "
				Strsql = Strsql & " WHERE codproced = '" & wProced & "' "
				Strsql = Strsql & " AND objid = '" & nID & "' "
				'response.write Strsql & " Strsql -1-"
				Set rsz = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsz.eof Then
					nTAccion  = rsz(4)
				End If
				Call ofv.cerrarconsulta(rsz)

				''Despues obtiene la descri de la accion de la tabla Tipoaccalt
				Strsql = "Select descri From Tipoaccalt "
				Strsql = Strsql & " Where tipoaccionalt = '" & nTAccion & "' "
				'response.write Strsql & " Strsql -2-"
				Set rsy = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rsy.eof Then
					nTdescri  = rsy(0)
					'''response.write nTdescri & " nTdescri " & "<br>"
				End If
				Call ofv.cerrarconsulta(rsy)

				''''

				If nblk = "SD" Then


					sHTML = sHTML & GenBuscaObj(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0")
					sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

					sHTML = sHTML & "<tr>"
					sHTML = sHTML & "<td  align=center   >"

					sHTML = sHTML & "<input  type=Button size=10%  value='Aceptar' class=bt "
					sHTML = sHTML & " onclick=" & Chr(34) & "fcomp3.submit();" & Chr(34) &  " >"
					sHTML = sHTML & "</td>"

					sHTML = sHTML & "<td  align=center   >"
					sHTML = sHTML & "<input  type=Button  size=10%  value='Cancelar' class=bt "
					sHTML = sHTML & " onclick=" & Chr(34) & "fcomp2.submit();" & Chr(34) &  " >"
					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"

					sHTML = sHTML & "</FORM>" & Chr(13)

					sHTML = sHTML & "</table><br><br>"

					sAccion =  Replace(vueltaC,"../","")
					sAccion =  Replace(sAccion,"_","&")

					sAccion = "EdGenLnk.asp?"
					sAccion = sAccion & "Id=" & nId  & "&nblk=S&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
					sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC

					sHTML = sHTML & "<FORM  name=fcomp2 ID=fcomp2 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
					sHTML = sHTML & "</FORM>" & Chr(13)

					If bModificar And XGUFCOBJ > 0  Then
						sAccion = "EdGenLnk.asp?"
						sAccion = sAccion & "Id=" & nId  & "&nblk=MLD&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
						sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
					Else
						sAccion = ""
					End If
					sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
					sHTML = sHTML & "<input name='comp1'   type='hidden'  size=1  value='" & XGUFCOBJ & "' class=bt >"
					sHTML = sHTML & "<input name='comp2'   type='hidden'  size=1  value='" & XGUFTOBJ & "' class=bt >"
					sHTML = sHTML & "<input name='compn1'  type='hidden'  size=1  value='" & XGUFNIV1 & "' class=bt >"
					sHTML = sHTML & "<input name='compn2'  type='hidden'  size=1  value='" & XGUFNIV2 & "' class=bt >"
					sHTML = sHTML & "<input name='compn3'  type='hidden'  size=1  value='" & XGUFNIV3 & "' class=bt >"
					sHTML = sHTML & "<input name='compn4'  type='hidden'  size=1  value='" & XGUFNIV4 & "' class=bt >"
					sHTML = sHTML & "<input name='compn5'  type='hidden'  size=1  value='" & XGUFNIV5 & "' class=bt >"
					sHTML = sHTML & "<input name='compn6'  type='hidden'  size=1  value='" & XGUFNIV6 & "' class=bt >"
					sHTML = sHTML & "</FORM>" & Chr(13)

					'''response.write XGUFNIV2 & " XGUFCOBJ linea 965 " & "<br>"
					'''response.write XGUFCOBJ & " XGUFTOBJ linea 966 " & "<br>"
					'''response.write XGUFNIV1 & " XGUFNIV1 linea 967 " & "<br>"
					'''response.write XGUFNIV2 & " XGUFNIV2 linea 968 " & "<br>"
					'''response.write XGUFNIV3 & " XGUFNIV3 linea 969 " & "<br>"
					'''esponse.write XGUFNIV4 & " XGUFNIV4 linea 970 " & "<br>"
					'''response.write XGUFNIV5 & " XGUFNIV5 linea 971 " & "<br>"
					'''response.write XGUFNIV6 & " XGUFNIV6 linea 972 " & "<br>"



				Else



					sHTML = sHTML & ofv.GenTabla("panel","aaa cellpadding=3  ","","60%","",0,0,0,0,"","","","no")

					sHTML = sHTML & "<tr>"
					sHTML = sHTML & "<td  colspan=2 class=ttt >"
					sHTML = sHTML & "Link&nbsp;a&nbsp;documento - "
					sHTML = sHTML & nTdescri
					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"


					sHTML = sHTML & "<tr>"
					sHTML = sHTML & "<td  colspan=2 align=center  >"





					XMST = "0"

					If wtipoL <> 0 Then

						XMST = "1"

						Call GETLISTOBJ("E",wtipoL,wtipOBJ)
						XGUFCOBJKEY = XGUTABLISKEY
						XGUFCOBJDAT = XGUTABLISDAT


						XGUFBTDF0 = XGUFBTDF0 & "<TR  "
						XGUFBTDF0 = XGUFBTDF0 & " >"

						XGUFBTDF0 = XGUFBTDF0 & "<TD CLASS=UT   >"
						XGUFBTDF0 = XGUFBTDF0 & XGUFCOBJDAT & "</TD>"
						XGUFBTDF0 = XGUFBTDF0 & "</TR>"

						If WNV1 <> "0" Then

							XMST = "2"


							Call GETLISTSLP("E",wtipoL,wtipOBJ,WNV1)
							XGUFCOBJKEY = XGUTABLISKEY
							XGUFCOBJDAT = XGUTABLISDAT


							XGUFBTDF1 = XGUFBTDF1 & "<TR  "
							XGUFBTDF1 = XGUFBTDF1 & " >"

							XGUFBTDF1 = XGUFBTDF1 & "<TD CLASS=UT   >"
							XGUFBTDF1 = XGUFBTDF1 & XGUFCOBJDAT & "</TD>"
							XGUFBTDF1 = XGUFBTDF1 & "</TR>"

							If WNV2 <> "0" Then

								XMST = "3"

								Call GETLISTNV2("E",wtipoL,wtipOBJ,WNV1,WNV2)
								XGUFCOBJKEY = XGUTABLISKEY
								XGUFCOBJDAT = XGUTABLISDAT


								XGUFBTDF2 = XGUFBTDF2 & "<TR  "
								XGUFBTDF2 = XGUFBTDF2 & " >"

								XGUFBTDF2 = XGUFBTDF2 & "<TD CLASS=UT   >"
								XGUFBTDF2 = XGUFBTDF2 & XGUFCOBJDAT & "</TD>"
								XGUFBTDF2 = XGUFBTDF2 & "</TR>"

								If WNV3 <> "0" Then

									XMST = "4"

									Call GETLISTNV3("E",wtipoL,wtipOBJ,WNV1,WNV2,WNV3)
									XGUFCOBJKEY = XGUTABLISKEY
									XGUFCOBJDAT = XGUTABLISDAT


									XGUFBTDF3 = XGUFBTDF3 & "<TR  "
									XGUFBTDF3 = XGUFBTDF3 & " >"

									XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT   >"
									XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
									XGUFBTDF3 = XGUFBTDF3 & "</TR>"


								End If
							End If
						End If
					End If



					If XMST > "0" Then

						XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR0""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
						XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF0 & "</TABLE>"


					End If

					If XMST > "1" Then

						XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR1""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
						XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF1 & "</TABLE>"


					End If

					If XMST > "2" Then

						XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR2""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
						XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF2 & "</TABLE>"


					End If

					If XMST > "3" Then

						XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSR3""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
						XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


					End If

					If XMST > "0" Then
						If request.querystring("base")="K"  Then
							sAccion = "EdGenLnk.asp?BASE=K&CMP=K&"
						Else


							sAccion = "EdGenLnk.asp?"
						End If
						sAccion = sAccion & "Id=" & nId  & "&nblk=SD&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
						sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC

						XGUFCOBJDAT = ""
						XGUFCOBJDAT = XGUFCOBJDAT & "<FORM  name=fcomp2 ID=fcomp2 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
						XGUFCOBJDAT = XGUFCOBJDAT & "<input  type=SUBMIT  size=10%  value='MODIFICAR' class=bt >"
						XGUFCOBJDAT = XGUFCOBJDAT & "</FORM>" & Chr(13)

						XGUFBTDF3 = ""
						XGUFBTDF3 = XGUFBTDF3 & "<TR  "
						XGUFBTDF3 = XGUFBTDF3 & " >"

						XGUFBTDF3 = XGUFBTDF3 & "<TD CLASS=UT ALIGN=RIGHT  >"
						XGUFBTDF3 = XGUFBTDF3 & XGUFCOBJDAT & "</TD>"
						XGUFBTDF3 = XGUFBTDF3 & "</TR>"

						XGUFBTAB1 = XGUFBTAB1 & "<table id=""MTBSUSRB""  class='aaa' width=100% cellspacing=1 cellpadding=2 "
						XGUFBTAB1 = XGUFBTAB1 & " >" & XGUFBTDF3 & "</TABLE>"


					End If


					sHTML = sHTML & XGUFBTAB1

					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"

					sHTML = sHTML & "<tr STYLE='display:none;' >"
					sHTML = sHTML & "<td  colspan=2 align=center  >"

					sHTML = sHTML & "<font class=at><select size=10  name=comp1  "
					sHTML = sHTML & " onchange=" & Chr(34) & "fcomp1.submit();" & Chr(34)
					sHTML = sHTML & " >"

					'Strsql = "Select * From csGenObjetos where estado = 3  OR ESTADO = 1 order by xdescri"

					'response.write chr(13) & strsql & chr(13)

					' set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

					' do while not rs2.eof
					'    sHTML = sHTML & "<option value=" & rs2(0)
					'   if clng(rs2(0)) = clng(wtipoL) then sHTML = sHTML & " selected "
					'   sHTML = sHTML & " > " & rs2(1) & "</option>"

					'  rs2.movenext
					'loop
					'call ofv.cerrarconsulta(rs2)
					sHTML = sHTML & "</select></font>"

					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"


					'-------- FIN muestra link local

				End If

			End If

			If nblk = "S" Then


				'''''''''''''''''''''''''''''''''''''''''''''''''''''' BLOQUE ASOCIACION AL TEXTO '''''''''''''''''''''''''''''''''''''''''''


				'----------' agregado: asociado a texto/sin asociar al texto ''''

				Strsql = "SELECT * From objlnk where objid = " & wlink

				'response.write chr(13) & strsql & chr(13)

				Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rs1.eof Then

					'RESPONSE.WRITE STRSQL & "<BR>" & wasotxt & "<BR>"


					wasotxt     = rs1(14) & ""
					wasotxtstr  = CLng("0" & retnum( rs1(15) & ""))
					wasotxtcan  = CLng("0" & retnum(rs1(16) & "" ))
					wasotxtdat  = rs1(17) & ""

					wasodisplay = ""


					Select Case UCase(ntabla)

						Case "OBJACCION"


						wtxtl2   = rs(11)
						wDAT2    = 11


						Case "OBJESQUEMA2"


						wtxtl2   = rs(6)
						wDAT2    = 6


						Case "OBJCONTROL2"


						wtxtl2   = rs(4)
						wDAT2    = 4

						Case "OBJESTDET"


						wtxtl2   = rs(8)
						wDAT2    = 8


						Case "OBJPORTADA"


						wtxtl2   = rs(8)
						wDAT2    = 8
						wasodisplay = " style='display:none;' "

						Case "REGIMENFIRMAS"


						wtxtl2   = rs(3)
						wDAT2    = 3
						wasodisplay = " style='display:none;' "

						'' case "OBJSECTOR"
						''
						''	     wtxtl2  = rs(12)
						''        wDAT2   = 12

					End Select

					If wasotxt = "" Then wasotxt = "N"

					sHTML = sHTML & "<tr " & wasodisplay & " >"
					sHTML = sHTML & "<td  colspan=2 class=ttt  >"
					sHTML = sHTML & "Presentacion&nbsp;del&nbsp;link"
					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"  & Chr(13)

					If bModificar Then
						sAccion = "asp/ActualizarEN.asp?ID=" & WLINK & "&Tabla=OBJLNK&volver=" & volver2
						sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
					Else
						sAccion = ""
					End If


					sHTML = sHTML & "<tr " & wasodisplay & " >"
					sHTML = sHTML & "<td  colspan=2 class=uti  >"
					sHTML = sHTML & "<FORM  name=fcomp9 ID=fcomp9 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
					scomK = "S" & Chr(9) & "N"
					scomD = "ASOCIADO AL TEXTO" & Chr(9) & "SIN ASOCIAR"


					If bModificar Then
						sInput = ofv.gencomboFX("fcomp9",RS1(14).NAME,wasotxt,scomK,scomD," ")
					Else
						sInput = ofv.GenerarInput("",wasotxt,"readonly","","10","dt","")
					End If

					sHTML = sHTML & sInput

					sHTML = sHTML & ofv.GenerarInput(RS1(15).name,"0","hidden","","10","","")
					sHTML = sHTML & ofv.GenerarInput(RS1(16).name,"0","hidden","","10","","")
					sHTML = sHTML & ofv.GenerarInput(RS1(17).name,rs(WDAT2).name,"hidden","","10","","")

					sHTML = sHTML & "</FORM>" & Chr(13)
					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"




					'response.write wtxtl2
					If wasotxt = "S" Then

						xdesde =  wasotxtstr
						xhasta =  wasotxtcan


						If InStr(wtxtl2 & "","||||$") > 0 Then

							wtxtl2 = Replace(wtxtl2,"<br>",Chr(13))
							wtxtl2 = Replace(wtxtl2,"<BR>",Chr(13))

							xdesde =  InStr(wtxtl2 & "","||||$")
							xhasta =  InStr(wtxtl2 & "","$||||") - xdesde - 5

							wtxtl2 = Replace(wtxtl2,"||||$","")
							wtxtl2 = Replace(wtxtl2,"$||||","")


							Strsql = "update " & ntabla & " set " & rs(WDAT2).name & " = '" & wtxtl2 & "'  where objid = " & nID
							'' response.write "<br><br><br><br>" & strsql & "<br>"
							Call ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if session("APLMDCLI")="BC" THEN
                             Strsql = "update objlnk set LurlDesc='' ,postxt = " & xdesde & ", canttxt = " & xhasta & "   where objid = " & wlink
                             wUrlDesc=""
                            ELSE
							 Strsql = "update objlnk set postxt = " & xdesde & ", canttxt = " & xhasta & "   where objid = " & wlink
							END IF
							'response.write "<br><br><br><br>" & strsql & "<br>"
							Call ofv.crearconsultaEx(StrSql,cn,1,parametros)


						End If

						'response.write  xdesde & "   " & xhasta
						If  xdesde = 0 And xhasta = 0 Then




							If bModificar Then

								sAccion = "asp/ActualizarEN.asp?ID=" & nId & "&Tabla=" & ntabla & "&volver=" & volver2
								sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

							Else
								sAccion = ""
							End If


							sHTML = sHTML & "<tr " & wasodisplay & " >"
							sHTML = sHTML & "<td  colspan=2 class=uti id='TDLINKTXT' style='display:block;' >"

							t3b = "<textarea cols=120 rows=20 class='DINDC' STYLE=""height:60%;overflow:'auto';"" "
							t3b = t3b & "name='LINKTXT' id='LINKTXT' onclick=""LinkK();"" "
							t3b = t3b & "  >"
							xdatos = wtxtl2 'replace(wtxtl2 & "",chr(13),"<br>")
							t3b = t3b & xdatos & "</textarea>"


							sHTML = sHTML & t3b
							sHTML = sHTML & "</td>"
							sHTML = sHTML & "</tr>"

							sHTML = sHTML & "<tr " & wasodisplay & " >"
							sHTML = sHTML & "<td  colspan=2 class=ut id='TD2LINKTXT' style='display:none;' >"

							t3b = "<DIV  STYLE=""height:60%;overflow:'auto';"" "
							t3b = t3b & "name='LINKTXT2' id='LINKTXT2' "
							t3b = t3b & "  >xx"
							t3b = t3b & "</div>"

							sHTML = sHTML & "<script language='javascript' >"
							sHTML = sHTML & " var xnomdato = '" & rs(WDAT2).name & "'; "
							sHTML = sHTML & "</script>"

							sHTML = sHTML & t3b
							'response.write sAccion
							sHTML = sHTML & "<FORM style='display:none;' name=fcomp10 ID=fcomp10 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)

							sHTML = sHTML & ofv.GenerarInput(rs(WDAT2).name,rs(WDAT2),"hidden","","10","","")


							If UCase(NTABLA) = "OBJACCION" Then

								sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
								sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")

							End If


							sHTML = sHTML & "</FORM>" & Chr(13)

							sHTML = sHTML & "</td>"
							sHTML = sHTML & "</tr>"


							sHTML = sHTML & "<tr " & wasodisplay & " >"
							sHTML = sHTML & "<td align='right' class=ut id='TD3LINKTXT' style='display:none;' >"
							sHTML = sHTML & "<FORM   action=''  method=post >" & Chr(13)
							sHTML = sHTML & "<input  type=Button size=10%  value='aceptar' class=bt "
							sHTML = sHTML & " onclick=" & Chr(34) & "fcomp10.submit();" & Chr(34) &  " >"
							sHTML = sHTML & "</td>"


							sHTML = sHTML & "</FORM>" & Chr(13)
							sHTML = sHTML & "</td>"

							sAccion = "EdGenLnk.asp?"
							sAccion = sAccion & "Id=" & nId  & "&nblk=" & nblk & "&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
							sAccion = sAccion & "&Tipobjeto=" & sTipo


							sHTML = sHTML & "<td  align='left' class=ut id='TD4LINKTXT' style='display:none;' >"
							sHTML = sHTML & "<FORM   action='" & sAccion & "'  method=post >" & Chr(13)

							sHTML = sHTML & "<input  type=submit size=10%  value='cancelar' class=bt "
							sHTML = sHTML & "  >"
							sHTML = sHTML & "</td>"

							sHTML = sHTML & "</FORM>" & Chr(13)
							sHTML = sHTML & "</td>"

							sHTML = sHTML & "</tr>"

						Else

							sHTML = sHTML & "<tr " & wasodisplay & " >"
							sHTML = sHTML & "<td  colspan=2 class=uti id='TDLINKTXTx' style='display:block;' >"

							t3b = "<DIV  STYLE=""height:80%;overflow:'auto';border-style:solid;border-width:1;border-color:#dddddd;"" "
							t3b = t3b & "name='LINKTXTx' id='LINKTXTx' "
							t3b = t3b & "  >"

							xdatos = wtxtl2
							xdatos = Left(xdatos,xdesde - 1) & "<u><b>" & Mid(xdatos,xdesde,xhasta) & "</b></u>" & Right(xdatos,Len(xdatos) - (xdesde + xhasta - 1))
							xdatos = Replace(xdatos,Chr(13),"<br>")

							t3b = t3b & xdatos & "</div>"
							sHTML = sHTML & t3b

							sHTML = sHTML & "</td>"
							sHTML = sHTML & "</tr>"

							sHTML = sHTML & "<tr " & wasodisplay & " >"

							If bModificar Then
								sAccion = "asp/ActualizarEN.asp?ID=" & WLINK & "&Tabla=OBJLNK&volver=" & volver2
								sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

							Else
								sAccion = ""
							End If

							sHTML = sHTML & "<td colspan=2 align='right' class=ut id='TD4LINKTXT' style='display:block;' >"
							sHTML = sHTML & "<FORM   action='" & sAccion & "'  method=post >" & Chr(13)
							sHTML = sHTML & ofv.GenerarInput(RS1(15).name,"0","hidden","","10","","")
							sHTML = sHTML & ofv.GenerarInput(RS1(16).name,"0","hidden","","10","","")
							sHTML = sHTML & ofv.GenerarInput(RS1(17).name,rs(WDAT2).name,"hidden","","10","","")
							sHTML = sHTML & "<input  type=submit size=10%  value='modificar' class=bt "
							sHTML = sHTML & "  >"

							sHTML = sHTML & "</FORM>" & Chr(13)
							sHTML = sHTML & "</td>"


							sHTML = sHTML & "</tr>"

						End If

					Else

						sHTML = sHTML & "<tr " & wasodisplay & " >"
						sHTML = sHTML & "<td  colspan=2 class=ttt  >"
						sHTML = sHTML & "Referencia&nbsp;para&nbsp;el&nbsp;link"
						sHTML = sHTML & "</td>"
						sHTML = sHTML & "</tr>"  & Chr(13)


						sHTML = sHTML & "<tr " & wasodisplay & " >"
						sHTML = sHTML & "<td  colspan=2 class=uti  >"
						If bModificar Then
							sAccion = "asp/ActualizarEN.asp?ID=" & nId & "&Tabla=" & ntabla & "&volver=" & volver2
							sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
						Else
							sAccion = ""
						End If

						sHTML = sHTML & "<FORM  name=fcomp0 ID=fcomp0 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)

						t3b = "<font class=at><textarea class=dt "
						t3b = t3b & "name=" & rs(WDAT).name & " cols=90 rows=6 "
						If bModificar Then
							t3b = t3b & " onblur=fcomp0.submit() "
							t3b = t3b & " onkeypress=" & Chr(34)
							t3b = t3b & " if (this.value.length > 490) {alert('El texto es demasiado"
							t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"
						Else
							t3b = t3b & " readonly >"
						End If
						xdatos = wtxtl
						t3b = t3b & xdatos & "</textarea></font>"

						sHTML = sHTML & t3b

						If UCase(NTABLA) = "OBJACCION" Then
							sHTML = sHTML & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","10","","")
							sHTML = sHTML & ofv.GenerarInput(rs(4).name,rs(4),"hidden","","10","","")
						End If
						sHTML = sHTML & "</FORM>" & Chr(13)
						sHTML = sHTML & "</td>"
						sHTML = sHTML & "</tr>"



					End If

					If bModificar Then
						sAccion = "EdGenLnk.asp?"
						sAccion = sAccion & "Id=" & nId  & "&nblk="

						If wtipo = "I" Then
							sAccion = sAccion & "MID"
						Else
							sAccion = sAccion & "ML"
						End If


						sAccion = sAccion & "&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
						sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
					Else
						sAccion = ""
					End If



					sHTML = sHTML & "</table><FORM  name=fcompn ID=fcompn action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)

					sHTML = sHTML & "<table><tr>"
					sHTML = sHTML & "<td  colspan=2 class=ttt height=40 >"
					sHTML = sHTML & "Descripcion&nbsp;para&nbsp;el&nbsp;link"
					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"  & Chr(13)

					sHTML = sHTML & "<tr>"
					sHTML = sHTML & "<td  colspan=2 class=uti height=40 >"

					sHTML = sHTML & ofv.GenerarInput("comp2",wUrlDesc,"textsb","fcompn","100","dt","sb")

					sHTML = sHTML & "</td>"
					sHTML = sHTML & "</tr>"
					' RESPONSE.WRITE REQUEST.FORM

				End If
				Call ofv.cerrarconsulta(rs1)

				'--------------FIN ASOCIACION AL TEXTO ''''''''''''''''''''''''''''''''''''''''''''''''''''''''


				'    END IF

				'  end if



				'''''''''''''''''''''''''''' FINAL : BOTONES DE ANULAR O VOLVER '''''''''''''''''''''''''''''''


				'    if nblk = "S" then


				sHTML = sHTML & "<tr>"
				sHTML = sHTML & "<td  align=center   >"

				sHTML = sHTML & "<input  type=Button size=10%  value='Anular Link' class=bt "
				sHTML = sHTML & " onclick=" & Chr(34) & "fcomp3.submit();" & Chr(34) &  " >"
				sHTML = sHTML & "</td>"

				sHTML = sHTML & "<td  align=center   >"
				sHTML = sHTML & "<input  type=Button  size=10%  value='Volver' class=bt "
				sHTML = sHTML & " onclick=" & Chr(34) & "fcomp5.submit();" & Chr(34) &  " >"
				sHTML = sHTML & "</td>"
				sHTML = sHTML & "</tr>"

				sHTML = sHTML & "</table></FORM>"
				'sHTML = sHTML & "</table><br><br>"
				sHTML = sHTML & "<br><br>"

				sAccion =  Replace(vueltaC,"../","")
				sAccion =  Replace(sAccion,"_","&")
				sHTML = sHTML & "<FORM  name=fcomp5 ID=fcomp5 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
				sHTML = sHTML & "</FORM>" & Chr(13)

				If bModificar Then
					sAccion = "EdGenLnk.asp?"
					sAccion = sAccion & "Id=" & nId  & "&nblk=A&Codobjeto=" & nCodigo & "&tabla=" & Ntabla
					sAccion = sAccion & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC
				Else
					sAccion = ""
				End If
				sHTML = sHTML & "<FORM  name=fcomp3 ID=fcomp3 action=" & Chr(34) & sAccion  & Chr(34) &  "  method=post >" & Chr(13)
				sHTML = sHTML & "</FORM>" & Chr(13)

			End If

		End If
		Call ofv.cerrarconsulta(rs)

	End If

	''''''' FIN Mostrar link '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




	Call ofv.cerrarconn(cn)

	Agregar = ""
	Buscar  = ""
	Celda   = ""




	sHTML = sHTML & "</center>" & WWWAPUNTA


	If UCase(ntabla) <> "REGIMENFIRMAS" Then
		'--------------------------------------------------------------------------------
		sHTML = sHTML & "</span></div>"

		Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		Fila = ofv.GenRow("","","","","","","",Celda,"si")
		sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

		'Estructura del Indice

		sHTML = sHTML & CorPnlIndx("N","")

		'--------------------------------------------------------------------------------
	End If

	sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,celda)

	sHTML = sHTML & "</body></html>"
End If
If nblk <> "A" Then response.write sHTML

%>