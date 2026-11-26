<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0


	CantRegMoverx = request.querystring("retorno")

	If CantRegMoverx = "" Then  CantRegMoverx = request.querystring("Cantregmover")

	nEditar=CDbl("0" &request.querystring("Editar"))

	diasExtendido=request.Form("diasextender")


	ID = request.querystring("ID")
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID & "_Editar=" & nEditar

	sHTML = ofv.FormHeader(session("FormN") & " - Propiedades del Mensaje")

	'Encabezados
	Formulario = ""


	'	response.write cdate (42973 + clng(cdate("01-01-1900")))





	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	Set cn = ofv.conectar(ofv.strconn2)

	posi = 0
	sImagenI4 = "<B><FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"

	' Actualizacion de Tablas al Usar la opcion de Editar

	If nEditar=1   And  diasExtendido <> "" Then

		msgcopete=request.form ("msgcopete")
		msgtitulo=request.form ("msgtitulo")
		texto=request.form ("texto")

		strsql= "UPDATE novegenpar  "
		strsql= strsql & " set  msgtitulo='" & msgtitulo & "', msgcopete ='"& msgcopete&"', texto ='"& texto & "',fecexp= fecexp+"& +diasExtendido &" , dias =dias +"& + diasExtendido
		strsql=  strsql & " where objid='"& ID &"'"

		'    response.write strsql &"<Br>"

		Set rs2= ofv.crearconsultaEx(StrSql,cn,1,parametros)


		strsql= "UPDATE novedades  "
		strsql= strsql & " set  msgtitulo='" & msgtitulo& "', msgcopete ='"& msgcopete &"', mensaje ='"& texto & "',fecexp=fecexp+"& diasExtendido &""
		strsql=  strsql & " where codigo='NGP"& ID &"'"

		'   response.write strsql &"<Br>"
		Set rs3= ofv.crearconsultaEx(StrSql,cn,1,parametros)

	End If





	Strsql = "Select * From novegenpar  where objid = " & ID

	'response.write Strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then




		sxopc1 ="<option  value=0 > SIN EXTENDER</option>"
		sxopc1 = sxopc1 & "<option  value='1' >1 DIA</option> "
		sxopc1 = sxopc1 & "<option value='2' >2 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='7' >7 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='10' >10 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='15' >15 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='20' >20 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='25' >25 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='30' >30 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='35' >35 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='40' >40 DIAS</option> "

		sxopc1 = Replace(sxopc1,"value=0 "  ,"value=0 selected")


		sxopc3 = "<option  value='1' >1 DIA</option> "
		sxopc3 = sxopc3 & "<option value='2' >2 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='7' >7 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='10' >10 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='15' >15 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='20' >20 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='25' >25 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='30' >30 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='35' >35 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='40' >40 DIAS</option> "



		sxopc3 = Replace(sxopc3,"value='" & rs(3) &"'","value='" & rs(3) & "' selected")



		sxopc4 = "<option value=0 >ALTA</option> "
		sxopc4 = sxopc4 & "<option value=1 >NORMAL</option> "
		sxopc4 = sxopc4 & "<option value=2 >BAJA</option> "

		sxopc4 = Replace(sxopc4,"value=" & rs(5),"value=" & rs(5) & " selected")

		sxopc5 = "<option value=1 >TODOS</option> "
		sxopc5 = sxopc5 & "<option value=2 >LISTA</option> "


		sxopc5 = Replace(sxopc5,"value=" & rs(6),"value=" & rs(6) & " selected")

		VDOPCION = RS(12) & ""

		If VDOPCION = "" Then VDOPCION = "0"


		sxopcVD = "<option value=0 >NO</option> "
		sxopcVD = sxopcVD & "<option value=1 >IMAGEN</option> "
		sxopcVD = sxopcVD & "<option value=2 >VIDEO</option> "

		sxopcVD = Replace(sxopcVD,"value=" & VDOPCION,"value=" & VDOPCION & " selected")


		fila = ""
		sForm = "formu"

		'  selacc = " onchange=" & chr(34) & "document." &  sForm & ".submit(); return true;" & chr(34)


		Celda =  ofv.GenCelda("","tt  ","","","","","","Codigo: ","si")
		sInput = "<b>NGP" & rs(0) & "</b>"
		Celda = celda & ofv.GenCelda("","cc colspan=5 ","","70%","","","",sInput,"si")

		Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



		Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")


		XFNAME   = request.querystring("XFNAME")
		XFSRD    = request.querystring("XFSRD")



		If XFNAME = "" Then XFNAME = "PV"
		If XFSRD  = "" Then XFSRD = 0


		' PRIMERA VEZ
		If XFNAME="PV"  Then
			XFSRD = CLng(rs(1))

		End If

		If XFNAME = "FD" Then
			XFSRD = request.form("DESDE")

		End If

		'response.write XFNAME


		sAccion = "AdNovedadesFec.ASP?XFNAME=NO&id=" & id & "&CantRegMover=" & CantRegMoverx & "&XFSRD=0" & rs(1)

		sinputE =""' "&nbsp;&nbsp;<a  class=UT href=" & sAccion

		If  CLng(XFSRD) = 0 Then
			sInput2 = Fechador(rs(1)+ CLng(CDate("01-01-1900")) )

		Else
			sInput2 = CDate(CLng(XFSRD) )
			swd = DatePart("w",sInput2)
			sInput2 = Fechador(XFSRD + CLng(CDate("01-01-1900")))

		End If

		'"frmCalendarioD","Desde"

		If  nEditar <> 1 Then

			sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

			Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")

			sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0167;</FONT></B>"
			sinputE = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & Chr(34)
			sinputE = sinputE  & "underline" & Chr(34) & ";this.style.backgroundColor=" & Chr(34)
			sinputE = sinputE  & "#f7aa22" & Chr(34) & ";' onmouseout='this.style.backgroundColor=" & Chr(34)
			sinputE = sinputE  & "" & Chr(34) & ";this.style.textDecoration=" & Chr(34)
			sinputE = sinputE  & "none" & Chr(34) & ";' "
			sinputE = sinputE  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(1).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
			sinputE = sinputE  & "  TITLE='Modificar fecha de vigencia.'  ><b>&nbsp;" & sImagenI4 & " " & "&nbsp;</b></a>"
          end if

			sInput = ofv.GenerarInput(rs(1).name,rs(1).value  + CLng(CDate("01-01-1900")) ,"hidden",sForm,"10","dt","")
			Celda = celda & ofv.GenCelda("","cc ","","25%","","","",sInput2 & sinputE,"si")

			' if  nEditar<>1 then

			' 	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

			' End if

			' Fecha de Expiracion

			If  nEditar=1 Then

				Celda = celda & ofv.GenCelda("","tt  ","","","","","","Fecha de Expiracion: ","si")
				sInput = fechador(rs(2)  + CLng(CDate("01-01-1900")))
				Celda = celda & ofv.GenCelda("","cc  ","","15%","","","",sInput,"si")

				' Dias de Vigencia Extender
				' Celda de Dias de Vigencia
				Celda = celda & ofv.GenCelda("","tt  ","","","","","","Extender Vigencia: ","si")
				sInput = "<select  class=at name=diasextender >" & sxopc1 & "</select>"


			Else

				' Dias de Vigencia
				' Celda de Dias de Vigencia
				Celda = celda & ofv.GenCelda("","tt  ","","","","","","Dias de Vigencia: ","si")
				sInput = "<select  class=at name=" & rs(3).name  & selacc &  " >" & sxopc3 & "</select>"

			End If








			Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")
			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


			sInput = ofv.GenerarInput(rs(9).name,"2","hidden",sForm,"1","dt","")

			Fila = fila & ofv.genrow("","","","","","","","<br>" & sInput,"si")



			Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","titulo","si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


			sInput = "<textarea name=" & rs(10).name  & " onkeypress=" & Chr(34)
			sInput = sInput & " if (this.value.length > 190) {alert('El texto es demasiado"
			sInput = sInput & " extenso'); return false;}" & Chr(34)
			''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
			sInput = sInput & " cols=120 rows=2>" & Replace(rs(10) & "","''","'") & "</textarea>"

			Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


			Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","copete","si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


			sInput = "<textarea name=" & rs(11).name  & " onkeypress=" & Chr(34)
			sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
			sInput = sInput & " extenso'); return false;}" & Chr(34)
			''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
			sInput = sInput & " cols=120 rows=3>" & Replace(rs(11) & "","''","'") & "</textarea>"

			Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




			Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","Mensaje","si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


			sInput = "<textarea name=" & rs(4).name  & " onkeypress=" & Chr(34)
			sInput = sInput & " if (this.value.length > 950) {alert('El texto es demasiado"
			sInput = sInput & " extenso'); return false;}" & Chr(34)
			''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
			sInput = sInput & " cols=120 rows=6>" & Replace(rs(4) & "","''","'") & "</textarea>"

			Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

			Fila = fila & ofv.genrow("","","","","","","","<br>","si")





			Celda =  ofv.GenCelda("","tt  ","","","","","","IMAGENES/VIDEO: ","si")

			If  nEditar<>1  Then

				sInput = "<select class=at name=" & rs(12).name  & selacc &  " >" & sxopcVD & "</select>"
				Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")

			Else

				Select Case rs(12)

					Case "0"  Sinput="NO"

					Case "1"  Sinput="IMAGEN"

					Case "2"  Sinput="VIDEO"


				End Select

				Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")


			End If

			' Creacion de Celda Boton Seleccionar Imagenes


			If (RS(12) & "") = "1" And nEditar <> 1 Then


				sInput = ofv.GenerarInput("","SELECCIONAR","BUTTON","","10","btSB","document.Botonx.submit();")
				Celda = celda & ofv.GenCelda("","cc colspan=6 ","","","","","",sInput,"si")

				sAccion = "Adnovedadesimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id
				FormuX = ofv.GenForm("Botonx","",sAccion,"","post","","&nbsp;","si")


			Else


				Celda = celda &  ofv.GenCelda("","cc colspan=6  ","","","","","","&nbsp;","si")

			End If




			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



			If (RS(12) & "") = "2" Then

				Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","URL DEL VIDEO:","si")
				Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


				sInput = "<textarea name=" & rs(14).name  & " onkeypress=" & Chr(34)
				sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
				sInput = sInput & " extenso'); return false;}" & Chr(34)
				sInput = sInput & " cols=120 rows=3>" & Replace(rs(14) & "","''","'") & "</textarea>"
				Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")
				Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

			End If





			' Creacion de Celda de Importancia

			Celda =  ofv.GenCelda("","tt  ","","","","","","Importancia: ","si")
			If  nEditar<>1 Then

				sInput = "<select class=at name=" & rs(5).name  & selacc &  " >" & sxopc4 & "</select>"
				Celda = celda & ofv.GenCelda("","cc","","35%","","","",sInput,"si")

			Else

				Select Case rs(5)

					Case "0"  Sinput="ALTA"

					Case "1"  Sinput="NORMAL"

					Case "2"  Sinput="BAJA"

				End Select


				Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")
			End If





			' Celda de Tipo de Distribucion

			Celda = celda &  ofv.GenCelda("","tt  ","","","","","","Tipo de Distribucion: ","si")

			If  nEditar<>1 Then

				sInput = "<select class=at name=" & rs(6).name  & selacc &  " >" & sxopc5 & "</select>"
				Celda = celda & ofv.GenCelda("","cc colspan=6 ","","35%","","","",sInput,"si")

			Else

				Select Case rs(6)

					Case "1"  Sinput="TODOS"

					Case "2"  Sinput="LISTA"

				End Select

				Celda = celda & ofv.GenCelda("","cc  colspan=6 ","","35%","","","",sInput,"si")

			End If

			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




			' Creacion de Celda de Contacto
			Celda =  ofv.GenCelda("","tt  ","","","","","","Contacto: ","si")
			sInput = "<b>" & rs(7) & "</b>"
			Celda = celda & ofv.GenCelda("","cc colspan=6 ","","70%","","","",sInput,"si")
			Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


Strsql = "Select * From novegenpar where MSGTIPOMESS = '1' Order By enviado, objid desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


			sAccion = ""
			If bModificar and rs(8)= "N" Then

				sInput = ofv.GenerarInput("","Actualizar","submit","","10","bt","")
				Celda = ofv.GenCelda("","aaa colspan=4","","10%","","","",sInput,"si")
				Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

				If nEditar <> 1 Then
					sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
					sAccion = sAccion & "&tabla=novegenpar&volver=" & vuelta & "&Editar="&nEditar
				Else
					sAccion = "adnovedadesprop.asp?Editar=1"& "&ID=" & ID
				End If


			End If

			tabla=ofv.GenTabla("","","","80% ","","0","","0","1","","",fila,"si")
			Formulario =  ofv.genform(sForm,"",sAccion,"","post","",tabla,"si")



			posi = posi + 1







		STARGET = ""

		sInputDesde="<input type=text name='" & rs(1).name & "' id='Desde' value='" & XFSRD & "' qstyle='visibility:hidden;' > "
		sAccionRetornoD= "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
		sAccionRetornoD = sAccionRetornoD & "&tabla=novegenpar&volver=" & vuelta & "&Editar="&nEditar

		' 	 formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")

		sHTML = sHTML & "<center><br>" & Formulario & "</center>" & FormuX & formRetornoD






		Call ofv.cerrarconsulta(rs)
		Call ofv.cerrarconn(cn)

		Agregar = ""
		buscar = ""

		sAccion = "Adnovedades.asp?CantRegMover=" & CantRegMoverx
		Celda = ofv.BotonVolver(sAccion)

		sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)




		sHTML = sHTML & GTCCALENDAR("")


		sHTML = sHTML & "</body></html>"





	End If
End If
Response.Write sHTML

Function Fechador(wkfecha)


	wkfecha=CLng(wkfecha)
	fec = CDate(wkfecha )
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))

	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>