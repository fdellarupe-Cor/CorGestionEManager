<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  'Encabezados
  formulario =  ofv.Encabezados("Nivel","Descripcion","CODIGO EXTERNO")

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  set cn = ofv.conectar(ofv.strconn0)



  ofv.ObtenerAtributos session("Form"),""
'call ObtenerAtributos( session("Form"),"",cn)
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
'response.write bModificar & "  "   & Session("CorApli")
  StrSql = "Select * From NivelConfiden order by NCCod"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  do until rs.eof
    sInput = ofv.GenerarInput(rs(0).name,rs(0),"hidden","niveles" & rs(0),"10","dt","")
    celdas = ofv.GenCelda("","tt colspan=2","","","","","",sinput & rs(0),"si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","niveles" & rs(0),"90","dt","")
      sInput2 = ofv.GenerarInput(rs(3).name,rs(3),"text","niveles" & rs(0),"20","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
      sInput2 = ofv.GenerarInput("",rs(3),"readonly","","20","dt","")

    end if

    celdas = celdas & ofv.GenCelda("","","","","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sInput2,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?ID=" & rs(0) & "&volver=" & vuelta & "&Tabla=NivelConfiden"
      formulario = formulario & ofv.GenForm("niveles" & rs(0),"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    rs.movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",formulario,"si")

  'Botonera
   Agregar = ""
   Buscar = ""
   Retornar = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


Public Sub ObtenerAtributos(sForm,  sNodo,cn)

		Call ObtenerAtributos2(sForm,  sNodo,  "",  "",cn)
	End Sub


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub ObtenerAtributos2(sForm,  sNodo,  sCampo,  sValor,cn)

		'On Error GoTo ObtenerAtributosErr

		'Dim strsql, atrok, cn, rs, sSQL, xcount , XMFI , XMFSEP, xsql
		Dim XMFTGRP
		' Dim a, b, c, d, e
		AEstado = True
		AConsultar = True
		AEliminar = True
		AModificar = True
		AAgregar = True

		Dim XTabPer(4)

		xcount = 0


	'	Set cn = ofv.Conectar(strconn0)

		atrok = "N"
		strsql = "Select Alias,usrEstado,usragregar,usrmodificar,usreliminar,usrconsultar "
		strsql = strsql & " From atrmXusuario Where "
		If sNodo <> "" Then
			strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
		Else
			strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
		End If
		strsql = strsql & " And usuario = '" & Session("segperfil") & "' And "
		strsql = strsql & "  Aplicacion = '" & Session("CorApli") & "'"
		'  xsql = strsql & Chr(13)
		response.write strsql
		Set rs = ofv.CrearConsultaEx(strsql, cn, 1, "")
		If rs.EOF Then
			atrok = "O"
		Else
			atrok = "S"
			AEstado = rs(1)
			AConsultar = rs(5)
			AEliminar = rs(4)
			AModificar = rs(3)
			AAgregar = rs(2)
		End If
		Call ofv.CerrarConsulta(rs)

		If atrok = "O" Then
			strsql = "Select Alias,grpEstado,grpagregar,grpmodificar,grpeliminar,grpconsultar "
			strsql = strsql & " From atrmXgrupo Where  "
			If sNodo <> "" Then
				strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
			Else
				strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
			End If
			strsql = strsql & " And Aplicacion = '" & Session("CorApli") & "'"

			If Session("DELLOG") = "S" Then

				XMFTGRP = Split(Session("SegGrpLst"), Chr(9))
				xcount = UBound(XMFTGRP) + 1
				XMFSEP = ""

				For XMFI = 0 To 4
					XTabPer(XMFI) = xcount
				Next

				strsql = strsql & " and grupo in ("

				For XMFI = 0 To UBound(XMFTGRP)
					strsql = strsql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
					XMFSEP = ","
				Next

				strsql = strsql & ") "
				' xsql = xsql & strsql & Chr(13)
				response.write "<br>" & strsql
				Set rs = ofv.CrearConsultaEx(strsql, cn, 1, "")
				If rs.EOF Then
					atrok = "O"
				Else
					atrok = "S"
					Do While Not rs.EOF


						If Not rs(5) Then XTabPer(1) = XTabPer(1) - 1 ''' consultar
						If Not rs(4) Then XTabPer(2) = XTabPer(2) - 1 ''' eliminar
						If Not rs(3) Then XTabPer(3) = XTabPer(3) - 1 ''' modificar
						If Not rs(2) Then XTabPer(4) = XTabPer(4) - 1 ''' agregar

						rs.movenext
					Loop

				End If

				If atrok = "S" Then


					If XTabPer(1) = 0 Or XTabPer(1) = False Then AConsultar = False
					If XTabPer(2) = 0 Or XTabPer(2) = False Then AEliminar = False
					If XTabPer(3) = 0 Or XTabPer(3) = False Then AModificar = False
					If XTabPer(4) = 0 Or XTabPer(4) = False Then AAgregar = False


				End If


			Else

				strsql = strsql & " And grupo = '" & Session("seggrp") & "'  "
				' xsql = xsql & strsql & Chr(13)
				response.write "<br>" & strsql
				Set rs = ofv.CrearConsultaEx(strsql, cn, 1, "")
				If rs.EOF Then
					atrok = "O"
				Else
					atrok = "S"

					AEstado = CBool(rs(1))
					AConsultar = CBool(rs(5))
					AEliminar = CBool(rs(4))
					AModificar = CBool(rs(3))
					AAgregar = CBool(rs(2))
				End If

			End If

			Call ofv.CerrarConsulta(rs)
		End If
       ' Call ofv.Cerrarconn(cn)

		'Response.WRITE sSQL



	End Sub

response.write sHTML

%>