<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If


html = headhtml()
html = html & "<body>"

tipo = request.querystring("tipo")
opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
ruta=server.mappath(request.servervariables("script_name"))


Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] Cnn.open "[credenciales eliminadas]"
' [UI] Conexion con la configuracion del sistema (scon1), como el resto de las pantallas.
' [UI] Antes se conectaba con el usuario administrador de SQL Server y su contrasena escritos en el codigo.
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

If opc = "ins" Then

	tabla = request.querystring("tabla")

	If tabla = "APL" Then

		nFecini = 0

		If nFecini < CLng(CDate("01-01-1900")) Then
			nFecini = CLng(CDate("01-01-1900"))
		Else
			nFecini = 0
		End If

		sFec = CLng(Date) - nFecini

		'[UI-ANT] ExecNonQuery Cnn, "sp_cg_Registro_Insertar @Entidad='apl', @Nombre='" & request.form("nom") & "', @IdCliente=" & request.form("cli") & ", @IdAplic=" & request.form("apli") & ", @ResponEnvio=" & SqlStr(request.form("re")) & ", @MotivoApl='" & request.form("mot") & "', @FechaIngresoApl=" & sFec & ", @Sec=" & request.form("sec") & ", @TipoApl=" & request.form("tipo")
		' [UI] Valores escapados (SqlStr) o validados como numero: antes se pegaba el texto del formulario en la consulta.
		' [UI] Un apostrofo en el nombre rompia el alta, y un texto malicioso se ejecutaba en la base.
		Dim aplSec
		aplSec = "NULL"
		If IsNumeric(request.form("sec") & "") Then aplSec = CLng(request.form("sec"))
		ExecNonQuery Cnn, "sp_cg_Registro_Insertar @Entidad='apl', @Nombre=" & SqlStr(request.form("nom")) & ", @IdCliente=" & CLng(request.form("cli")) & ", @IdAplic=" & CLng(request.form("apli")) & ", @ResponEnvio=" & SqlStr(request.form("re")) & ", @MotivoApl=" & SqlStr(request.form("mot")) & ", @FechaIngresoApl=" & sFec & ", @Sec=" & aplSec & ", @TipoApl=" & CLng(request.form("tipo"))

		xid = dameUltId("A",cnn)
		CgLog "APL", xid, "ALTA"

		vvv = ObjAcc(xid,"A",cnn,"C")

		'Set fso = SERVER.CreateObject("Scripting.FileSystemObject")
	    'If fso.FolderExists(ruta  & folderx) Then


	End If

End If


If opc <> "ins" Then

	Select Case tipo

		Case "APL"

		tabla = "APL"
		tit = "APL"
		volver = "apl.asp"

	End Select

End If

' [UI] Reconexion innecesaria (la conexion ya esta abierta) con credenciales en el codigo: desactivada
'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] Cnn.open "[credenciales eliminadas]"

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & cabecera()

'[UI-ANT] If tipo = "APL" Then

	'[UI-ANT] 'html = html & "<div style='height:8%;width:75%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;margin:0 auto;'>"
	'[UI-ANT] html = html & "<table class='titulo-seccion ui-page-header' >"
	'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"
	'[UI-ANT] html = html & "<td><h2>NUEVO " & tit & "</h2></td>"

	'[UI-ANT] formx5 = "<form id='formx5' xstyle='height:70%;' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

	'[UI-ANT] botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"

	'[UI-ANT] html = html & "<td>" & botonx5 & "</td>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "</table>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'Formulario de ingreso de datos

	'[UI-ANT] 'Estilos
	'[UI-ANT] 'st0 = "font-family:verdana;font-size:9pt;width:250px;"
	'[UI-ANT] 'st1 = "width:180px;height:25px;border-radius:2.5px;line-height:25px;font-family:calibri;border-color:#cecece;"
	'[UI-ANT] 'st2 = "width:180px;height:200px;border-radius:2.5px;font-family:calibri;border-color:#cecece;padding:3px;width:250px;"

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO NUEVO APL ---
	'[UI-ANT] 'html = html & "<div style='height:73%;width:80%;margin:0 auto;'><br>"
	'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 0 auto; padding: 2rem;'>"

	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar3.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:50%;' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"
	
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"
	
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

	'[UI-ANT] ' --- INICIO GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] ' --- NOMBRE ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st0 & "'  >NOMBRE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil' '><input class='inpu' style='" & st1 & "' id='nom' name='nom' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE</div>"
	'[UI-ANT] html = html & "<div><input id='nom' name='nom' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] ' --- CLIENTE ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'  style='" & st0 & "'  >CLIENTE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='" & st1 & "' id='cli' name='cli'>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='cli' name='cli'>"

		'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- APLICACION ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st0 & "'  >APLICACION:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='" & st1 & "'  id='apli' name='apli'>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='apli' name='apli'>"

		'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- RESPONSABLE DE ENVIO ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st0 & "'  >RESPONSABLE DE ENVIO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='" & st1 & "' id='re' name='re'>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESPONSABLE DE ENVIO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='re' name='re'>"

		'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Responsables @Modo='listar_funcion', @Funcion='FUNC'")
		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- SECUENCIA ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st0 & "'  >SECUENCIA:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='" & st1 & "'  id='sec' name='sec' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>SECUENCIA</div>"
	'[UI-ANT] html = html & "<div><input id='sec' name='sec' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] ' --- TIPO ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st0 & "'  > TIPO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='" & st1 & "' id='tipo' name='tipo'>"
	'[UI-ANT] html = html & "<div class='dato-label'>TIPO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='tipo' name='tipo'>"

		'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='tipo_apl'")
		'[UI-ANT] html = html & "<option value=Seleccione selected>--- Seleccione Tipo ---</option>"
		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- FIN GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- CAMPO FULL WIDTH: MODULOS IMPLICADOS ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>MODIFIC Y MODULOS IMPLICADOS:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><textarea class='txtar' style='" & st2 & "'  id='mot' name='mot' ></textarea></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div style='margin-top: 2rem;'>"
	'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>MODIFIC Y MODULOS IMPLICADOS</div>"
	'[UI-ANT] html = html & "<div><textarea id='mot' name='mot' class='input-moderno ui-input' style='height: 150px; resize: vertical;'></textarea></div>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- BOTON AGREGAR ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4 align=center><input class='bot' type='button' value='Agregar' onclick='agregar();' ></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 1rem;margin-top:1rem;'>"
	'[UI-ANT] html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar();'>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- CIERRES FINALES ---
	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] html = html & "</form>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] End If


'[UI-ANT] 'html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/jquery-ui.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function agregar(){ "
'[UI-ANT] html = html & " var nom = document.getElementById('nom').value;"
'[UI-ANT] html = html & " if (nom == ''){ "
'[UI-ANT] html = html & " alert('Nombre no debe ser nulo'); return false; "
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & " document.getElementById('agre').submit();"
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script>"
'[UI-ANT] html = html & "  $( function() {"

'[UI-ANT] html = html & "$.datepicker.regional['es'] = {"
'[UI-ANT] html = html & "closeText: 'Cerrar',"
'[UI-ANT] html = html & "prevText: '< Ant',"
'[UI-ANT] html = html & "nextText: 'Sig >',"
'[UI-ANT] html = html & "currentText: 'Hoy',"
'[UI-ANT] html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
'[UI-ANT] html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
'[UI-ANT] html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
'[UI-ANT] html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi?','Juv','Vie','S?b'],"
'[UI-ANT] html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
'[UI-ANT] html = html & "weekHeader: 'Sm',"
'[UI-ANT] html = html & "dateFormat: 'dd/mm/yy',"
'[UI-ANT] html = html & "firstDay: 1,"
'[UI-ANT] html = html & "isRTL: false,"
'[UI-ANT] html = html & "showMonthAfterYear: false,"
'[UI-ANT] html = html & "yearSuffix: ''"
'[UI-ANT] html = html & "};"
'[UI-ANT] html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

'[UI-ANT] html = html & "    $( ""#fe"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fi"" ).datepicker();"
'[UI-ANT] html = html & "  } );"
'[UI-ANT] html = html & "  </script>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Ids y nombres de los campos, formulario y validaciones: los mismos de antes.
' =========================================================================
Dim accTop

html = html & "<div>"
html = html & cabecera()
html = html & "</div>"

If tipo = "APL" Then

	formx5 = "<form id='formx5' class='form-oculto' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	accTop = UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")
	html = html & UiPantallaAbre("Nuevo APL", accTop, "")

	html = html & "<form id='agre' method='post' action='agregar3.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	html = html & UiCardAbre("Datos del APL", "") & UiNotaObligatorios() & UiFormGridAbre()

	html = html & UiCampo("nom", "Nombre", UiInputTexto("nom", True, "autocomplete='off'"), True, True)
	html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'"), "", "", "", ""), False, False)
	html = html & UiCampo("apli", "Aplicaci&oacute;n", UiSelectRs("apli", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'"), "", "", "", ""), False, False)
	html = html & UiCampo("re", "Responsable de env&iacute;o", UiSelectRs("re", Exec(Cnn, "sp_cg_Responsables @Modo='listar_funcion', @Funcion='FUNC'"), "", "", "", ""), False, False)
	html = html & UiCampo("sec", "Secuencia", UiInputTexto("sec", False, "inputmode='numeric' autocomplete='off'"), False, False)
	' El valor "Seleccione" de la primera opcion es el de antes; la validacion ahora exige elegir un tipo
	html = html & UiCampo("tipo", "Tipo", UiSelectRs("tipo", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='tipo_apl'"), "", "Seleccionar tipo", "Seleccione", "aria-required='true'"), True, False)
	html = html & UiCampo("mot", "Modificaciones y m&oacute;dulos implicados", UiTextarea("mot", 7, False), False, True)

	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonPrincipal("Agregar APL", "agregar();"))
	html = html & "</form>"
	html = html & UiPantallaCierra()

End If

html = html & "<script type='text/javascript'>"
html = html & " function agregar(){ "
html = html & " var nom = document.getElementById('nom');"
html = html & " var tip = document.getElementById('tipo');"
html = html & " if (nom.value == ''){ "
html = html & " alert('Complet\u00e1 el nombre.'); nom.focus(); return false; "
html = html & " } "
' Sin tipo, el alta fallaba en la base (se enviaba el texto "Seleccione" como numero)
html = html & " if (tip.value == 'Seleccione'){ "
html = html & " alert('Eleg\u00ed el tipo de APL.'); tip.focus(); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"



If opc = "ins" Then

	vuelta = request.querystring("volver")

	html = html & "<FORM name='init' method='post'  ACTION='" & vuelta & "?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
	html = html & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
	html = html & "</FORM>"

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Agregado exitosamente');"
	html = html & " document.all.init.submit();"
	html = html & "</script>"

End If

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html
%>