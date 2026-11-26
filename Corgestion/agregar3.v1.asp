<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = headhtml()
html = html & "<body>"

tipo = request.querystring("tipo")
opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
ruta=server.mappath(request.servervariables("script_name"))


Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

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

		ExecNonQuery Cnn, "sp_cg_Registro_Insertar @Entidad='apl', @Nombre='" & request.form("nom") & "', @IdCliente=" & request.form("cli") & ", @IdAplic=" & request.form("apli") & ", @ResponEnvio=" & SqlStr(request.form("re")) & ", @MotivoApl='" & request.form("mot") & "', @FechaIngresoApl=" & sFec & ", @Sec=" & request.form("sec") & ", @TipoApl=" & request.form("tipo")

		xid = dameUltId("A",cnn)

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

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

html = html & "<div>"

html = html & cabecera()

If tipo = "APL" Then

	'html = html & "<div style='height:8%;width:75%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;margin:0 auto;'>"
	html = html & "<table class='titulo-seccion ui-page-header' >"
	html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"
	html = html & "<td><h2>NUEVO " & tit & "</h2></td>"

	formx5 = "<form id='formx5' xstyle='height:70%;' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

	botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"

	html = html & "<td>" & botonx5 & "</td>"
	html = html & "</tr>"
	html = html & "</table>"
	html = html & "</div>"

	'Formulario de ingreso de datos

	'Estilos
	'st0 = "font-family:verdana;font-size:9pt;width:250px;"
	'st1 = "width:180px;height:25px;border-radius:2.5px;line-height:25px;font-family:calibri;border-color:#cecece;"
	'st2 = "width:180px;height:200px;border-radius:2.5px;font-family:calibri;border-color:#cecece;padding:3px;width:250px;"

	' --- INICIO NORMALIZACION: FORMULARIO NUEVO APL ---
	'html = html & "<div style='height:73%;width:80%;margin:0 auto;'><br>"
	html = html & "<div style='max-width: 900px; margin: 0 auto; padding: 2rem;'>"

	html = html & "<form id='agre' method='post' action='agregar3.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"

	'html = html & "<table class='tabl' style='width:50%;' align=center cellspacing=0 cellpadding=3>"
	html = html & "<div class='bloque-datos ui-panel'>"
	
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"
	
	html = html & "<div class='bloque-body ui-panel__body'>"

	' --- INICIO GRILLA 2 COLUMNAS ---
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	' --- NOMBRE ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='" & st0 & "'  >NOMBRE:</td>"
	'html = html & "<td class='fil' '><input class='inpu' style='" & st1 & "' id='nom' name='nom' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>NOMBRE</div>"
	html = html & "<div><input id='nom' name='nom' type='text' class='input-moderno ui-input'></div>"

	' --- CLIENTE ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'  style='" & st0 & "'  >CLIENTE:</td>"
	'html = html & "<td class='fil'><select class='selec' style='" & st1 & "' id='cli' name='cli'>"
	html = html & "<div class='dato-label'>CLIENTE</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='cli' name='cli'>"

		Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
		Do While Not rs.eof
			html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			rs.movenext
		Loop

	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "</select></div>"

	' --- APLICACION ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='" & st0 & "'  >APLICACION:</td>"
	'html = html & "<td class='fil'><select class='selec' style='" & st1 & "'  id='apli' name='apli'>"
	html = html & "<div class='dato-label'>APLICACION</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='apli' name='apli'>"

		Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
		Do While Not rs.eof
			html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			rs.movenext
		Loop

	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "</select></div>"

	' --- RESPONSABLE DE ENVIO ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='" & st0 & "'  >RESPONSABLE DE ENVIO:</td>"
	'html = html & "<td class='fil'><select class='selec' style='" & st1 & "' id='re' name='re'>"
	html = html & "<div class='dato-label'>RESPONSABLE DE ENVIO</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='re' name='re'>"

		Set rs = Exec(Cnn, "sp_cg_Responsables @Modo='listar_funcion', @Funcion='FUNC'")
		Do While Not rs.eof
			html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			rs.movenext
		Loop

	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "</select></div>"

	' --- SECUENCIA ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='" & st0 & "'  >SECUENCIA:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='" & st1 & "'  id='sec' name='sec' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>SECUENCIA</div>"
	html = html & "<div><input id='sec' name='sec' type='text' class='input-moderno ui-input'></div>"

	' --- TIPO ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='" & st0 & "'  > TIPO:</td>"
	'html = html & "<td class='fil'><select class='selec' style='" & st1 & "' id='tipo' name='tipo'>"
	html = html & "<div class='dato-label'>TIPO</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='tipo' name='tipo'>"

		Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='tipo_apl'")
		html = html & "<option value=Seleccione selected>--- Seleccione Tipo ---</option>"
		Do While Not rs.eof
			html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			rs.movenext
		Loop

	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "</select></div>"

	' --- FIN GRILLA 2 COLUMNAS ---
	html = html & "</div>"

	' --- CAMPO FULL WIDTH: MODULOS IMPLICADOS ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='" & st0 & "'>MODIFIC Y MODULOS IMPLICADOS:</td>"
	'html = html & "<td class='fil'><textarea class='txtar' style='" & st2 & "'  id='mot' name='mot' ></textarea></td>"
	'html = html & "</tr>"
	html = html & "<div style='margin-top: 2rem;'>"
	html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>MODIFIC Y MODULOS IMPLICADOS</div>"
	html = html & "<div><textarea id='mot' name='mot' class='input-moderno ui-input' style='height: 150px; resize: vertical;'></textarea></div>"
	html = html & "</div>"

	' --- BOTON AGREGAR ---
	'html = html & "<tr>"
	'html = html & "<td colspan=4 align=center><input class='bot' type='button' value='Agregar' onclick='agregar();' ></td>"
	'html = html & "</tr>"
	html = html & "<div style='text-align: center; margin-bottom: 1rem;margin-top:1rem;'>"
	html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar();'>"
	html = html & "</div>"

	' --- CIERRES FINALES ---
	'html = html & "</table>"
	html = html & "</div>" ' Fin .bloque-body
	html = html & "</div>" ' Fin .bloque-datos

	html = html & "</form>"
	html = html & "</div>" ' Fin max-width

End If


'html = html & "<script src='assets/js/jquery.min.js'></script>"
'html = html & "<script src='assets/js/jquery-ui.js'></script>"
'html = html & "<script src='assets/js/skel.min.js'></script>"
'html = html & "<script src='assets/js/util.js'></script>"
'html = html & "<script src='assets/js/main.js'></script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar(){ "
html = html & " var nom = document.getElementById('nom').value;"
html = html & " if (nom == ''){ "
html = html & " alert('Nombre no debe ser nulo'); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"

html = html & "<script>"
html = html & "  $( function() {"

html = html & "$.datepicker.regional['es'] = {"
html = html & "closeText: 'Cerrar',"
html = html & "prevText: '< Ant',"
html = html & "nextText: 'Sig >',"
html = html & "currentText: 'Hoy',"
html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi�','Juv','Vie','S�b'],"
html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
html = html & "weekHeader: 'Sm',"
html = html & "dateFormat: 'dd/mm/yy',"
html = html & "firstDay: 1,"
html = html & "isRTL: false,"
html = html & "showMonthAfterYear: false,"
html = html & "yearSuffix: ''"
html = html & "};"
html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

html = html & "    $( ""#fe"" ).datepicker();"
html = html & "    $( ""#fi"" ).datepicker();"
html = html & "  } );"
html = html & "  </script>"



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