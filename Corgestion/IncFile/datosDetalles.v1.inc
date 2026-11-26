<%
'==============================================================================
' IncFile/datosDetalles.inc
' Capa de ACCESO A DATOS especifica de detalles.asp. Cada funcion ejecuta
' un Stored Procedure y devuelve el mismo Recordset que antes devolvia el
' SQL crudo (mismo orden/alias de columnas esperado por la vista), para que
' detalles.asp NO tenga que cambiar la forma en que lee rs(0), rs(1), etc.
'
'==============================================================================

Function EjecutarSP_ADO(cn, spName, paramNames, paramValues)
	Dim cmd, i
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = spName
	For i = 0 To UBound(paramNames)
		If IsNull(paramValues(i)) Then
			cmd.Parameters.Append cmd.CreateParameter(paramNames(i), 200, 1, 200, Null)
		Else
			cmd.Parameters.Append cmd.CreateParameter(paramNames(i), 200, 1, 200, CStr(paramValues(i)))
		End If
	Next
	On Error Resume Next
	Set EjecutarSP_ADO = cmd.Execute()
	If Err.Number <> 0 Then
		Response.Write "ERROR " & Err.Number & ": " & Err.Description
		Response.End
	End If
	On Error Goto 0
	Set cmd = Nothing
End Function

'------------------------------------------------------------------------
' 1. Detalle de Incidente (Modo Lectura)  ->  sp_cg_ObtenerDetalleIncidente
'------------------------------------------------------------------------
Function ObtenerDetalleIncidente(cn, id)
	Set ObtenerDetalleIncidente = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerDetalleIncidente", Array("@id"), Array(id))
End Function

'------------------------------------------------------------------------
' 2. APLs Relacionados por Entidad  ->  sp_cg_ObtenerAplRelac
'    tipoRelac: "I" (incidente) o "P" (proyecto), igual que el filtro
'    a.tipo_relac original.
'------------------------------------------------------------------------
Function ObtenerAplRelac(cn, id, tipoRelac)
	Set ObtenerAplRelac = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerAplRelac", Array("@id_relac", "@id_apl", "@tipoRelac"), Array(id, Null, tipoRelac))
End Function

'------------------------------------------------------------------------
' 2b. Entidades (incidente/proyecto) relacionadas a un APL  ->  sp_ObtenerRelacionesApl
'     Direccion inversa de ObtenerAplRelac: recibe el ID del APL y devuelve
'     los incidentes/proyectos que lo referencian. Usada en detalles2.asp
'     (pestaña "INCIDENTES O PROYECTOS RELACIONADOS").
'------------------------------------------------------------------------
Function ObtenerRelacionesApl(cn, id, tipoRelac)
	Set ObtenerRelacionesApl = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerAplRelac", Array("@id_relac", "@id_apl", "@tipo_relac"), Array(Null, id, tipoRelac))
End Function


'------------------------------------------------------------------------
' 3. Detalle de Proyecto (Modo Lectura)  ->  sp_cg_ObtenerDetalleProyecto
'------------------------------------------------------------------------
Function ObtenerDetalleProyecto(cn, id)
	Set ObtenerDetalleProyecto = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerDetalleProyecto", Array("@id"), Array(id))
End Function

'------------------------------------------------------------------------
' 4. Historial de Comentarios  ->  sp_cg_ObtenerComentarios
'    En el original habia DOS variantes de filtro segun la pestaña:
'      - modo "inc": WHERE c.id=@id AND a.tipoobjeto='inc' AND a.codobjeto=@id
'                    (pestaña Avances del incidente)
'      - modo "cmd": WHERE c.id=@id AND a.tipo='cmd'
'                    (pestaña Avances del proyecto -- OJO: esta consulta ya
'                    en el original hacia join contra "incidentes" para un
'                    listado de comentarios de PROYECTO, posible bug
'                    preexistente. Se preserva tal cual, sin corregir.)
'    Se pasa "modo" como parametro para que el SP decida el filtro exacto;
'    confirmar el nombre de ese parametro contra tu implementacion real.
'------------------------------------------------------------------------
Function ObtenerComentarios(cn, codobjeto, tipoobjeto)
	Set ObtenerComentarios = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerComentarios", Array("@codobjeto", "@tipoobjeto"), Array(codobjeto, tipoobjeto))
End Function

'------------------------------------------------------------------------
' 5. Incidentes Asociados a un Proyecto  ->  sp_cg_ObtenerIncidentesPorProyecto
'------------------------------------------------------------------------
Function ObtenerIncidentesPorProyecto(cn, id)
	Set ObtenerIncidentesPorProyecto = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerIncidentesPorProyecto", Array("@id"), Array(id))
End Function

'------------------------------------------------------------------------
' 6. Incidente en Modo Edicion  ->  SP SUGERIDO (no estaba en tu lista)
'    Nombre propuesto: sp_cg_ObtenerIncidenteEdicion
'    Equivale, para "incidentes", a lo que sp_cg_ObtenerProyectoEdicion hace
'    para "proyectos" (junta ids + descripciones para poblar el form de
'    edicion, no solo lectura). Crear el SP y, si el nombre final es otro,
'    ajustar unicamente la linea de abajo.
'------------------------------------------------------------------------
Function ObtenerIncidenteEdicion(cn, id)
	Set ObtenerIncidenteEdicion = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerIncidenteEdicion", Array("@id"), Array(id))
End Function

'------------------------------------------------------------------------
' 7. Proyecto en Modo Edicion  ->  sp_cg_ObtenerProyectoEdicion
'------------------------------------------------------------------------
Function ObtenerProyectoEdicion(cn, id)
	Set ObtenerProyectoEdicion = EjecutarSP_ADO(cn, "dbo.sp_cg_ObtenerProyectoEdicion", Array("@id"), Array(id))
End Function


'------------------------------------------------------------------------
' 8. Existencia de un APL (chequeo rapido, usado en detalles2.asp cuando
'    tipo="cons" antes de mostrar el detalle)
'------------------------------------------------------------------------
Function ExisteApl(cn, id)
	Dim cmd, rs
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "sp_cg_ConsultarEntidad"
	cmd.Parameters.Append cmd.CreateParameter("@Entidad", 200, 1, 15, "apl")
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "existe")
	cmd.Parameters.Append cmd.CreateParameter("@Id", 3, 1, 4, CLng(id))
	On Error Resume Next
	Set rs = cmd.Execute()
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ExisteApl (id=" & Server.HTMLEncode(id) & ")<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
	ExisteApl = Not rs.EOF
End Function

'------------------------------------------------------------------------
' 9. Detalle de APL (Modo Lectura) -> detalles2.asp, opc="ver"
'------------------------------------------------------------------------
Function ObtenerDetalleApl(cn, id)
	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "sp_cg_ConsultarEntidad"
	cmd.Parameters.Append cmd.CreateParameter("@Entidad", 200, 1, 15, "apl")
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "detalle")
	cmd.Parameters.Append cmd.CreateParameter("@Id", 3, 1, 4, CLng(id))
	On Error Resume Next
	Set ObtenerDetalleApl = cmd.Execute()
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerDetalleApl (id=" & Server.HTMLEncode(id) & ")<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function


'------------------------------------------------------------------------
' 11. Detalle de APL (Modo Edicion) -> detalles2.asp, opc="act"
'------------------------------------------------------------------------
Function ObtenerAplEdicion(cn, id)
	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "sp_cg_ConsultarEntidad"
	cmd.Parameters.Append cmd.CreateParameter("@Entidad", 200, 1, 15, "apl")
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "edicion")
	cmd.Parameters.Append cmd.CreateParameter("@Id", 3, 1, 4, CLng(id))
	On Error Resume Next
	Set ObtenerAplEdicion = cmd.Execute()
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerAplEdicion (id=" & Server.HTMLEncode(id) & ")<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

Function ObtenerContacto(cn, id)
	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "sp_cg_Contactos"
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "obtener")
	cmd.Parameters.Append cmd.CreateParameter("@Id", 3, 1, 4, CLng(id))
	On Error Resume Next
	Set ObtenerContacto = cmd.Execute()
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerContacto (id=" & Server.HTMLEncode(id) & ")<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

' ' ------------------------------------------------------------------------------
' '-- Capa de Negocio / Acceso a Datos para Proyectos
' '-- Pre-requisito: Incluir "IncFile/funciones.inc" antes que este archivo para tener 'cnn' y 'scon1'
' ' ------------------------------------------------------------------------------

' Función genérica para cargar combos (Clientes, Aplicativos)
' Retorna un Recordset desconectado
Function GetListaCombo( cn,  tabla )
	dim cmdtxt
	Dim cmd
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "sp_cg_Catalogo_Listar"

	Select Case UCase(tabla)
		Case "CLIENTES"
			cmd.CommandText = "sp_cg_Catalogo_Listar"
			cmd.Parameters.Append cmd.CreateParameter("@Catalogo", 200, 1, 30, "clientes")
		Case "APLICATIVOS"
			cmd.CommandText = "sp_cg_Catalogo_Listar"
			cmd.Parameters.Append cmd.CreateParameter("@Catalogo", 200, 1, 30, "aplicativos")
		Case "TIPO_APL"
			cmd.CommandText = "sp_cg_Catalogo_Listar"
			cmd.Parameters.Append cmd.CreateParameter("@Catalogo", 200, 1, 30, "tipo_apl")
		Case "RESPONSABLES_FUNC"
			cmd.CommandText = "sp_cg_Combo_Listar"
			cmd.Parameters.Append cmd.CreateParameter("@Origen", 200, 1, 20, "respon_gestion")
	End Select

	On Error Resume Next
	Set GetListaCombo = cmd.Execute()
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: GetListaCombo (tabla=" & Server.HTMLEncode(tabla) & ")<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

	
	
' Función principal para obtener el listado de proyectos llamando al SP
' Retorna un Recordset desconectado
Function GetProyectosListado(cn, IdCliente, IdAplicativo, TextoBusqueda, EstadoFiltro, FechaIngreso, Orden, RegistroInicial, CantidadMostrar)
    Dim cmd, rs
    
    '-- Constantes ADODB (si no están definidas en funciones.inc, definirlas aquí)
    Const adCmdStoredProc = 4
    Const adInteger = 3
    Const adVarChar = 200
    Const adDate = 7
    Const adParamInput = 1
    Const adUseClient = 3
    Const adOpenStatic = 3
    Const adLockReadOnly = 1
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = cn
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "dbo.sp_cg_Proyectos_Listar"
    
    '-- Definición estricta de parámetros para seguridad y rendimiento
    cmd.Parameters.Append cmd.CreateParameter("@IdCliente", adInteger, adParamInput, , Clng(IdCliente))
    cmd.Parameters.Append cmd.CreateParameter("@IdAplicativo", adInteger, adParamInput, , Clng(IdAplicativo))
    
    ' Manejo de Nulos para strings
    If Trim(TextoBusqueda) = "" Then
        cmd.Parameters.Append cmd.CreateParameter("@TextoBusqueda", adVarChar, adParamInput, 100, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@TextoBusqueda", adVarChar, adParamInput, 100, TextoBusqueda)
    End If
    
    cmd.Parameters.Append cmd.CreateParameter("@EstadoFiltro", adInteger, adParamInput, , Clng(EstadoFiltro))
    
    ' El SP espera @FechaIngreso como INT: dias desde 1900-01-01
    If IsDate(FechaIngreso) Then
         cmd.Parameters.Append cmd.CreateParameter("@FechaIngreso", adInteger, adParamInput, , CLng(CDate(FechaIngreso)) - CLng(CDate("01-01-1900")))
    ElseIf IsNumeric(FechaIngreso) Then
         If CLng(FechaIngreso) > 0 Then
              cmd.Parameters.Append cmd.CreateParameter("@FechaIngreso", adInteger, adParamInput, , CLng(FechaIngreso))
         Else
              cmd.Parameters.Append cmd.CreateParameter("@FechaIngreso", adInteger, adParamInput, , Null)
         End If
    Else
         cmd.Parameters.Append cmd.CreateParameter("@FechaIngreso", adInteger, adParamInput, , Null)
    End If

    cmd.Parameters.Append cmd.CreateParameter("@Orden", adVarChar, adParamInput, 50, Orden)
    cmd.Parameters.Append cmd.CreateParameter("@RegistroInicial", adInteger, adParamInput, , Clng(RegistroInicial))
    cmd.Parameters.Append cmd.CreateParameter("@CantidadMostrar", adInteger, adParamInput, , Clng(CantidadMostrar))
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = adUseClient
    On Error Resume Next
    rs.Open cmd, , adOpenStatic, adLockReadOnly
    If Err.Number <> 0 Then
        Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
        Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
        Response.Write "Contexto: GetProyectosListado<br>"
        Response.Write "Codigo: " & Err.Number & "<br>"
        Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
        Response.End
    End If
    On Error Goto 0
    
    Set rs.ActiveConnection = Nothing '-- Desconectar Recordset
    Set GetProyectosListado = rs
    Set cmd = Nothing
End Function

''Funciones auxiliares de UI (Movidas de la vista para limpiar código)
 Function GetBadgePrioridad(prioridadDesc)
    Dim classes, uiClasses, texto
    texto = UCASE(prioridadDesc & "")
    Select Case texto
        Case "ALTA"
            classes = "badge-peligro"
            uiClasses = "ui-badge ui-badge--danger"
        Case "MEDIA"
            classes = "badge-media"
            uiClasses = "ui-badge ui-badge--warning"
        Case Else
            classes = "badge-pendiente"
            uiClasses = "ui-badge ui-badge--pending"
    End Select
     GetBadgePrioridad = "<span class='badge " & classes & " " & uiClasses & "' aria-label='Prioridad " & Server.HTMLEncode(texto) & "'>" & Server.HTMLEncode(texto) & "</span>"
 End Function

Function GetBadgeEstado(estadoDesc, responNombre)
    Dim classes, uiClasses, texto
    texto = UCASE(estadoDesc & "")
    Select Case texto
        Case "PENDIENTE"
            classes = "badge-media"
            uiClasses = "ui-badge ui-badge--warning"
            texto = "PEND"
        Case "OK"
            classes = "badge-ok"
            uiClasses = "ui-badge ui-badge--ok"
        Case "BAJA"
            classes = "badge-peligro"
            uiClasses = "ui-badge ui-badge--danger"
        Case "TESTING EN CLIENTE"
            classes = "badge-media"
            uiClasses = "ui-badge ui-badge--warning"
            texto = "TEST"
        Case Else
            classes = ""
            uiClasses = "ui-badge"
    End Select

    GetBadgeEstado = "<span class='badge " & classes & " " & uiClasses & "' aria-label='" & Server.HTMLEncode(texto) & "'>" & Server.HTMLEncode(texto) & "</span>" & _
                      " (" & "&nbsp;" & Server.HTMLEncode(responNombre & "") & ")"
End Function

' Arma un <select> a partir de un Recordset ya obtenido (por ejemplo con
' GetListaCombo). Reemplaza a la vieja Function armaCombo que vivia en
' detalles2.asp (armaba el combo con SQL crudo pasado como string, y ademas
' tenia un bug de nombre -- asignaba a "armarCombo" en vez de "armaCombo",
' por eso el llamador necesitaba pegarle un "</select>" extra a mano; esta
' version devuelve el <select> ya completo y cerrado).
Function ArmarComboHTML(rsCombo, clave, comboid)
	Dim h
	h = "<select class='selec ui-select' style='width:90%;border-radius:5px;font-family:arial;font-size:8pt;height:25px;line-height:normal;border-color:#dedede;color:#7d7f7d;' id='" & comboid & "' name='" & comboid & "'>"

	Do While Not rsCombo.eof
		If clave = rsCombo(0) Then
			h = h & "<option style='font-family:calibri;font-size:9pt;color:#7d7f7d;' value='" & rsCombo(0) & "' selected>" & rsCombo(1) & "</option>"
		Else
			h = h & "<option value='" & rsCombo(0) & "'>" & rsCombo(1) & "</option>"
		End If
		rsCombo.movenext
	Loop
h = h & "</select>"
	ArmarComboHTML = h
End Function
%>
