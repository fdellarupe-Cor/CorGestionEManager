<%
'==============================================================================
' IncFile/datosAgregar.inc
' Capa de ACCESO A DATOS especifica de agregar2.asp.
'
' Parte A: los 3 altas (Incidente/Proyecto/Licencia), antes armadas con SQL
'          concatenado desde request.form (vulnerable a inyeccion SQL), ahora
'          via los SP nuevos (ver sp_agregar2.sql) con parametros tipados.
' Parte B: los combos de catalogo (clientes, aplicativos, prioridades,
'          responsables, via_pedido, browsers, contactos). Se dejo el mismo
'          SELECT que tenia el original (no se inventaron columnas), solo se
'          saco de la vista. No requieren SP porque no reciben datos del
'          usuario (no hay riesgo de inyeccion), pero si preferis unificarlos
'          con sp_cg_Catalogo_Listar (el de detalles.asp), avisame la forma
'          real de esa SP para no romper columnas (aca hay combos que leen 3
'          columnas -id, nombre, apellido- no solo 2).
'==============================================================================

Function EjecutarSP_ADO(cn, spName, paramDefs)
	' Helper interno generico: paramDefs es un array de arrays
	' {nombre, tipoADO, tamano, valor}. Devuelve el Command ya armado
	' (sin ejecutar) para reusar tanto en SELECT como en INSERT/UPDATE.
	Dim cmd, i
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = spName

	For i = 0 To UBound(paramDefs)
		cmd.Parameters.Append cmd.CreateParameter(paramDefs(i)(0), paramDefs(i)(1), 1, paramDefs(i)(2), paramDefs(i)(3))
	Next

	Set EjecutarSP_ADO = cmd
End Function

Sub EjecutarSP_NonQuery(cn, spName, paramDefs)
	Dim cmd
	Set cmd = EjecutarSP_ADO(cn, spName, paramDefs)
	On Error Resume Next
	cmd.Execute
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "SP: " & Server.HTMLEncode(spName) & "<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Sub

'------------------------------------------------------------------------
' Parte A - Altas (INSERT via SP)
' Tipos ADO usados: 3 = adInteger, 5 = adDouble, 200 = adVarChar
'------------------------------------------------------------------------

Sub InsertarIncidente(cn, idCliente, contacto, idAplic, detalles, prioridad, _
						responGestion, responTec, reqPresup, fechaIngreso, browser, viaPedido, titulo, responsableRol)

	Dim p(12)
	p(0)  = Array("@id_cliente", 3, 4, CLng(idCliente))
	p(1)  = Array("@contacto", 3, 4, CLng(contacto))
	p(2)  = Array("@id_aplic", 3, 4, CLng(idAplic))
	p(3)  = Array("@detalles", 200, 4000, CStr(detalles))
	p(4)  = Array("@prioridad", 3, 4, CLng(prioridad))
	p(5)  = Array("@respon_gestion", 200, 100, CStr(responGestion))
	p(6)  = Array("@respon_tec", 200, 100, CStr(responTec))
	p(7)  = Array("@req_presup", 200, 10, CStr(reqPresup))
	p(8)  = Array("@fecha_ingreso", 3, 4, CLng(fechaIngreso))
	p(9)  = Array("@browser", 3, 4, CLng(browser))
	p(10) = Array("@via_pedido", 3, 4, CLng(viaPedido))
	p(11) = Array("@titulo", 200, 200, CStr(titulo))
	p(12) = Array("@responsable", 200, 20, CStr(responsableRol))

	EjecutarSP_NonQuery cn, "dbo.sp_cg_Incidentes_Insertar", p

End Sub

Sub InsertarProyecto(cn, solicNro, tituloP, descripcion, contacto, idCliente, idAplic, _
						responGestion, responTec, fechaIngreso, fechaEntrega, prioridad, _
						tRelev, tDiseno, tDesarr, tTest, browser, viaPedido, reqPresup, responsableRol, tipoproyecto)

	Dim p(19)
	p(0)  = Array("@solic_nro", 200, 50, CStr(solicNro))
	p(1)  = Array("@titulo", 200, 200, CStr(tituloP))
	p(2)  = Array("@descripcion", 200, 4000, CStr(descripcion))
	p(3)  = Array("@contacto", 3, 4, CLng(contacto))
	p(4)  = Array("@id_cliente", 3, 4, CLng(idCliente))
	p(5)  = Array("@id_aplic", 3, 4, CLng(idAplic))
	p(6)  = Array("@respon_gestion", 200, 100, CStr(responGestion))
	p(7)  = Array("@respon_tec", 200, 100, CStr(responTec))
	p(8)  = Array("@fecha_ingreso", 3, 4, CLng(fechaIngreso))
	p(9)  = Array("@fecha_entrega", 3, 4, CLng(fechaEntrega))
	p(10) = Array("@prioridad", 3, 4, CLng(prioridad))
	p(11) = Array("@tiempo_relevamiento", 5, 8, CDbl(tRelev))
	p(12) = Array("@tiempo_diseno", 5, 8, CDbl(tDiseno))
	p(13) = Array("@tiempo_desarr", 5, 8, CDbl(tDesarr))
	p(14) = Array("@tiempo_test", 5, 8, CDbl(tTest))
	p(15) = Array("@browser", 3, 4, CLng(browser))
	p(16) = Array("@via_pedido", 3, 4, CLng(viaPedido))
	p(17) = Array("@req_presup", 200, 10, CStr(reqPresup))
	p(18) = Array("@responsable", 200, 20, CStr(responsableRol))
	p(19) = Array("@TipoProyecto", 3, 4, CLng(tipoproyecto))

	EjecutarSP_NonQuery cn, "dbo.sp_cg_Proyectos_Insertar", p

End Sub

Sub InsertarLicencia(cn, idRespon, fInicio, fFin, motivo)

	Dim p(3)
	p(0) = Array("@id_respon", 200, 100, CStr(idRespon))
	p(1) = Array("@f_inicio", 3, 4, CLng(fInicio))
	p(2) = Array("@f_fin", 3, 4, CLng(fFin))
	p(3) = Array("@motivo", 200, 4000, CStr(motivo))

	EjecutarSP_NonQuery cn, "dbo.sp_cg_Licencias_Insertar", p

End Sub

'------------------------------------------------------------------------
' Parte B - Combos de catalogo (via SP consolidado sp_cg_Catalogo_Listar /
' sp_cg_Combo_Listar). Cada uno respeta la cantidad de columnas que ya
' consumia agregar2.asp.
'------------------------------------------------------------------------

Function ObtenerClientes(cn)
	On Error Resume Next
	Set ObtenerClientes = cn.Execute("sp_cg_Catalogo_Listar @Catalogo='clientes'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerClientes<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

Function ObtenerAplicativos(cn)
	On Error Resume Next
	Set ObtenerAplicativos = cn.Execute("sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerAplicativos<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

Function ObtenerPrioridades(cn)
	On Error Resume Next
	Set ObtenerPrioridades = cn.Execute("sp_cg_Catalogo_Listar @Catalogo='prioridades'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerPrioridades<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

' funcion: "FUNC" o "TEC" (mismo filtro que usaba el original)
Function ObtenerResponsables(cn, funcion)
	Dim cmd
	'respose.write "func:" & cstr(funcion)
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "sp_cg_Responsables"
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 25, "listar_funcion")
	cmd.Parameters.Append cmd.CreateParameter("@Id", 200, 1, 50, Null)
	cmd.Parameters.Append cmd.CreateParameter("@Funcion", 200, 1, 50, funcion)
	'Set ObtenerResponsables = cmd.Execute()
	
	On Error Resume Next
    Set ObtenerResponsables = cmd.Execute
    If Err.Number <> 0 Then
        Response.Write "ERROR " & Err.Number & ": " & Err.Description & " Valor:"  & cstr(funcion)
        Response.End
    End If
    On Error Goto 0
	
End Function

Function ObtenerViaPedido(cn)
	On Error Resume Next
	Set ObtenerViaPedido = cn.Execute("sp_cg_Catalogo_Listar @Catalogo='via_pedido'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerViaPedido<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

Function ObtenerBrowsers(cn)
	On Error Resume Next
	Set ObtenerBrowsers = cn.Execute("sp_cg_Catalogo_Listar @Catalogo='browsers'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerBrowsers<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

' Variante usada en el form de INCIDENTE (con join a clientes, aunque el
' join no se usa para el order by; se preserva tal cual estaba)
Function ObtenerContactosParaIncidente(cn)
	On Error Resume Next
	Set ObtenerContactosParaIncidente = cn.Execute("sp_cg_Combo_Listar @Origen='contacto'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerContactosParaIncidente<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function

' Variante usada en el form de PROYECTO (sin join)
Function ObtenerContactosParaProyecto(cn)
	On Error Resume Next
	Set ObtenerContactosParaProyecto = cn.Execute("sp_cg_Combo_Listar @Origen='contacto'")
	If Err.Number <> 0 Then
		Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
		Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
		Response.Write "Contexto: ObtenerContactosParaProyecto<br>"
		Response.Write "Codigo: " & Err.Number & "<br>"
		Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
		Response.End
	End If
	On Error Goto 0
End Function
%>
