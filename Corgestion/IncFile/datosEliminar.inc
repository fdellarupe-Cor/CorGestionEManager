<%
'==============================================================================
' IncFile/datosEliminar.inc
' Capa de ACCESO A DATOS para eliminar.asp. Toda la logica de baja/borrado
' (logica vs fisico, chequeos de uso, cascadas) vive ahora en dbo.sp_cg_Eliminar.
'
' EliminarRegistro devuelve True/False. False significa que el SP disparo un
' RAISERROR (por ejemplo, "el registro esta en uso") -- el ASP decide que
' mensaje mostrar segun el opc, igual que hacia el archivo original.
'==============================================================================

Function EliminarRegistro(cn, opt, id, usuario)
	Dim cmd
	On Error Resume Next

	Err.Clear
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4 ' adCmdStoredProc
	cmd.CommandText = "dbo.sp_cg_Registro_Eliminar"
	cmd.Parameters.Append cmd.CreateParameter("@Opt", 200, 1, 20, opt)
	cmd.Parameters.Append cmd.CreateParameter("@Id", 3, 1, 4, CLng(id))
	If usuario = "" Then
		cmd.Parameters.Append cmd.CreateParameter("@usuario", 200, 1, 100, Null)
	Else
		cmd.Parameters.Append cmd.CreateParameter("@usuario", 200, 1, 100, usuario)
	End If
	cmd.Execute

	If Err.Number <> 0 Then
		EliminarRegistro = False
		Err.Clear
	Else
		EliminarRegistro = True
		' [LOG] Baja de incidente / proyecto / APL / recordatorio de agenda (CG_LOG)
		Select Case LCase(opt & "")
			Case "inc": CgLog "INCIDENTE", id, "BAJA"
			Case "proy": CgLog "PROYECTO", id, "BAJA"
			Case "apl": CgLog "APL", id, "BAJA"
			Case "rec": CgLog "AGENDA", id, "BAJA"
			Case "apli": CgLog "APLICATIVO", id, "BAJA"
			Case "brow": CgLog "BROWSER", id, "BAJA"
			Case "cli": CgLog "CLIENTE", id, "BAJA"
			Case "con": CgLog "CONTACTO", id, "BAJA"
			Case "corr": CgLog "RESPONSABLE", id, "BAJA"
			Case "est": CgLog "ESTADO", id, "BAJA"
			Case "prio": CgLog "PRIORIDAD", id, "BAJA"
			Case "viap": CgLog "VIA_PEDIDO", id, "BAJA"
			Case "lic": CgLog "LICENCIA", id, "BAJA"
		End Select
	End If

	On Error Goto 0
End Function
%>
