
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok = ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	shtml = ok
Else

	'response.write request.querystring
	empre=request.querystring("empre")
	Set cn = ofv.conectar(ofv.strconn2)

	WSPFRMN="frm"
	Session("USRSET")=1
	sql="select * from clienteplan where id=0" & empre
	'  response.write sql
	Set rsx = ofv.crearconsultaEx(sql, cn, 1, parametros)

	If Not rsx.eof Then
		Session("USRSET")=rsx("codset")
		Session("Cliente") = rsx(1)
		Session("CliLogo") = "../COREss20/Empresa/" & rsx(5)
		If (rsx(11) & "" ) <> "" Then Session("CliLogoP") = "../COREss20/Empresa/" & rsx(11)
	End If

	'  XX=ConUsrset()
	Call ofv.cerrarconsulta(rsx)

	xx=ConUsrset()
	SESSION("UN")=empre

	WSPFORME2 = "<FORM ID='" & WSPFRMN & "2' NAME='" & WSPFRMN & "2' ACTION='../CORESS20/ESTILO.ASP?op=ex&empre=" & empre & "' TARGET='estilo' style='display:none;' METHOD='POST' ></FORM>"
	'WSPFORME3 = "<FORM ID='" & WSPFRMN & "3' NAME='" & WSPFRMN & "3' ACTION='../coremanager/novedmenu.asp?EMPRE=" & empre & "' TARGET='info' style='display:none;' METHOD='POST' ></FORM>"
	shtml= WSPFORME2 & WSPFORME3
	shtml=shtml & "<script>"
	shtml=shtml & "document." & WSPFRMN & "2.submit();"
	'shtml=shtml & "document." & WSPFRMN & "3.submit();"

	shtml=shtml & "</script>"


	response.write shtml
	Call ofv.cerrarconn(cn)
End If


%>