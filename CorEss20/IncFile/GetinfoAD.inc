<%


Dim strBase
Dim strBase2
Dim userN
Dim pass
Dim ZZ
Dim Srbac
Dim sMailLdap

Srbac=session("Srbac")
sMailLdap=session("sMailLdap")
'response.write REQUEST.servervariables("LOGON_USER") & "<br>"
PageSizead=100
DOMC= REQUEST.servervariables("LOGON_USER")

'DOMC= "SANTAFE\jurer"
DOMC=Split(DOMC,"\")

If UBound(DOMC) >=0 Then
	SESSION("DOMINIOX")=DOMC(0)
Else
	SESSION("DOMINIOX")=""
End If

'SESSION("DOMINIOX")="SJ"

'SESSION("DOMINIOX")="DOMINIO ER"
'SESSION("DOMINIOX")="DOMINIO SJ"
'SESSION("DOMINIOX")="SANTACRUZ"


'response.write session("dominioX") & "222"


XXX=ValDominio(session("dominioX"))

strBase =  session("ldap")
strBase2 =Replace(strBase & "" ,"<","")
strBase2=Replace(strBase2 & "" ,">","")

userN = session("usrldap")
pass = session("pwdldap")



Function getUsrIngroupAD(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0))
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL OR DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")

		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
		Set objRecordSet = objCommand.Execute

		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		objRecordSet = Nothing
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function

'--------------------------------------------------------------------------------------------------
Function getUsrIngroupADSV(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )

		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&"")) 'rsdom("pwdldap")

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")
		GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'  and userAccountControl  <> 66050 and userAccountControl  <> 546 order by name"

		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		'     objRecordSet = NOTHING
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function


Function getUsrIngroupADLacaja(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )

		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")
		GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		'     objRecordSet = NOTHING
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function


Function getUsrIngroupADBANELCO(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )

		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")
		GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		'     objRecordSet = NOTHING
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function



Function getUsrIngroupAD_FINANSUR(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )

		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")
		GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		'     objRecordSet = NOTHING
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function


Function getUsrIngroupAD_BPN(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)

	' -- Inicio nuevo Bloque
	'  Ahora antes de Preguntar a AD buscara los integrantes de un GRUPO
	' en la tabla OBJALOG
	
	

	Strsqlq = "select USRNET as usuario, USRNAME, MAIL,GRUPO  from vw_UserGrupos	 where  grupo= '" & grupo & "'"
	
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )
	
	
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & rsdom("usuario")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & rsdom("USRNAME")
			xgusep = Chr(9)
			rsdom.movenext
	
	Loop
	Call ofv.cerrarconsulta(rsdom)
 
	
	if session("SGU_UIGuserid") = "" Then ' Si no se encontro mienbros del grupo en OBJALOG busco en AD
		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
		session("SGU_UIGuserid") = ""
		session("SGU_UIGdescri") = ""

		Do While (Not rsdom.eof )

			sldap =rsdom("raizldap")
			usernx =  rsdom("usrldap")
			passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

			strBase =  rsdom("raizldap")
			strBase2 =Replace(strBase ,"<","")
			strBase2=Replace(strBase2 ,">","")
			GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
			Set objConnection = server.CreateObject("ADODB.Connection")
			Set objCommand = server.CreateObject("ADODB.Command")
			objConnection.Provider = "ADsDSOObject"
			objConnection.Properties("User ID") = usernx
			objConnection.Properties("Password") = passx
			objConnection.Properties("Encrypt Password") = True
			objConnection.OPEN "Active Directory Provider"
			Set objCommand.ActiveConnection = objConnection

			objCommand.Properties("Page Size") = PageSizead
			objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

			objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

			Set objRecordSet = objCommand.Execute


			Do Until objRecordSet.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
				xgusep = Chr(9)

				objRecordSet.MoveNext

			Loop


			objRecordSet.CLOSE
			'     objRecordSet = NOTHING
			objConnection.CLOSE
			Set objCommand = Nothing
			Set objConnectionn = Nothing
			rsdom.movenext
		Loop
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnAD)
	end If
	
End Function

Function getUsrIngroupAD_GP(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree

	' RESPONSE.WRITE ofv.strconn2

	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	'  RESPONSE.WRITE Strsql
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		'GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)

	If Session("EMPMUN")  = "S" Then
		Strsqlq = "Select * From dominios WHERE cliente=" & session("SGU_site")
	Else
		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	End If
	'response.write Strsqlq
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )

		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")
		GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail,displayName FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
		' RESPONSE.WRITE "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			If rsdom("id")=7 Then
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("displayName")
			Else
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			End If

			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		'     objRecordSet = NOTHING
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function


Function getUsrIngroupAD_BCR(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)

	If Trim(GrupoADX & "")<> "" Then
		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
		session("SGU_UIGuserid") = ""
		session("SGU_UIGdescri") = ""

		Do While (Not rsdom.eof )

			sldap =rsdom("raizldap")
			usernx =  rsdom("usrldap")
			passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

			strBase =  rsdom("raizldap")
			strBase2 =Replace(strBase ,"<","")
			strBase2=Replace(strBase2 ,">","")
			GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
			Set objConnection = server.CreateObject("ADODB.Connection")
			Set objCommand = server.CreateObject("ADODB.Command")
			objConnection.Provider = "ADsDSOObject"
			objConnection.Properties("User ID") = usernx
			objConnection.Properties("Password") = passx
			objConnection.Properties("Encrypt Password") = True
			objConnection.OPEN "Active Directory Provider"
			Set objCommand.ActiveConnection = objConnection

			objCommand.Properties("Page Size") = PageSizead
			objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

			objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail,displayName FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

			Set objRecordSet = objCommand.Execute


			Do Until objRecordSet.EOF

				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("displayName")
				xgusep = Chr(9)

				objRecordSet.MoveNext

			Loop


			objRecordSet.CLOSE
			'     objRecordSet = NOTHING
			objConnection.CLOSE
			Set objCommand = Nothing
			Set objConnectionn = Nothing
			rsdom.movenext
		Loop
		Call ofv.cerrarconsulta(rsdom)
	End If
	Call ofv.cerrarconn(cnAD)

End Function


Function GetUserInsectAd(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0) &"*"
	'session("xxx")=WXGU_SECT
	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set connaXz = server.CreateObject("ADODB.Connection")


	connaXz.provider = "ADSDSOObject"
	connaXz.Properties("User ID") = usern
	connaXz.Properties("Password") = pass
	connaXz.Properties("Encrypt Password") = True

	connaXz.OPEN

	Set oCommand = server.CreateObject("ADODB.Command")
	Set oCommand.ActiveConnection = connaXz
	oCommand.Properties("Searchscope")=2


	oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

	'response.write oCommand.CommandText
	Set RSAXz = oCommand.Execute


	'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

	' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
	' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery

	'  SET RSAXz = connaXz.EXECUTE(strQuery)



	Do Until RSAXz.EOF
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
		xgusep = Chr(9)
		RSAXz.movenext
	Loop

	RSAXz.CLOSE
	Set RSAXz = Nothing
	connaXz.CLOSE
	Set connaXz = Nothing

End Function



Function getUserInSectAdBANELCO(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0) &"*"
	'session("xxx")=WXGU_SECT
	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set connaXz = server.CreateObject("ADODB.Connection")


	connaXz.provider = "ADSDSOObject"
	connaXz.Properties("User ID") = usern
	connaXz.Properties("Password") = pass
	connaXz.Properties("Encrypt Password") = True

	connaXz.OPEN

	Set oCommand = server.CreateObject("ADODB.Command")
	Set oCommand.ActiveConnection = connaXz
	oCommand.Properties("Searchscope")=2


	oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

	'response.write oCommand.CommandText
	Set RSAXz = oCommand.Execute


	'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

	' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
	' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery

	'  SET RSAXz = connaXz.EXECUTE(strQuery)



	Do Until RSAXz.EOF
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
		xgusep = Chr(9)
		RSAXz.movenext
	Loop

	RSAXz.CLOSE
	Set RSAXz = Nothing
	connaXz.CLOSE
	Set connaXz = Nothing

End Function




Function getUserInSectAd_FINANSUR(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0) &"*"
	'session("xxx")=WXGU_SECT
	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set connaXz = server.CreateObject("ADODB.Connection")


	connaXz.provider = "ADSDSOObject"
	connaXz.Properties("User ID") = usern
	connaXz.Properties("Password") = pass
	connaXz.Properties("Encrypt Password") = True

	connaXz.OPEN

	Set oCommand = server.CreateObject("ADODB.Command")
	Set oCommand.ActiveConnection = connaXz
	oCommand.Properties("Searchscope")=2


	oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

	'response.write oCommand.CommandText
	Set RSAXz = oCommand.Execute


	'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

	' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
	' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery

	'  SET RSAXz = connaXz.EXECUTE(strQuery)



	Do Until RSAXz.EOF
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
		xgusep = Chr(9)
		RSAXz.movenext
	Loop

	RSAXz.CLOSE
	Set RSAXz = Nothing
	connaXz.CLOSE
	Set connaXz = Nothing

End Function
'---------------------------------------------------------------------------------------------------------------------------------------------



Function getUserInSectAd_BPN(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0) &"*"
	'session("xxx")=WXGU_SECT
	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set connaXz = server.CreateObject("ADODB.Connection")


	connaXz.provider = "ADSDSOObject"
	connaXz.Properties("User ID") = usern
	connaXz.Properties("Password") = pass
	connaXz.Properties("Encrypt Password") = True

	connaXz.OPEN

	Set oCommand = server.CreateObject("ADODB.Command")
	Set oCommand.ActiveConnection = connaXz
	oCommand.Properties("Searchscope")=2


	oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

	'response.write oCommand.CommandText
	Set RSAXz = oCommand.Execute


	'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

	' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
	' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery

	'  SET RSAXz = connaXz.EXECUTE(strQuery)



	Do Until RSAXz.EOF
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
		xgusep = Chr(9)
		RSAXz.movenext
	Loop

	RSAXz.CLOSE
	Set RSAXz = Nothing
	connaXz.CLOSE
	Set connaXz = Nothing

End Function


Function getUserInSectAd_GP(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0) &"*"
	'session("xxx")=WXGU_SECT
	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set connaXz = server.CreateObject("ADODB.Connection")


	connaXz.provider = "ADSDSOObject"
	connaXz.Properties("User ID") = usern
	connaXz.Properties("Password") = pass
	connaXz.Properties("Encrypt Password") = True

	connaXz.OPEN

	Set oCommand = server.CreateObject("ADODB.Command")
	Set oCommand.ActiveConnection = connaXz
	oCommand.Properties("Searchscope")=2


	oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

	'response.write oCommand.CommandText
	Set RSAXz = oCommand.Execute


	'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

	' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
	' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery

	'  SET RSAXz = connaXz.EXECUTE(strQuery)



	Do Until RSAXz.EOF
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
		xgusep = Chr(9)
		RSAXz.movenext
	Loop

	RSAXz.CLOSE
	Set RSAXz = Nothing
	connaXz.CLOSE
	Set connaXz = Nothing

End Function

Function getUserInSectAd_BCR(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0) &"*"
	'session("xxx")=WXGU_SECT
	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set connaXz = server.CreateObject("ADODB.Connection")


	connaXz.provider = "ADSDSOObject"
	connaXz.Properties("User ID") = usern
	connaXz.Properties("Password") = pass
	connaXz.Properties("Encrypt Password") = True

	connaXz.OPEN

	Set oCommand = server.CreateObject("ADODB.Command")
	Set oCommand.ActiveConnection = connaXz
	oCommand.Properties("Searchscope")=2


	oCommand.CommandText = "SELECT postOfficeBox,wwwHomePage,sAMAccountName,name FROM '"& strBase2 &"' WHERE "  & "objectCategory='person' and objectclass='user' and wwwHomePage='"& WXGU_SECT  &"' order by postOfficeBox"

	'response.write oCommand.CommandText
	Set RSAXz = oCommand.Execute


	'strAttributes = "wwwHomePage,sAMAccountName,name,title,mail,memberOf"

	' strFilter = "(&(objectCategory=person)(objectClass=user)" & "(wwwHomePage=" & WXGU_SECT &"))" '& "(samAccountName=" & userid &"))"
	' strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery

	'  SET RSAXz = connaXz.EXECUTE(strQuery)



	Do Until RSAXz.EOF
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAXz("sAMAccountName")
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAXz("name")
		xgusep = Chr(9)
		RSAXz.movenext
	Loop

	RSAXz.CLOSE
	Set RSAXz = Nothing
	connaXz.CLOSE
	Set connaXz = Nothing

End Function
'---------------------------------------------------------------------------------------------------------------------------------------------
Function getUsrIngroupADBCOLU(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
		grupoadx2=  rsADX(0)
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""

	Do While (Not rsdom.eof )

		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&"")) 'rsdom("pwdldap")

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")
		GrupoADX =  FullNameGrupo(GrupoADX2,strBase )
		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"

		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("name")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE
		'     objRecordSet = NOTHING
		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function







'-----------------------------------------------------------------------------------







Function getGroupsUserAD(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strsql
	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)

	session("SGU_grupos")=""
	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				' xgrupo=split(strGroup,",")

				session("SGU_grupos") = session("SGU_grupos") & xgusep  & strGroup
				xgusep = Chr(9)


			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function

'--------------------------------------------------------------------------------------------------
Function getGroupsUserADBCOLU(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)


	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")
				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function


Function getGroupsUserADLACAJA(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)


	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")
				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function


Function getGroupsUserADBANELCO(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)


	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")

				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function


Function getGroupsUserAD_FINANSUR(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)


	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")

				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------


Function getGroupsUserAD_BPN(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True
	'response.write pass
	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)


	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")

				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function

Function getGroupsUserAD_GP(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True
	'response.write userN
	'EXIT FUNCTION
	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)

	session("SGU_grupos")=""
	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				' strGroup = UCASE(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")

				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo

				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing
	'response.write session("SGU_grupos")
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------------
Function getGroupsUserADSV(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write pass

	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)


	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")
				' strGroup=replace(strGroup,"CN=","")
				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")
				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
				xgusep = Chr(9)
			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing


End Function


Function getInfouserWLACAJA(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = legajo & "@lacaja.com.ar" 'RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function



Function getInfouserWBANELCO(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = getperfilExtBanelco(rsax("title"))
			If IsArray(rsax("description"))Then
				xnivel=rsax("description")
				session("SGU_idnivel") =xnivel(0)
			Else
				xnivel=rsax("description")
				session("SGU_idnivel") =xnivel
			End If
			session("SGU_idnivel")=getnivelExtBanelco(UCase(session("SGU_idnivel")))
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function

Function getInfouserW_FINANSUR(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

			session("SGU_idnivel")="9"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function


Function getInfouserW_BPN(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

			session("SGU_idnivel")="9"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function

Function getInfouserW_COR(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")

			If IsNull(rsax("description").Value) Then
				valorDescription = "" ' Está vacío en Active Directory
			ElseIf IsArray(rsax("description").Value) Then
				' Es un arreglo. Usamos Join para convertirlo a texto si tiene varios valores,
				' o simplemente tomamos la posición (0)
				valorDescription = Join(rsax("description").Value, " ")
			Else
				valorDescription = rsax("description").Value & ""
			End If
			 session("SGU_codigopuesto") = getperfil(valorDescription)
             session("SGU_descripuesto") = getperfildescri( session("SGU_codigopuesto"))
			'session("SGU_codigopuesto") =getperfil(rsax("description"))

			session("SGU_idnivel")="9"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function


Function getInfouserW_GP(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	'Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

	If Session("EMPMUN")  = "S" Then
		If SESSION("CLIEMP")=0 Then SESSION("CLIEMP")=session("usr_dominio")
		Strsqlq = "Select * From dominios WHERE cliente=" & SESSION("CLIEMP")
	Else
		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	End If



	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"
			session("SGU_site")=1

		Else

			session("SGU_userid") = Legajo
			If rsdom("id")=7 Then
				session("SGU_nombre") = RSAX("displayName")
			Else
				session("SGU_nombre") = RSAX("name")
			End If
			'   session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

			session("SGU_idnivel")="9"
			session("SGU_site")=session("usr_dominio")
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function


Function getInfouserWEX_GP(legajo,DATOSE)
	' DATOSE codigo de empresa en DATOSE(0)
	'chr(9) para separar datos por si hay que agregar mas paramentros


	Dim cempre
	dim DATOSEx
	cempre=0
	DATOSEx=DATOSE
	If Trim(DATOSEx) = "" Then
	Else
		DATOSEx=Split(Trim(DATOSEx & ""),Chr(9))

		If Len(DATOSEx(0) & "")  > 0 Then
			cempre=DATOSEx(0)
		End If
	End If


	Set cnDOM = ofv.conectar(ofv.strconn0)
	'Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	If cempre > 0 Then
		Strsqlq = "Select * From dominios WHERE  cliente=" & cempre
	Else
		If Session("EMPMUN")  = "S" Then
			If SESSION("CLIEMP")=0 Then SESSION("CLIEMP")=session("usr_dominio")
			Strsqlq = "Select * From dominios WHERE cliente=" & SESSION("CLIEMP")
		Else
			Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
		End If
	End If


	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"
			session("SGU_site")=1

		Else

			session("SGU_userid") = Legajo
			If rsdom("id")=7 Then
				session("SGU_nombre") = RSAX("displayName")
			Else
				session("SGU_nombre") = RSAX("name")
			End If
			'   session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

			session("SGU_idnivel")="9"
			session("SGU_site")=session("usr_dominio")
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function


Function getInfouserw_BCR(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		'strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"
		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel")="9"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("displayName")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = getsectorExtBCR(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") ="NO" 'getperfilExtBanelco(rsax("title"))

			session("SGU_idnivel")=getnivelExtBCR(rsax("title")&"")
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function


Function getInfouserWAD(legajo)

	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------
Function getInfouserWBCOLU(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------




Function getInfouserWSV(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function


Function getInfouserWBC(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,name,cn,title,mail,memberOf"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"


		Else

			Set Cnpeople=ofv.conectar(ofv.strconn7)
		'	legajoZ=Mid(legajo,3,Len(legajo))
			' 21/2/2019 por pedido del banco
            legajoz=legajo
			strsqlCn= "select ALTER_EMPLID  ,NAME1, DEPTID  ,JOBCODE   from PS_BC_NORMAS_VW  where UPPER(ALTER_EMPLID)='" & UCASE(legajoZ) &"'"
			Set CnViewPeople=  ofv.crearconsultaEx(strsqlCn ,Cnpeople ,1,volver)

			If Not CnViewPeople.eof Then

				session("SGU_userid") = Legajo
				session("SGU_nombre") =Replace(CnViewPeople("NAME1")& "" ,"'"," ") ' & " " &  CnViewPeople("MIDDLE_NAME") & " " & CnViewPeople("LAST_NAME"),"'"," ")
				session("SGU_mail") = ChangeMail(RSAX("mail"))
				session("SGU_descrisector") = "NO"
				SDEPTID = CnViewPeople("DEPTID") & ""
				SDEPTID=Trim(SDEPTID)
				session("SGU_idcostcenter") = getsectorExt(SDEPTID) 'CnViewPeople("DEPTID")
				session("SGU_descripuesto") = ""
				SXJOB=CnViewPeople("JOBCODE") & ""
				SXJOB=Trim(SXJOB)
				session("SGU_codigopuesto") = SXJOB
				session("SGU_idnivel")      =getnivelExt(SXJOB)
			Else

				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
				session("SGU_mail") = "NO@disponible.com"
				session("SGU_descrisector") = "NO disponible"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO disponible"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel")      ="9"
			End If

			Call ofv.cerrarconsulta(CnViewPeople)
			Call ofv.cerrarconn(Cnpeople)
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function







Function getUsrInTextAD(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	TEXTO=Replace(TEXTO,"*","")
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="NO"
	ElseIf Len(texto)=1 Then
		texto= texto & "*"
	Else
		texto="*"  & texto & "*"
	End If
	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,name"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		' response.write strQuery

		Set RSAX = connaX.Execute(strQuery)



		Do Until RSAX.EOF
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
			xgusep = Chr(9)
			RSAX.movenext
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.movenext
	Loop

End Function


Function getUsrInTextADBANELCO(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,name"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		' response.write strQuery

		Set RSAX = connaX.Execute(strQuery)



		Do Until RSAX.EOF
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
			xgusep = Chr(9)
			RSAX.movenext
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.movenext
	Loop

End Function


Function getUsrInTextAD_FINANSUR(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	TEXTO=Replace(TEXTO,"*","")
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="NO"
	Else
		If Trim(Len(texto&""))=1 Then
			texto=texto & "*"
		Else
			texto="*"  & texto & "*"
		End If
	End If

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,name,displayName"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		strFilter="(&(objectCategory=person)(objectClass=user)(displayName=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		' response.write strQuery

		Set objCommand = server.CreateObject("ADODB.Command")

		Set objCommand.ActiveConnection = connaX

		objCommand.Properties("Size Limit") =  PageSizead
		objCommand.Properties("Searchscope") = 2

		objCommand.CommandText=strQuery
		' response.write  strQuery
		Set RSAX = objCommand.Execute '(strQuery)

		totu=0

		Do While Not  RSAX.EOF And totu < PageSizead
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("displayName")
			xgusep = Chr(9)
			RSAX.movenext
			totu=totu+1
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		Set objCommand=Nothing
		connaX.CLOSE
		Set connaX = Nothing

		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function



Function getUsrInTextAD_BPN(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	TEXTO=Replace(TEXTO,"*","")
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="NO"
	ElseIf Len(texto)=1 Then
		texto= texto & "*"
	Else
		texto="*"  & texto & "*"
	End If

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		Set objCommand = server.CreateObject("ADODB.Command")

		Set objCommand.ActiveConnection = connaX

		objCommand.Properties("Size Limit") = 500
		objCommand.Properties("Searchscope") = 2
		' objCommand.Properties("Sort On") = "NAME"

		Set objCommand.ActiveConnection = connaX
		strAttributes = "sAMAccountName,cn,title,mail,name"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		objCommand.CommandText=strQuery
		Set RSAX = objCommand.EXECUTE'(strQuery)


		totu=0
		Do While Not  RSAX.EOF And totu < PageSizead
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
			xgusep = Chr(9)
			RSAX.movenext
			totu=totu+1
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		' objCommand.close
		Set objCommand = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.movenext
	Loop

End Function


Function getUsrInTextAD_GP(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)

	' Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

	If Session("EMPMUN")  = "S" Then
		Strsqlq = "Select * From dominios WHERE cliente=0" & session("SETIDEMPRE")
	Else
		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	End If
	'       response.write Strsqlq
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	TEXTO=Replace(TEXTO,"*","")
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="NO"
	ElseIf Len(texto)=1 Then
		texto= texto & "*"
	Else
		texto="*"  & texto & "*"
	End If

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		Set objCommand = server.CreateObject("ADODB.Command")

		Set objCommand.ActiveConnection = connaX

		objCommand.Properties("Size Limit") = 500
		objCommand.Properties("Searchscope") = 2
		' objCommand.Properties("Sort On") = "NAME"

		Set objCommand.ActiveConnection = connaX
		strAttributes = "sAMAccountName,cn,title,mail,name,displayname"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		If rsdom("id") = 7 Then
			strFilter="(&(objectCategory=person)(objectClass=user)(displayname=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"
		Else
			strFilter="(&(objectCategory=person)(objectClass=user)(name=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"
		End If

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		objCommand.CommandText=strQuery
		Set RSAX = objCommand.EXECUTE'(strQuery)


		totu=0
		Do While Not  RSAX.EOF And totu < PageSizead
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			If rsdom("id") = 7 Then
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("displayname")
			Else
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
			End If

			xgusep = Chr(9)
			RSAX.movenext
			totu=totu+1
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		' objCommand.close
		Set objCommand = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.movenext
	Loop

End Function



Function getGruposMail_BPN(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	TEXTO=Trim(Replace(TEXTO,"*",""))
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="*"
	Else
		texto="*"  & texto & "*"
	End If

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True


		connaX.OPEN
		Set objCommand = server.CreateObject("ADODB.Command")

		Set objCommand.ActiveConnection = connaX

		objCommand.Properties("Size Limit") = PageSizead
		objCommand.Properties("Searchscope") = 2
		objCommand.Properties("Sort On") = "name"
		strAttributes = "name,mail"
		' strFilter = "(&(objectCategory=group)((&(groupType:1.2.840.113556.1.4.803:=8)(!(groupType:1.2.840.113556.1.4.803:=-2147483648)))))"


	'	strFilter = "(&((&(objectCategory=group)(&(groupType:1.2.840.113556.1.4.803:=8)(!(groupType:1.2.840.113556.1.4.803:=-2147483648))))) (name=" & texto &  ") )"

	'	 activar siguiente linea en cor porque no hay exchange
		strFilter =  	"(&(objectCategory=group)(name=" &texto&"))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		objCommand.CommandText=strQuery
		' response.write  strQuery
		Set RSAX = objCommand.Execute '(strQuery)

		' response.write RSAX.EOF


		Do While Not  RSAX.EOF

			If ValidaMail(rsax("mail")& "") = 2 Then
				session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("name")
				session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("Name")
				xgusep = Chr(9)
			End If
			RSAX.movenext
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.movenext
		if Session("EMPMUN") = "N"  then exit do
	Loop

End Function

Function getUsrInTextAD_BCR(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	TEXTO=Replace(TEXTO,"*","")
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="NO"
	Else
		If Trim(Len(texto&""))=1 Then
			texto=texto & "*"
		Else
			texto="*"  & texto & "*"
		End If
	End If

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,name,displayName"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		strFilter="(&(objectCategory=person)(objectClass=user)(displayName=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		' response.write strQuery

		Set objCommand = server.CreateObject("ADODB.Command")

		Set objCommand.ActiveConnection = connaX

		objCommand.Properties("Size Limit") =  PageSizead
		objCommand.Properties("Searchscope") = 2

		objCommand.CommandText=strQuery
		' response.write  strQuery
		Set RSAX = objCommand.Execute '(strQuery)

		totu=0

		Do While Not  RSAX.EOF And totu < PageSizead
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("displayName")
			xgusep = Chr(9)
			RSAX.movenext
			totu=totu+1
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		Set objCommand=Nothing
		connaX.CLOSE
		Set connaX = Nothing

		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function

Function getMailDelgrupo_BPN(texto)



	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	TEXTO=Trim(Replace(TEXTO,"*",""))
	getMailDelgrupo_BPN=""

	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True


		connaX.OPEN
		Set objCommand = server.CreateObject("ADODB.Command")

		Set objCommand.ActiveConnection = connaX

		objCommand.Properties("Size Limit") = 500
		objCommand.Properties("Searchscope") = 2

		strAttributes = "name,mail"

		strFilter =  	"(&(objectCategory=group)(name=" &texto&"))"
		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		objCommand.CommandText=strQuery
		'response.write  strQuery
		Set RSAX = objCommand.Execute

		' response.write RSAX.EOF


		If Not  RSAX.EOF Then

			' if ValidaMail(rsax("mail")& "") = 2 then
			getMailDelgrupo_BPN=rsax("mail")&""
			Exit Do
			'  end if
		End If
		RSAX.CLOSE
		Set RSAX = Nothing

		rsdom.movenext
	Loop


	connaX.CLOSE
	Set connaX = Nothing

End Function





Function getinfoUserAD(legajo)

	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"
	' RESPONSE.WRITE Strsqlq
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	' session("SGU_nombre")=""
	Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
		sldap =rsdom("raizldap")

		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		Set RSAX = connaX.Execute(strQuery)



		If RSAX.EOF Then
			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
			session("SGU_site")=1

		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") =  RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			session("SGU_site")=1
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)


End Function


Function getinfoUserADlacaja(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
		sldap =rsdom("raizldap")


		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		Set RSAX = connaX.Execute(strQuery)



		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"

		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = legajo & "@lacaja.com.ar" ' RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)

End Function



Function getinfoUserADBanelco(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
		sldap =rsdom("raizldap")


		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		Set RSAX = connaX.Execute(strQuery)



		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel") ="9"

		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = getperfilExtBanelco(rsax("title"))
			If IsArray(rsax("description"))Then
				xnivel=rsax("description")
				session("SGU_idnivel") =xnivel(0)
			Else
				xnivel=rsax("description")
				session("SGU_idnivel") =xnivel
			End If
			session("SGU_idnivel")=getnivelExtBanelco(UCase(session("SGU_idnivel")))
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)

End Function


Function getinfoUserAD_FINANSUR(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
		sldap =rsdom("raizldap")


		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		Set RSAX = connaX.Execute(strQuery)



		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
			session("SGU_idnivel") ="9"

		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("name")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
			session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)

End Function



Function getinfoUserAD_BPN(legajo)

	If Len(Trim(LEGAJO & "")) = 0 Then

		session("SGU_userid") = "--" & Legajo & "-"
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel") ="9"
	Else
		Set cnDOM = ofv.conectar(ofv.strconn0)

		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

		Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
			sldap =rsdom("raizldap")


			usernx =  rsdom("usrldap")
			passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

			Set connaX = server.CreateObject("ADODB.Connection")


			connaX.provider = "ADSDSOObject"
			connaX.Properties("User ID") = usernx
			connaX.Properties("Password") = passx
			connaX.Properties("Encrypt Password") = True
			' response.write userN & "----" & pass & "---" & strBase
			connaX.OPEN

			strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

			strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


			strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
			'response.write strQuery
			Set RSAX = connaX.Execute(strQuery)



			If RSAX.EOF Then


				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO"
				session("SGU_mail") = "NO"
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel") ="9"

			Else

				session("SGU_userid") = Legajo
				session("SGU_nombre") = RSAX("name")
				session("SGU_mail") = RSAX("mail")
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
				session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
				Exit Do
			End If

			RSAX.CLOSE
			Set RSAX = Nothing
			connaX.CLOSE
			Set connaX = Nothing
			rsdom.MOVENEXT
		Loop
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnDOM)
	End If
End Function



Function getinfoUserAD_COR(legajo)

	If Len(Trim(LEGAJO & "")) = 0 Then

		session("SGU_userid") = "--" & Legajo & "-"
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel") ="9"
	Else
		Set cnDOM = ofv.conectar(ofv.strconn0)

		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
		
		

		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

		Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
			sldap =rsdom("raizldap")


			usernx =  rsdom("usrldap")
			passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

			Set connaX = server.CreateObject("ADODB.Connection")


			connaX.provider = "ADSDSOObject"
			connaX.Properties("User ID") = usernx
			connaX.Properties("Password") = passx
			connaX.Properties("Encrypt Password") = True
			' response.write userN & "----" & pass & "---" & strBase
			connaX.OPEN

			strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description"

			strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


			strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
			'response.write strQuery
			Set RSAX = connaX.Execute(strQuery)



			If RSAX.EOF Then


				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO"
				session("SGU_mail") = "NO"
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel") ="9"

			Else

				session("SGU_userid") = Legajo
				session("SGU_nombre") = RSAX("name")
				session("SGU_mail") = RSAX("mail")
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
				session("SGU_descripuesto") = "NO"


				If IsNull(rsax("description").Value) Then
				    valorDescription = "" ' Está vacío en Active Directory
				ElseIf IsArray(rsax("description").Value) Then
				    ' Es un arreglo. Usamos Join para convertirlo a texto si tiene varios valores,
				    ' o simplemente tomamos la posición (0)
				    valorDescription = Join(rsax("description").Value, " ")
				Else
				    valorDescription = rsax("description").Value & ""
				End If
				session("SGU_codigopuesto") = getperfil(valorDescription)' & "2222222222222222" 'getperfilExtBanelco(rsax("description"))
				 session("SGU_descripuesto") = getperfildescri( session("SGU_codigopuesto"))
				session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
				Exit Do
			End If

			RSAX.CLOSE
			Set RSAX = Nothing
			connaX.CLOSE
			Set connaX = Nothing
			rsdom.MOVENEXT
		Loop
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnDOM)
	End If
End Function


Function getinfoUserAD_GP(legajo)

	If Len(Trim(LEGAJO & "")) = 0 Then

		session("SGU_userid") = "--" & Legajo & "-"
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel") ="9"
	Else
		Set cnDOM = ofv.conectar(ofv.strconn0)

		'    Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
		If Session("EMPMUN")  = "S" Then
			If CStr(SESSION("CLIEMP") &"") =CStr("0") Then SESSION("CLIEMP")=session("usr_dominio")
			'Strsqlq = "Select * From dominios WHERE descripcion='" & SESSION("DOMINIOX") & "' and cliente=" & SESSION("CLIEMP")
			Strsqlq = "Select * From dominios WHERE descripcion='" & SESSION("DOMINIOX") & "'"
		Else
			Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
		End If

		'      response.write strsqlq
		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

		Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
			sldap =rsdom("raizldap")


			usernx =  rsdom("usrldap")
			passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))
			' RESPONSE.WRITE passx
			Set connaX = server.CreateObject("ADODB.Connection")


			connaX.provider = "ADSDSOObject"
			connaX.Properties("User ID") = usernx
			connaX.Properties("Password") = passx
			connaX.Properties("Encrypt Password") = True
			' response.write userN & "----" & pass & "---" & strBase
			connaX.OPEN

			strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

			strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


			strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

			Set RSAX = connaX.Execute(strQuery)



			If RSAX.EOF Then


				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO"
				session("SGU_mail") = "NO"
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel") ="9"
				session("SGU_site")=1

			Else

				session("SGU_userid") = Legajo
				If rsdom("id")=7 Then
					session("SGU_nombre") = RSAX("displayName")
				Else
					session("SGU_nombre") = RSAX("name")
				End If
				session("SGU_mail") = RSAX("mail")
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
				session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))

				'session("SGU_site")=rsdom("cliente") 'session("usr_dominio")

				If SESSION("SITECORP")="1" Then
					StrSqlq="SELECT SITE FROM USUARIOS WHERE ID='" & SESSION("USUARIO")&"'"
					Set rsSITE = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
					If Not rsSITE.EOF Then

						session("SGU_site")=rsSITE(0)
					Else
						session("SGU_site")=session("usr_dominio")
					End If
					Call OFV.CERRARCONSULTA(rsSITE)
				Else
					session("SGU_site")=session("usr_dominio")
				End If

				'  RESPONSE.WRITE session("SGU_site")
				Exit Do
			End If

			RSAX.CLOSE
			Set RSAX = Nothing
			connaX.CLOSE
			Set connaX = Nothing
			rsdom.MOVENEXT
		Loop
		'response.write strQuery & session("SGU_site")
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnDOM)
	End If
End Function

Function getinfoUserADEX_GP(legajo,DATOSE)

	' DATOSE codig de empresa en DATOSE(0)
	'chr(9) para seperar datos por si hay que agregar mas paramentros

	Dim cempre
	dim DATOSEx
	cempre=0
	DATOSEx=DATOSE
	If Trim(DATOSEx) = "" Then
	Else
		DATOSEx=Split(Trim(DATOSEx & ""),Chr(9))

		If Len(DATOSEx(0) & "")  > 0 Then
			cempre=DATOSEx(0)
		End If
	End If


	If Len(Trim(LEGAJO & "")) = 0 Then

		session("SGU_userid") = "--" & Legajo & "-"
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel") ="9"
	Else
		Set cnDOM = ofv.conectar(ofv.strconn0)


		If cempre > 0 Then
			Strsqlq = "Select * From dominios WHERE  cliente=" & cempre
		Else
			'    Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
			If Session("EMPMUN")  = "S" Then
				If CStr(SESSION("CLIEMP") &"") =CStr("0") Then SESSION("CLIEMP")=session("usr_dominio")
				'Strsqlq = "Select * From dominios WHERE descripcion='" & SESSION("DOMINIOX") & "' and cliente=" & SESSION("CLIEMP")
				Strsqlq = "Select * From dominios WHERE descripcion='" & SESSION("DOMINIOX") & "'"
			Else
				Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
			End If
		End If
		'      response.write strsqlq
		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

		Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
			sldap =rsdom("raizldap")


			usernx =  rsdom("usrldap")
			passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))
			' RESPONSE.WRITE passx
			Set connaX = server.CreateObject("ADODB.Connection")


			connaX.provider = "ADSDSOObject"
			connaX.Properties("User ID") = usernx
			connaX.Properties("Password") = passx
			connaX.Properties("Encrypt Password") = True
			' response.write userN & "----" & pass & "---" & strBase
			connaX.OPEN

			strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

			strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


			strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

			Set RSAX = connaX.Execute(strQuery)



			If RSAX.EOF Then


				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO"
				session("SGU_mail") = "NO"
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel") ="9"
				session("SGU_site")=1

			Else

				session("SGU_userid") = Legajo
				If rsdom("id")=7 Then
					session("SGU_nombre") = RSAX("displayName")
				Else
					session("SGU_nombre") = RSAX("name")
				End If
				session("SGU_mail") = RSAX("mail")
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0" 'getsectorExt(rsax("department")&"")
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO " 'getperfilExtBanelco(rsax("title"))
				session("SGU_idnivel")="9" 'getnivelExtBanelco(ucase(session("SGU_idnivel")))
				session("SGU_site")=rsdom("cliente") 'session("usr_dominio")
				'  RESPONSE.WRITE session("SGU_site")
				Exit Do
			End If

			RSAX.CLOSE
			Set RSAX = Nothing
			connaX.CLOSE
			Set connaX = Nothing
			rsdom.MOVENEXT
		Loop
		'response.write strQuery & session("SGU_site")
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnDOM)
	End If
End Function


Function getinfoUserAD_BCR(legajo)

	If Len(Trim(LEGAJO & "")) = 0 Then

		session("SGU_userid") = "--" & Legajo & "-"
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel") ="9"
	Else
		Set cnDOM = ofv.conectar(ofv.strconn0)

		Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

		Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

		Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
			sldap =rsdom("raizldap")


			usernx =  rsdom("usrldap")
			passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

			Set connaX = server.CreateObject("ADODB.Connection")


			connaX.provider = "ADSDSOObject"
			connaX.Properties("User ID") = usernx
			connaX.Properties("Password") = passx
			connaX.Properties("Encrypt Password") = True
			' response.write userN & "----" & pass & "---" & strBase
			connaX.OPEN

			strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,department,description,displayName"

			strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


			strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
			'response.write strQuery
			Set RSAX = connaX.Execute(strQuery)



			If RSAX.EOF Then


				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO"
				session("SGU_mail") = "NO"
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel") ="9"

			Else

				session("SGU_userid") = Legajo
				session("SGU_nombre") = RSAX("displayName")
				session("SGU_mail") = RSAX("mail")
				session("SGU_descrisector") = "NO"
				session("SGU_idcostcenter") = getsectorExtbcr(rsax("department")&"")
				session("SGU_descripuesto") = "NO"
				session("SGU_codigopuesto") ="NO" 'getperfilExtbcr(rsax("title"))
				session("SGU_idnivel")=getnivelExtBCR(rsax("title")&"")
				Exit Do
			End If

			RSAX.CLOSE
			Set RSAX = Nothing
			connaX.CLOSE
			Set connaX = Nothing
			rsdom.MOVENEXT
		Loop
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnDOM)
	End If


End Function






Function getinfoUserADBC(legajo)


	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"

	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof)
		sldap =rsdom("raizldap")


		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		Set RSAX = connaX.Execute(strQuery)



		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"

		Else

			Set Cnpeople =ofv.conectar(ofv.strconn7)
			'legajoZ=Mid(legajo,3,Len(legajo))
			' 21/2/2019 por pedido del banco
            legajoz=legajo
			strsqlCn = "select ALTER_EMPLID  ,FIRST_NAME,MIDDLE_NAME,LAST_NAME, DEPTID  ,JOBCODE   from PS_BC_NORMAS_VW  where UPPER(ALTER_EMPLID)='" & UCASE(legajoZ) &"' ORDER BY NAME1"

			Set CnViewPeople=ofv.crearconsultaEx(strsqlCn ,Cnpeople ,1,volver)

			If Not CnViewPeople.eof Then

				session("SGU_userid") = Legajo
				session("SGU_nombre") =CnViewPeople("FIRST_NAME") & "&nbsp;" &  CnViewPeople("MIDDLE_NAME") & " " & CnViewPeople("LAST_NAME")
				session("SGU_mail") = ChangeMail(RSAX("mail"))'RSAX("mail")
				session("SGU_descrisector") = "NO"
				SDEPTID = CnViewPeople("DEPTID") & ""
				SDEPTID=Trim(SDEPTID)
				session("SGU_idcostcenter") = getsectorExt(SDEPTID) 'CnViewPeople("DEPTID")
				session("SGU_descripuesto") = ""
				SXJOB=CnViewPeople("JOBCODE") & ""
				SXJOB=Trim(SXJOB)
				session("SGU_codigopuesto") = SXJOB
				session("SGU_idnivel")      = getnivelExt(SXJOB)
			Else

				session("SGU_userid") = Legajo
				session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION '& strsqlCn
				session("SGU_mail") = "NO@disponible.com"
				session("SGU_descrisector") = "NO disponible"
				session("SGU_idcostcenter") = "0"
				session("SGU_descripuesto") = "NO disponible"
				session("SGU_codigopuesto") = "NO"
				session("SGU_idnivel")      ="9"
			End If

			Call ofv.cerrarconsulta(CnViewPeople)
			Call ofv.cerrarconn(Cnpeople)

			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)

End Function


Function GetUserInsectBC(XGU_SECT)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	If IsNumeric(XGU_SECT) Then
		WXGU_SECT = XGU_SECT
	Else
		WXGU_SECT = 1
	End If

	Set xcn = ofv.conectar(ofv.strconn2)

	xStrSql = "Select codinter From eco WHERE eco = 0" & WXGU_SECT
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")

	If Not xrs.eof Then WXGU_SECT= xrs(0)

	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

	Set Cnpeople=ofv.conectar(ofv.strconn7)
	strsqlCn= "select ALTER_EMPLID ,FIRST_NAME,MIDDLE_NAME,LAST_NAME from PS_BC_NORMAS_VW  where DEPTID='" & WXGU_SECT &"' ORDER BY NAME1"
	Set CnViewPeople=ofv.crearconsultaEx(strsqlCn ,Cnpeople ,1,volver)

	Do While Not CnViewPeople.EOF

		SALTER_EMPLID=CnViewPeople("ALTER_EMPLID") & ""
		SALTER_EMPLID=Trim(SALTER_EMPLID)
		legajo=getInfouserBCEx(SALTER_EMPLID)

		session("SGU_UIGuserid") =  session("SGU_UIGuserid") & xgusep & legajo
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & CnViewPeople("FIRST_NAME") & " " &  CnViewPeople("MIDDLE_NAME") & " " & CnViewPeople("LAST_NAME")
		xgusep = Chr(9)
		CnViewPeople.movenext
	Loop

	Call ofv.cerrarconsulta(CnViewPeople)
	Call ofv.cerrarconn(Cnpeople)


End Function

'------------------------------------------------------------------------------------------------------------

Function ValDominio(dominio)

	session("ldap")=""
	If session("ldap") = "" Then
		session("usr_dominio")=1
		'RESPONSE.WRITE ofv.strconn0
		Set cnDOM = ofv.conectar(ofv.strconn0)

		Strsql = "Select * From dominios WHERE DESCRIPCION='" & dominio & "'"
		'RESPONSE.WRITE STRSQL
		Set rsdom = ofv.crearconsultaEx(StrSql, cnDOM, 1, parametros)
		If Not rsdom.EOF Then
			session("usr_dominio")=rsdom("cliente")

			session("ldap") =rsdom("raizldap")
			session("usrldap") =  rsdom("usrldap")
			If IsNumeric(rsdom("pwdldap")&"") Then
				session("pwdldap") = ofv.GetCript("D", CStr(rsdom("pwdldap")&"")) 'rsdom("pwdldap") '
			Else
				session("pwdldap") = rsdom("pwdldap")
			End If
		End If
		Call ofv.cerrarconsulta(rsdom)
		Call ofv.cerrarconn(cnDOM)

	End If
	' RESPONSE.WRITE  session("usr_dominio") 'strsql  & "---" & dominio 'session("ldap")

End Function

'--------------------------------------------------------------------------------------------------------------

Function FullNameGrupo(idgrupo,strbase1)
	Dim rsad
	Dim cnad


	Set cnad = server.CreateObject("ADODB.Connection")


	cnad.provider = "ADSDSOObject"
	cnad.Properties("User ID") = userN
	cnad.Properties("Password") = pass
	cnad.Properties("Encrypt Password") = True

	cnAd.OPEN

	strAttributes = "name,distinguishedName"

	strFilter = "(&(objectCategory=group) " & "(Name=" & idgrupo &"))"


	strQuery = strBase1 & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery
	Set rsad = cnAd.Execute(strQuery)

	If Not rsad.EOF Then
		FullNameGrupo = rsad("distinguishedName")
	Else
		FullNameGrupo = idgrupo
	End If

	rsad.CLOSE
	Set rsad = Nothing
	cnAd.CLOSE
	Set cnad = Nothing

End Function


Function FullNameGrupoEx(idgrupo,strbase1,userN,pass)
	Dim rsad
	Dim cnad


	Set cnad = server.CreateObject("ADODB.Connection")


	cnad.provider = "ADSDSOObject"
	cnad.Properties("User ID") = userN
	cnad.Properties("Password") = ofv.GetCript("D", CStr( pass &""))
	cnad.Properties("Encrypt Password") = True

	cnAd.OPEN

	strAttributes = "name,distinguishedName"

	strFilter = "(&(objectCategory=group) " & "(Name=" & idgrupo &"))"


	strQuery = strBase1 & ";" & strFilter & ";" & strAttributes & ";subtree"
	'  response.write strQuery
	Set rsad = cnAd.Execute(strQuery)

	If Not rsad.EOF Then
		FullNameGrupoEx = rsad("distinguishedName")
	Else
		FullNameGrupoEx = idgrupo
	End If

	rsad.CLOSE
	Set rsad = Nothing
	cnAd.CLOSE
	Set cnad = Nothing

End Function

'----------RBAC------------------------------------------------RBAC-----------------------------------------------------------------


Function getinfoUserADBCRED(legajo)

	Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

	Dirc=Dirc  & "coressdll\getuser.xml"

	Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")

	xml4=leerArc(dirc)
	xml4=Replace(XML4,"XUSUARIO",legajo)
	xml4=Replace(XML4,"XAPP",session("USRPRV"))


	xmlHttp2.open "POST", sRBAC, False
	xmlHttp2.send (xml4 )

	Set xml3 =xmlHTTP2.responsexml

	' Set xmldoc=xml3.selectSingleNode("//response")
	'Set z1=xmldoc.selectSingleNode("//error")
	' serror= cbool(z1.text)

	Set xmldoc=xml3.selectSingleNode("//userData")
	If Not (xmldoc Is Nothing)  Then
		session("SGU_userid") = legajo

		Set z1=xmldoc.selectSingleNode("//dsFirstName")
		nombre=z1.text
		Set z2=xmldoc.selectSingleNode("//dsLastName")
		apellido=z2.text

		session("SGU_nombre") = nombre & " " & apellido
		session("SGU_mail") = GETMAIL(legajo)
		session("SGU_descrisector") = "MySect"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel")      ="8"
	Else
		session("SGU_userid") = Legajo
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel")      ="8"
	End If

	Set xmlHTTP2=Nothing

End Function

'-------------------------------------------------------------------------------------

Function getinfoUserWBCRED(legajo)

	Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

	Dirc=Dirc  & "coressdll\getuser.xml"

	Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")

	xml4=leerArc(dirc)
	xml4=Replace(XML4,"XUSUARIO",legajo)
	xml4=Replace(XML4,"XAPP",session("USRPRV"))


	xmlHttp2.open "POST", sRBAC, False
	xmlHttp2.send (xml4 )

	Set xml3 =xmlHTTP2.responsexml

	' Set xmldoc=xml3.selectSingleNode("//response")
	'Set z1=xmldoc.selectSingleNode("//error")
	' serror= ucase(z1.text&"")

	' if serror="FALSE" THEN
	Set xmldoc=xml3.selectSingleNode("//userData")
	If Not (xmldoc Is Nothing)  Then

		session("SGU_userid") = legajo

		Set z1=xmldoc.selectSingleNode("//dsFirstName")
		nombre=z1.text
		Set z2=xmldoc.selectSingleNode("//dsLastName")
		apellido=z2.text

		session("SGU_nombre") = nombre & " " & apellido
		session("SGU_mail") = GETMAIL(legajo)
		session("SGU_descrisector") = "MySect"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel")      ="8"
	Else
		session("SGU_userid") = Legajo
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel")      ="8"
	End If

	Set xmlHTTP2=Nothing

End Function

'-------------------------------------------------------------------------------------

Function getGroupsUserBCRED(legajo)

	session("SGU_grupos")=""
	xgusep=""
	Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

	Dirc=Dirc  & "coressdll\getroles.xml"

	Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")

	xml4=leerArc(dirc)
	xml4=Replace(XML4,"XUSUARIO",legajo)
	xml4=Replace(XML4,"XAPP",session("USRPRV"))


	xmlHttp2.open "POST", sRBAC, False
	xmlHttp2.send (xml4)

	Set xml3 =xmlHTTP2.responsexml


	Set xmldoc=xml3.selectSingleNode("//return")


	Set nList = xmldoc.selectNodes("roles")

	For Each n In nList
		session("SGU_grupos") = session("SGU_grupos") & xgusep  & UCase(n.text &"")
		xgusep = Chr(9)
	Next

	Set xmlHTTP2=Nothing
End Function


Function getInfouserPstBANELCOEX(Legajo,IDPST,tipo)


	session("SGU_userid")      = Legajo
	session("SGU_nombre")      = "NO"
	session("SGU_mail")	   = "NO"
	session("SGU_descrisector")= "NO"
	session("SGU_idcostcenter")= "0"
	session("SGU_descripuesto")= "NO"
	session("SGU_codigopuesto")= "NO"
	session("SGU_direccion")   ="NO"
	session("SGU_telefono")   ="NO"
	session("SGU_DEPTO")   =""
	session("SGU_IMG")=""


	Select Case tipo


		Case "USUARIO"



		Set xcnPSB = ofv.conectar(ofv.strconn2)

		session("SGU_descrisector")="NO"
		session("SGU_descripuesto")="NO"

		XSQL="    Select C.DESCRI,C1.DESCRI From PUESTOREL "
		XSQL=XSQL & "  INNER JOIN LISTADEPUESTOS C ON(C.OBJID=PUESTOREL.IDPST) "
		XSQL=XSQL & "   	   LEFT JOIN ECO C1 ON (C1.ECO=C.SECTOR) "
		XSQL=XSQL & " WHERE PUESTOREL.idpst =" &  IDPST

		Set xrs2 = OFV.crearconsultaEx(XSQL, xcnPSB, 1, "")

		xx=getinfouser(legajo)

		session("SGU_codigopuesto")= IDPST
		session("SGU_DEPTO")=xrs2(1)




		Call ofv.cerrarconsulta(xrs2)

		Call ofv.cerrarconn(xcnPSB)






		Case "PUESTO"

		xgusep                   = ""
		session("SGU_UIGuserid") = ""
		session("SGU_UIGdescri") = ""

		Select Case GTUcase

			Case "I"

			Set xcn = ofv.conectar(ofv.strconn2)

			xStrSql = "Select cod From relapuestos WHERE idpst = '" & Legajo
			Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
			If Not xRS.EOF Then

				Do While Not XRS.EOF

					session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
					xx=getinfouser(xrs(0) & "")
					session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & session("SGU_nombre")

					xgusep                   = Chr(9)

					XRS.MOVENEXT
				Loop

			Else

			End If

			Call ofv.cerrarconsulta(xrs)
			Call ofv.cerrarconn(xcn)
		End Select

	End Select
End Function
'---------------------------------------------------------------------------


'---------------------------------------------------------------------------

Function getInfouserPstBcEX(Legajo,IDPST,tipo)


	session("SGU_userid")      = Legajo
	session("SGU_nombre")      = "NO"
	session("SGU_mail")	   = "NO"
	session("SGU_descrisector")= "NO"
	session("SGU_idcostcenter")= "0"
	session("SGU_descripuesto")= "NO"
	session("SGU_codigopuesto")= "NO"
	session("SGU_direccion")   ="NO"
	session("SGU_telefono")   ="NO"
	session("SGU_DEPTO")   =""
	session("SGU_IMG")=""


	Select Case tipo


		Case "USUARIO"

		Set xcnPS = ofv.conectar(ofv.strconn7)
		Set xcnPSB = ofv.conectar(ofv.strconn2)



		xStrSql="select * from PS_BC_NORMAS_VW  where	RTRIM(UPPER(ALTER_EMPLID))='" & UCase(Legajo & "") &"' ORDER BY NAME1"

		' response.write xStrSql

		session("SGU_descrisector")="NO"
		session("SGU_descripuesto")="NO"
		Set xrs = OFV.crearconsultaEx(xStrSql, xcnPS, 1, "")
		If Not xRS.EOF Then

			XSQL="    Select C.DESCRI,C1.DESCRI From PUESTOREL "
			XSQL=XSQL & "  INNER JOIN LISTADEPUESTOS C ON(C.OBJID=PUESTOREL.IDPST) "
			XSQL=XSQL & "   	   LEFT JOIN ECO C1 ON (C1.ECO=C.SECTOR) "
			XSQL=XSQL & " WHERE PUESTOREL.idpst =" &  IDPST
			'RESPONSE.WRITE XSQL
			Set xrs2 = OFV.crearconsultaEx(XSQL, xcnPSB, 1, "")

			session("SGU_userid")      = Trim(xrs("ALTER_EMPLID")&"")
			session("SGU_nombre")      = Replace(Trim(xrs("NAME1")&""),"'"," ")
			session("SGU_mail")	   = Trim(xrs("EMAILID") &"")
			'session("SGU_descrisector")= trim(xrs2(1)&"")&"&nbsp;"

			If Not XRS2.EOF Then
				session("SGU_descrisector")= Trim(xrs2(1)&"")&"&nbsp;"
				session("SGU_descripuesto")= Trim(xrs2(0) &"")&"&nbsp;"
			Else
				session("SGU_descrisector")="NO"
				session("SGU_descripuesto")="NO"


			End If
			session("SGU_idcostcenter")= xrs(3)
			'session("SGU_descripuesto")= trim(xrs2(0) &"")&"&nbsp;"
			session("SGU_codigopuesto")= IDPST
			session("SGU_direccion")   =Trim(xrs("DESCR200")&"")
			session("SGU_telefono")   =Trim(xrs("DESCR254_MIXED")&"")
			session("SGU_DEPTO")   =Trim(xrs("DESCR") &"-" & xrs("DESCR1")&"")
			ximg=Trim(xrs("BC_DESCR250_1") &"")
			session("SGU_IMG")   =Right(ximg,Len(ximg)- InStrRev(ximg,"\"))

		Else
			session("SGU_userid")      = Legajo
			session("SGU_nombre")      = "NO"
			session("SGU_mail")	   = "NO"
			session("SGU_descrisector")= "NO"
			session("SGU_idcostcenter")= "0"
			session("SGU_descripuesto")= "NO"
			session("SGU_codigopuesto")= "NO"
			session("SGU_direccion")   ="NO"
			session("SGU_telefono")   ="NO"
			session("SGU_DEPTO")   =""
			session("SGU_IMG")=""
		End If

		Call ofv.cerrarconsulta(xrs)
		Call ofv.cerrarconn(xcnPS)
		Call ofv.cerrarconn(xcnPSB)


		Case "PUESTO"

		xgusep                   = ""
		session("SGU_UIGuserid") = ""
		session("SGU_UIGdescri") = ""

		Select Case GTUcase

			Case "I"

			Set xcn = ofv.conectar(ofv.strconn2)

			xStrSql = "Select cod From relapuestos WHERE idpst = '" & Legajo
			Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")
			If Not xRS.EOF Then

				Do While Not XRS.EOF

					session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & xrs(0)
					xx=getinfouser(xrs(0) & "")
					session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & session("SGU_nombre")

					xgusep                   = Chr(9)

					XRS.MOVENEXT
				Loop

			Else

			End If

			Call ofv.cerrarconsulta(xrs)
			Call ofv.cerrarconn(xcn)
		End Select

	End Select
End Function
'---------------------------------------------------------------------------






Function getUsrInTEXTBCEX(XGU_TEXT)



	Set xcn = ofv.conectar(ofv.strconn7)

	XGU_TEXT=Replace(UCase(XGU_TEXT & "" ),"'","''")

	xStrSql="select ALTER_EMPLID,NAME1 from PS_BC_NORMAS_VW  where upper(NAME1) LIKE '" & XGU_TEXT &"%' ORDER BY NAME1"
	' response.write xStrSql
	Set xrs = OFV.crearconsultaEx(xStrSql, xcn, 1, "")




	xgusep                   = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""




	If Not xRS.EOF Then

		Do While Not XRS.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & Trim(xrs(0)&"")

			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & Replace(Trim(XRS(1)&""),"'" ," ")

			xgusep                   = Chr(9)

			XRS.MOVENEXT
		Loop

	Else

	End If

	Call ofv.cerrarconsulta(xrs)
	Call ofv.cerrarconn(xcn)

End Function
'---------------------------------------------------------------------------

Function getUsrIngroupBCRED(GRUPO)


	session("SGU_grupos")=""
	xgusep=""



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)

	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)

	Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

	Dirc=Dirc  & "coressdll\USERWITHROLEv1.xml"

	Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")

	xml4=leerArc(dirc)
	xml4=Replace(XML4,"XROLE",GrupoADX)
	xml4=Replace(XML4,"XAPP",session("USRPRV"))


	xmlHttp2.open "POST", sRBAC, False
	xmlHttp2.send (xml4)

	Set xml3 =xmlHTTP2.responsexml


	Set xmldoc=xml3.selectSingleNode("//return")


	Set nList = xmldoc.selectNodes("users")

	For Each n In nList
		session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & n.text
		xxx=getinfoUserADBCRED(n.text)
		session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & session("SGU_nombre")
		xgusep = Chr(9)
	Next

	Set xmlHTTP2=Nothing




End Function

'--------------------------------------------------------------------------

Function GETMAIL(LEGAJO)

	Dirc=  Request.ServerVariables("APPL_PHYSICAL_PATH")

	Dirc=Dirc  & "coressdll\getmailLdap.xml"

	Set xmlHttp2=SERVER.CreateObject("Microsoft.XMLHTTP")

	xml4=leerArc(dirc)
	xml4=Replace(XML4,"XUSUARIO",legajo)



	xmlHttp2.open "POST", sMailLdap, False
	xmlHttp2.send (xml4)

	Set xml3 =xmlHTTP2.responsexml

	Set xmldoc=xml3.selectSingleNode("//return")
	Set n = xmldoc.selectSingleNode("mail")
	If Not (n Is Nothing) Then
		GETMAIL=n.text
	Else
		getmail="NO"
	End If

	Set xmlHTTP2=Nothing

End Function


Function getsectorExt(Sdato)

	'Sdato=mid(sdato,1,11)

	If SDATO & "" ="" Then
		getsectorExt="NO"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select * From eco WHERE codinter='" & sdato & "'"

		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getsectorExt=rsdomeco(0)
		Else
			getsectorExt="0" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function





Function getnivelExt(Sdato)

	' Sdato=RTRIM(Sdato & "")
	Set cnext = ofv.conectar(ofv.strconn0)

	Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  >=0" & sdato

	Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

	If Not rsdomeco.eof Then
		getnivelExt=rsdomeco(0)
	Else
		getnivelExt ="9" 'Sdato
	End If
	Call ofv.cerrarconsulta(rsdomeco)
	Call ofv.cerrarconn(cnext)

End Function

Function getnivelExtBanelco(Sdato)

	Sdato=RTrim(Sdato & "")
	If  SDATO="" Then
		getnivelExtBanelco="9"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  ='" & sdato & "'"

		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getnivelExtBanelco=rsdomeco(0)
		Else
			getnivelExtBanelco ="9" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function


Function getperfilExtBanelco(Sdato)

	Sdato=RTrim(Sdato & "")
	If SDATO="" Then
		getperfilExtBanelco="NO"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select tipoPerfil  From TipoPerfil WHERE CODpuestoE  ='" & sdato & "'"
		'RESPONSE.WRITE Strsqle3
		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getperfilExtBanelco=rsdomeco(0)
		Else
			getperfilExtBanelco ="NO" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function

Function getperfil(Sdato)

	Sdato=RTrim(Sdato & "")

	If SDATO="" Then
		getperfil="NO"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select tipoPerfil  From TipoPerfil WHERE CODpuestoE  ='" & sdato & "'"
		'RESPONSE.WRITE Strsqle3
		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getperfil=rsdomeco(0)
		Else
			getperfil ="NO" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function

Function getperfildescri(Sdato)

	Sdato=RTrim(Sdato & "")

	If SDATO="" Then
		getperfildescri="NO"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select descri  From TipoPerfil WHERE tipoperfil  ='" & sdato & "'"
		'RESPONSE.WRITE Strsqle3
		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getperfildescri=rsdomeco(0)
		Else
			getperfildescri ="NO" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function



Function getInfouserBCEx(legajo)

	' legajo=RTRIM(legajo &"")
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=*" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then
			getInfouserBCEx=legajo
		Else
			getInfouserBCEx=RSAX("samAccountName")
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function

Function ChangeMail(Smail)
	Smail = Smail & ""
	Armail = Split(Smail, "@")

	If UBound(Armail) <> -1 Then
		ChangeMail = Armail(0) & "@ciudad.bco"
	Else
		ChangeMail = "NO"
	End If

End Function


Function GrupoDelUsrLoguin(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("GrupoDelLegajo")

	strSoapReq = strSoapReq & "<GrupoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</GrupoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("GrupoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "GrupoDelLegajo"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	Set strResult = oHttp.responseXML
	xgusep=""
	GrupoDelUsrLoguin=""
	Set nList = strResult.selectNodes("//Grupo")

	Dim iLoop
	For Each n In nList
		If n.text <>"" Then
			GrupoDelUsrLoguin = GrupoDelUsrLoguin & xgusep  &  n.text &""
			xgusep = Chr(9)
		End If

	Next

	Set oHttp=Nothing
	Set strResult=Nothing
End Function


Function LegajosSector(sector)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("LegajosSector")

	strSoapReq = strSoapReq & "<LegajosSectorRequest>" & GenerateSoapParameter("Sector", sector) & "</LegajosSectorRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("LegajosSector")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "LegajosSector"
	oHttp.send strSoapReq
	'	response.write oHttp.responseText
	Set strResult = oHttp.responseXML
	Set nList = strResult.selectNodes("//Empleados/item")
	xgusep=""
	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	Dim iLoop
	For Each n In nList


		Legajo=n.selectSingleNode("Legajo").text &""

		nombre=n.selectSingleNode("Nombre").text & ""

		If Legajo <>"" Then
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & Legajo '& objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & nombre
			xgusep = Chr(9)
		End If

	Next



	session("SGU_UIGdescri")=Konvert(session("SGU_UIGdescri"))

End Function

Function GrupoDelLegajo(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("GrupoDelLegajo")

	strSoapReq = strSoapReq & "<GrupoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</GrupoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("GrupoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "GrupoDelLegajo"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	Set strResult = oHttp.responseXML
	xgusep=""
	session("SGU_grupos")=""
	Set nList = strResult.selectNodes("//Grupo")

	Dim iLoop
	For Each n In nList
		If n.text <>"" Then
			session("SGU_grupos") = session("SGU_grupos") & xgusep  &  n.text &""
			xgusep = Chr(9)

		End If

	Next

	Set oHttp=Nothing
	Set strResult=Nothing
End Function

Function Konvert(Sin)
	Dim oIn
	Set oIn = SERVER.CreateObject("ADODB.Stream")
	oIn.Open
	oIn.CharSet = "ISO-8859-1"
	oIn.WriteText Sin
	oIn.Position = 0
	oIn.CharSet =  "UTF-8"
	Konvert = oIn.ReadText
	oIn.Close
	Set OIN=Nothing
End Function

Function LegajosGrupo(grupo)
	' Create the http request text
	Dim strSoapReq

	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		grupo=rsADX(0)

	End If

	Call ofv.cerrarconsulta(rsADX)


	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("LegajosGrupo")

	strSoapReq = strSoapReq & "<LegajosGrupoRequest>" & GenerateSoapParameter("Grupo", grupo) & "</LegajosGrupoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("LegajosGrupo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "LegajosGrupo"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	Set strResult = oHttp.responseXML

	Set nList = strResult.selectNodes("//Empleados/item")
	xgusep=""
	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	Dim iLoop
	For Each n In nList
		Legajo=n.selectSingleNode("Legajo").text &""

		nombre=n.selectSingleNode("Nombre").text & ""

		If Legajo <>"" Then
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & Legajo '& objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & nombre
			xgusep = Chr(9)



		End If

	Next

	session("SGU_UIGdescri")=Konvert(session("SGU_UIGdescri"))
	Set oHttp=Nothing
	Set strResult=Nothing
End Function

Function Loguin(p1,p2)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("Login")

	strSoapReq = strSoapReq & "<LoginRequest>" & GenerateSoapParameter("Legajo", p1)
	strSoapReq = strSoapReq & GenerateSoapParameter("Identificacion", p2) &   "</LoginRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("Login")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttp
	Dim strResult
	Set oHttp = SERVER.CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "Login"
	oHttp.send strSoapReq
	'response.write oHttp.responseText
	Set strResult = oHttp.responseXML


	ErrorT=strResult.selectSingleNode("//Error ").Text &""

	Loguin=Konvert(ErrorT)
	Set oHttp=Nothing
	Set strResult=Nothing
End Function





Function GetResult(ByVal responseText, ByVal resultParam)

	Dim oXml
	Set oXml = CreateObject("Msxml2.DOMDocument")
	oXml.Async = True
	oXml.LoadXml responseText

	Dim strPath
	strPath = "/*/*/*/" & resultParam
	Dim oNode
	Set oNode = oXml.documentElement.SelectSingleNode(strPath)
	GetResult = oNode.Text

End Function


Function GenerateSoapBodyStart()

	Dim strSoap
	strSoap = "<?xml version=""1.0"" encoding=""utf-8""?>"
	strSoap = strSoap & "<soap12:Envelope "
	strSoap = strSoap & "xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" "
	strSoap = strSoap & "xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" "
	strSoap = strSoap & "xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope""> "
	strSoap = strSoap & "<soap12:Body>"
	GenerateSoapBodyStart = strSoap

End Function

Function GenerateSoapBodyEnd()

	Dim strSoap
	strSoap = "</soap12:Body>"
	strSoap = strSoap & "</soap12:Envelope>"
	GenerateSoapBodyEnd = strSoap

End Function

Function GenerateSoapFunctionCallStart(ByVal strFunction)

	Dim strSoap
	strSoap = "<" & strFunction & " xmlns=""" & nsUrl & """>"
	GenerateSoapFunctionCallStart = strSoap

End Function

Function GenerateSoapFunctionCallEnd(ByVal strFunction)

	Dim strSoap
	strSoap = "</" & strFunction & ">"
	GenerateSoapFunctionCallEnd = strSoap

End Function

Function GenerateSoapParameter(ByVal strParam, ByVal strValue)

	Dim strSoap
	strSoap = "<" & strParam & ">"
	strSoap = strSoap & strValue
	strSoap = strSoap & "</" + strParam & ">"
	GenerateSoapParameter = strSoap

End Function





Function LOGUINEx(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("LoginLegajo")

	strSoapReq = strSoapReq & "<LoginLRequest >" & GenerateSoapParameter("Legajo", legajo) & "</LoginLRequest >"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("LoginLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()
	' RESPONSE.WRITE strSoapReq
	Dim oHttpstrSoapReq
	Dim strResult
	Set oHttp = server.CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "LoginLegajo"
	oHttp.send strSoapReq

	Set strResult = oHttp.responseXML



	If Not (strResult Is Nothing)  Then


		LOGUINEx=Konvert(strResult.selectSingleNode("//Error").Text &"")
		If LOGUINEx ="" Then
			LOGUINEx="OK"
			Session("Usuario")   = Legajo
			Call InfousrLOGUIN(Legajo)
		End If

	Else

		LOGUINEx="NO"

	End If


	Set oHttp=Nothing
	Set strResult=Nothing



End Function

Function InfoUsrLOGUIN(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("InfoDelLegajo")

	strSoapReq = strSoapReq & "<InfoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</InfoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("InfoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()
	' RESPONSE.WRITE strSoapReq
	Dim oHttpstrSoapReq
	Dim strResult
	Set oHttp = server.CreateObject("Msxml2.XMLHTTP")
	oHttp.open "POST", URLRiva, False
	oHttp.setRequestHeader "Content-Type", "text/xml"
	oHttp.setRequestHeader "SOAPAction", URLRiva & "InfoDelLegajo"
	oHttp.send strSoapReq

	Set strResult = oHttp.responseXML

	If Not (strResult Is Nothing)  Then
		session("SGU_userid") = legajo

		UsuarioNombre=Konvert(strResult.selectSingleNode("//UsuarioNombre").Text &"")
		UsuarioMail=Konvert(strResult.selectSingleNode("//UsuarioMail").Text &"")
		UsuarioGrupo=strResult.selectSingleNode("//UsuarioGrupo").Text &""
		UsuarioSector=strResult.selectSingleNode("//UsuarioSector").Text &""
		UsuarioNivel=strResult.selectSingleNode("//UsuarioNivel").Text &""
		UsuarioPerfil=strResult.selectSingleNode("//UsuarioPerfil").Text &""
		Error=Konvert(strResult.selectSingleNode("//Error ").Text &"")


		Session("Usuario")   = legajo
		Session("UsrOpcion") = UsuarioNombre 'session("SGU_nombre")
		Session("SkrDefT")   = "0"
		Session("SkrDefE")   = "0"
		Session("SkrDefI")   = "0"
		Session("UsrPmin")   = "0"
		Session("UsrPmax")   = "0"
		Session("UsrHist")   = "0"
		Session("SegUsuario") = UsuarioNombre 'session("SGU_nombre")
		Session("SegNivel")  = getnivelExtRiva(UsuarioNivel)
		Session("SegGrp")    = "DEF"
		Session("SegPerfil") = getperfilExtRiva(UsuarioPerfil) 'Session("Usuario")
		Session("UsuarNT")   = legajo 'Session("Usuario")

		Session("UsrSector") = getsectorExtRiva(UsuarioSector) '"" 'session("SGU_idcostcenter")
		Session("UsrSecDsc") = "" 'session("SGU_descrisector")
		Session("UsrPuesto") = "" 'session("SGU_codigopuesto")
		Session("UsrPstDsc") = "" 'session("SGU_descripuesto")
		Session("UsrMail")   = UsuarioMail 'session("SGU_mail")

		Session("SegGrpOpt")    = "S"
		Session("SegGrpLst")    = GrupoDelUsrLoguin(Legajo)


	End If


	Set oHttp=Nothing
	Set strResult=Nothing



End Function

Function ValidaMail(email)
	'Funcion que comprueba si un e-mail es valido. Devuelve:
	' 0 -> si es nulo o esta vacio
	' 1 -> Si el e-mail es incorrecto
	' 2 -> Si el e-mail es valido

	If(email)="" Then
		ValidaMail=0
		Exit Function
	End If
	Dim partes, parte, i, c
	'rompo el email en dos partes, antes y después de la arroba
	partes = Split(email, "@")
	If UBound(partes) <> 1 Then
		'si el mayor indice del array es distinto de 1 es que no he obtenido las dos partes
		ValidaMail = 1
		Exit Function
	End If
	'para cada parte, compruebo varias cosas
	For Each parte In partes
		'Compruebo que tiene algún caracter
		If Len(parte) <= 0 Then
			ValidaMail = 1
			Exit Function
		End If
		'para cada caracter de la parte
		For i = 1 To Len(parte)
			'tomo el caracter actual
			c = LCase(Mid(parte, i, 1))
			'miro a ver si ese caracter es uno de los permitidos
			If InStr("._-abcdefghijklmnopqrstuvwxyz", c) <= 0 And Not IsNumeric(c) Then
				ValidaMail = 1
				Exit Function
			End If
		Next
		'si la parte actual acaba o empieza en punto la dirección no es válida
		If Left(parte, 1) = "." Or Right(parte, 1) = "." Then
			ValidaMail = 1
			Exit Function
		End If
	Next
	'si en la segunda parte del email no tenemos un punto es que va mal
	If InStr(partes(1), ".") <= 0 Then
		ValidaMail = 1
		Exit Function
	End If
	'calculo cuantos caracteres hay después del último punto de la segunda parte del mail
	i = Len(partes(1)) - InStrRev(partes(1), ".")
	'si el número de caracteres es distinto de 2 y 3
	If Not (i = 2 Or i = 3) Then
		ValidaMail = 1
		Exit Function
	End If
	'si encuentro dos puntos seguidos tampoco va bien
	If InStr(email, "..") > 0 Then
		ValidaMail=0
		Exit Function
	End If
	ValidaMail = 2
End Function

Function getnivelExtBCR(Sdato)

	' Sdato=RTRIM(Sdato & "")
	Set cnext = ofv.conectar(ofv.strconn0)

	Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  ='" & sdato & "'"

	Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

	If Not rsdomeco.eof Then
		getnivelExtBCR=CDbl("0" & rsdomeco(0))
	Else
		getnivelExtBCR =CDbl("8") 'Sdato
	End If
	Call ofv.cerrarconsulta(rsdomeco)
	Call ofv.cerrarconn(cnext)

End Function




Function InfoDelLegajo(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("InfoDelLegajo")

	strSoapReq = strSoapReq & "<InfoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</InfoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("InfoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()

	Dim oHttpex2
	Dim strResultex2
	Set oHttpex2 = server.CreateObject("Msxml2.XMLHTTP")
	oHttpex2.open "POST", URLRiva, False
	oHttpex2.setRequestHeader "Content-Type", "text/xml"
	oHttpex2.setRequestHeader "SOAPAction", URLRiva & "InfoDelLegajo"
	oHttpex2.send strSoapReq

	Set strResultex2 = oHttpex2.responseXML

	If Not (strResultex2 Is Nothing)  Then
		session("SGU_userid") = legajo

		UsuarioNombre=Konvert(strResultex2.selectSingleNode("//UsuarioNombre").Text &"")
		UsuarioMail=Konvert(strResultex2.selectSingleNode("//UsuarioMail").Text &"")
		UsuarioGrupo=strResultex2.selectSingleNode("//UsuarioGrupo").Text &""
		UsuarioSector=strResultex2.selectSingleNode("//UsuarioSector").Text &""
		UsuarioNivel=strResultex2.selectSingleNode("//UsuarioNivel").Text &""
		UsuarioPerfil=strResultex2.selectSingleNode("//UsuarioPerfil").Text &""
		Error=strResultex2.selectSingleNode("//Error ").Text &""


		session("SGU_nombre") =UsuarioNombre
		session("SGU_mail") = UsuarioMail
		session("SGU_descrisector") =""
		session("SGU_idcostcenter") = getsectorExtRiva(UsuarioSector)
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") =getperfilExtRiva(UsuarioPerfil)
		session("SGU_idnivel")      =getnivelExtRiva(UsuarioNivel)
	Else
		session("SGU_userid") = Legajo
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel")      ="9"
	End If


	Set oHttpex2=Nothing
	Set strResultex2=Nothing



End Function
Function InfoDelLegajow(Legajo)
	' Create the http request text
	Dim strSoapReq
	strSoapReq = GenerateSoapBodyStart()
	strSoapReq = strSoapReq & GenerateSoapFunctionCallStart("InfoDelLegajo")

	strSoapReq = strSoapReq & "<InfoDelLegajoRequest>" & GenerateSoapParameter("Legajo", legajo) & "</InfoDelLegajoRequest>"
	strSoapReq = strSoapReq & GenerateSoapFunctionCallEnd("InfoDelLegajo")
	strSoapReq = strSoapReq & GenerateSoapBodyEnd()
	'response.write  strSoapReq
	Dim oHttpex
	Dim strResultex
	Set oHttpex = server.CreateObject("Msxml2.XMLHTTP")
	oHttpex.open "POST", URLRiva, False
	oHttpex.setRequestHeader "Content-Type", "text/xml"
	oHttpex.setRequestHeader "SOAPAction", URLRiva & "InfoDelLegajo"
	oHttpex.send strSoapReq

	Set strResultex = oHttpex.responseXML

	If Not (strResultex Is Nothing)  Then
		session("SGU_userid") = legajo

		UsuarioNombre=Konvert(strResultex.selectSingleNode("//UsuarioNombre").Text &"")
		UsuarioMail=Konvert(strResultex.selectSingleNode("//UsuarioMail").Text &"")
		UsuarioGrupo=strResultex.selectSingleNode("//UsuarioGrupo").Text &""
		UsuarioSector=strResultex.selectSingleNode("//UsuarioSector").Text &""
		UsuarioNivel=strResultex.selectSingleNode("//UsuarioNivel").Text &""
		UsuarioPerfil=strResultex.selectSingleNode("//UsuarioPerfil").Text &""
		Error=strResultex.selectSingleNode("//Error ").Text &""


		session("SGU_nombre") =UsuarioNombre
		session("SGU_mail") = UsuarioMail
		session("SGU_descrisector") =""
		session("SGU_idcostcenter") = getsectorExtRiva(UsuarioSector)
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = getperfilExtRiva(UsuarioPerfil)
		session("SGU_idnivel")      ="00" & UsuarioNivel
	Else
		session("SGU_userid") = Legajo
		session("SGU_nombre") = "NO"
		session("SGU_mail") = "NO"
		session("SGU_descrisector") = "NO"
		session("SGU_idcostcenter") = "0"
		session("SGU_descripuesto") = "NO"
		session("SGU_codigopuesto") = "NO"
		session("SGU_idnivel")      ="9"
	End If


	Set oHttpex=Nothing
	Set strResultex=Nothing



End Function
Function getsectorExtBCR(Sdato)

	'Sdato=mid(sdato,1,11)

	If SDATO & "" ="" Then
		getsectorExtBCR="0"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select * From eco WHERE codinter='" & sdato & "'"

		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getsectorExtBCR=rsdomeco(0)
		Else
			getsectorExtBCR="0" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function
Function getperfilExtRiva(Sdato)

	Sdato=RTrim(Sdato & "")
	If SDATO="" Then
		getperfilExtRiva="NO"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select tipoPerfil  From TipoPerfil WHERE CODpuestoE  ='" & sdato & "'"
		'RESPONSE.WRITE Strsqle3
		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getperfilExtRiva=rsdomeco(0)
		Else
			getperfilExtRiva ="NO" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function

Function getnivelExtRiva(Sdato)

	' Sdato=RTRIM(Sdato & "")
	Set cnext = ofv.conectar(ofv.strconn0)

	Strsqle3 = "Select Nccod  From NivelConfiden WHERE CODEXTERNO  ='" & sdato & "'"

	Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

	If Not rsdomeco.eof Then
		getnivelExtRiva=CDbl("0" & rsdomeco(0))
	Else
		getnivelExtRiva =CDbl("8") 'Sdato
	End If
	Call ofv.cerrarconsulta(rsdomeco)
	Call ofv.cerrarconn(cnext)

End Function

Function getsectorExtRiva(Sdato)

	'Sdato=mid(sdato,1,11)

	If SDATO & "" ="" Then
		getsectorExtRiva="0"
	Else
		Set cnext = ofv.conectar(ofv.strconn0)

		Strsqle3 = "Select * From eco WHERE codinter='" & sdato & "'"

		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			getsectorExtRiva=rsdomeco(0)
		Else
			getsectorExtRiva="0" 'Sdato
		End If
		Call ofv.cerrarconsulta(rsdomeco)
		Call ofv.cerrarconn(cnext)
	End If
End Function


Function buscarGrupoAdGP(mapgrp)

	XKEYGRP = Split(mapgrp,Chr(9))
	XDELIMG=""
	XName=""
    Set cnext = ofv.conectar(ofv.strconn0)
	For XGUINX=0 To UBound(XKEYGRP)
	    Strsqle3 = "Select mapgrp  From grupos WHERE ID  ='" & XKEYGRP(XGUINX) & "'"
		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)
        XNAMEG="NO"
		IF NOT rsdomeco.EOF  THEN XNAMEG=rsdomeco(0)
		Call ofv.cerrarconsulta(rsdomeco)
		XName= XName & XDELIMG & "  Name='" & XNAMEG & "'"
		XDELIMG= " OR "
	Next


	'Set cnext = ofv.conectar(ofv.strconn0)

	Set connaX = server.CreateObject("ADODB.Connection")
	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = session("usrEmpre")
	connaX.Properties("Password") = CStr(ofv.getCript("D",session("pwdEmpre")))
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN
	ldapEmpreldapEmpre= Replace(session("ldapEmpre"),"<","")
	ldapEmpreldapEmpre= Replace(ldapEmpreldapEmpre ,">","")



	strQuery="SELECT Name FROM '" & ldapEmpreldapEmpre & "' WHERE objectCategory='group' AND  (" & XName & ")"

	' response.write strQuery
	Set RSAX = connaX.Execute(strQuery)

	session("SGU_UIGuserid")=""
	session("SGU_UIGdescri")=""
	xgusep=""

	Do While Not   rsax.eof

		'    session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & rsax(0)
		Strsqle3 = "Select descripcion,ID  From grupos WHERE mapgrp  ='" & rsax(0) & "'"
		Set rsdomeco = ofv.crearconsultaEx(StrSqle3, cnext, 1, parametros)

		If Not rsdomeco.eof Then
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & rsdomeco(0)
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & rsdomeco(1)
		Else
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & "SD"
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & "SD"
		End If
		Call ofv.cerrarconsulta(rsdomeco)

		xgusep = Chr(9)
		RSAX.MOVENEXT
	Loop

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing
	Call ofv.cerrarconn(cnext)
End Function


Function getInfouserWADBP(legajo)

	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)

	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx =CStr(ofv.getCript("D",rsdom("pwdldap")&""))' rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,name,cn,title,mail,memberOf,DisplayName"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

		Set RSAX = connaX.Execute(strQuery)


		If RSAX.EOF Then


			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO DISPONIBLE - " & Err.NUMBER & "&nbsp;" & Err.DESCRIPTION
			session("SGU_mail") = "NO@disponible.com"
			session("SGU_descrisector") = "NO disponible"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO disponible"
			session("SGU_codigopuesto") = "NO"


		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("DisplayName")
			session("SGU_mail") = RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			Exit Do




		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)
End Function

Function getinfoUserADBP(legajo)

	Set cnDOM = ofv.conectar(ofv.strconn0)
   ' on error resume next
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"
	' RESPONSE.WRITE Strsqlq
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	' session("SGU_nombre")=""

	'if legajo="" then legajo="22"
	Do While (Not rsdom.eof)
		sldap =rsdom("raizldap")

		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,DisplayName"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
	'	response.write strQuery & "<br>"
		Set RSAX = connaX.Execute(strQuery)



		If RSAX.EOF Then
			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
			session("SGU_site")=1

		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("DisplayName")
			session("SGU_mail") =  RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			session("SGU_site")=1
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)


End Function

Function getGroupsUserADBP(legajo)

	Set connaX = server.CreateObject("ADODB.Connection")


	connaX.provider = "ADSDSOObject"
	connaX.Properties("User ID") = userN
	connaX.Properties("Password") = pass
	connaX.Properties("Encrypt Password") = True

	connaX.OPEN

	strAttributes = "sAMAccountName,cn,title,mail,memberOf,name"

	strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


	strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
	'response.write strQuery
	Set RSAX = connaX.Execute(strQuery)

	'   session("SGU_grupos") = "DEF" & CHR(9)

	session("SGU_grupos")=""

	If Not RSAX.EOF Then

		arrGroups = RSAX.Fields("memberOf")

		If Not IsNull(arrGroups) Then
			xgusep = ""
			For Each strGroup In arrGroups

				strGroup = UCase(strGroup &"")

				xgrupo=Split(strGroup,",")
				sxgrupo=Replace(xgrupo(0),"CN=","")
				session("SGU_grupos") = session("SGU_grupos") & xgusep  & sxgrupo
				xgusep = Chr(9)


			Next
		End If



	End If

	RSAX.CLOSE
	Set RSAX = Nothing
	connaX.CLOSE
	Set connaX = Nothing

End Function

Function getUsrIngroupADBP(grupo)
	Dim ADS_SCOPE_BASE
	ADS_SCOPE_BASE = 2 'subtree



	Set cnAD = ofv.conectar(ofv.strconn2)
	Strsql = "Select MapGrp From grupos where id='" & grupo & "'"
	Set rsADX = ofv.crearconsultaEx(StrSql, cnAd, 1, parametros)
	If Not rsADX.EOF Then
		GrupoADX=rsADX(0)
	'	response.write GrupoADX
		GrupoADX = FullNameGrupo(rsADX(0),strbase)
	Else
		GrupoADX = ""
	End If

	Call ofv.cerrarconsulta(rsADX)


	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL OR DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnAD, 1, parametros)
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
        xgusep=""
	Do While (Not rsdom.eof )
		sldap =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))

		strBase =  rsdom("raizldap")
		strBase2 =Replace(strBase ,"<","")
		strBase2=Replace(strBase2 ,">","")

		Set objConnection = server.CreateObject("ADODB.Connection")
		Set objCommand = server.CreateObject("ADODB.Command")
		objConnection.Provider = "ADsDSOObject"
		objConnection.Properties("User ID") = usernx
		objConnection.Properties("Password") = passx
		objConnection.Properties("Encrypt Password") = True
		objConnection.OPEN "Active Directory Provider"
		Set objCommand.ActiveConnection = objConnection

		objCommand.Properties("Page Size") = PageSizead
		objCommand.Properties("Searchscope") = ADS_SCOPE_BASE

		objCommand.CommandText = "SELECT sAMAccountName,name,cn,title,mail,DisplayName FROM '"& strBase2 &"' WHERE " & "objectCategory='user' and memberof='"& GrupoADX &"'"
		Set objRecordSet = objCommand.Execute


		Do Until objRecordSet.EOF

			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & objRecordSet("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & objRecordSet("DisplayName")
			xgusep = Chr(9)

			objRecordSet.MoveNext

		Loop


		objRecordSet.CLOSE

		objConnection.CLOSE
		Set objCommand = Nothing
		Set objConnectionn = Nothing
		rsdom.movenext
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnAD)

End Function

Function getUsrInTextADBP(texto)

	xgusep = ""
	session("SGU_UIGuserid") = ""
	session("SGU_UIGdescri") = ""
	Set cnDOM = ofv.conectar(ofv.strconn0)
	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT'"
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	TEXTO=Replace(TEXTO,"*","")
	If TEXTO="*" Or TEXTO=""   Then
		TEXTO="NO"
	ElseIf Len(texto)=1 Then
		texto= texto & "*"
	Else
		texto="*"  & texto & "*"
	End If
	Do While (Not rsdom.eof )
		strBase =rsdom("raizldap")
		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap")

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True

		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,name,DisplayName"

		'  strFilter = "(&(objectCategory=person)(objectClass=user)" & "(Name=" & texto &"))"
		strFilter="(&(objectCategory=person)(objectClass=user)(DisplayName=" & texto & ")(!userAccountControl:1.2.840.113556.1.4.803:=2))"

		strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
		' response.write strQuery

		Set RSAX = connaX.Execute(strQuery)



		Do Until RSAX.EOF
			session("SGU_UIGuserid") = session("SGU_UIGuserid") & xgusep & RSAX("sAMAccountName")
			session("SGU_UIGdescri") = session("SGU_UIGdescri") & xgusep & RSAX("DisplayName")
			xgusep = Chr(9)
			RSAX.movenext
		Loop

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.movenext
	Loop

End Function

%>
