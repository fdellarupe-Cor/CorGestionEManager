<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	SESSION("BASE")="D"



	Set cn = ofv.conectar(ofv.strconn2)
	sHTML = ofv.FORMheader("")
	'WWWAPUNTA = " <script language=javascript >  clactfrm('11','0','0','0','0','0','0','0','0','0'); document.XGUFRM.submit();</script>"
'response.write request.form
	wtipoL   = 0
	wtipOBJ  = 0

	wNV1    = "0"
	wNV2    = "0"
	wNV3    = "0"


	sHTML = sHTML &  GenBuscaObj(wtipoL,wtipOBJ,WNV1,WNV2,WNV3,"0","0") '&WWWAPUNTA
	sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,celda)

	sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
	sAccion = "AdNovedadesProp.asp?" & request.querystring & "&cobj=" & XGUFCOBJ & "&tobj=" & XGUFTOBJ & "&solapa=" & XGUFNIV1 & "&act=SI"
	FormuA = ofv.GenForm("BotonA","",sAccion,"","post","",sInput,"si")
	Celda =  ofv.GenCelda("","","","","","","",FormuA,"si")

	sInput = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
	sAccion = "AdNovedadesProp.asp?" & request.querystring & "&act=NO"
	FormuC = ofv.GenForm("BotonC","",sAccion,"","post","",sInput,"si")
	Celda = Celda & ofv.GenCelda("","","","","","","",FormuC,"si")

	Fila = ofv.genrow("","","","","","","",Celda,"si")

	tablax = ofv.GenTabla("","","center","15% ","","0","","0","1","","",Fila,"si")

	sHTML = sHTML & tablax & "</body></html>"
	response.write sHTML
	Call ofv.cerrarconn(cn)

End If
%>