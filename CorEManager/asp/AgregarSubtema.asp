<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sTema = request.querystring("Tema")
  vuelta = request.querystring("Volver") & "&Tema=" & sTema
  bAgregar = request.querystring("Agregar")

  on error resume next

  set cn = ofv.conectar(ofv.strconn2)

  nErrores = 0

  if bAgregar = "True" then
    strsql = "Select * From Prodserv"
    set rs = ofv.crearconsulta2(strsql,cn)

    strsql = "Insert Into Prodserv (codtema"

    For Each x In Request.form
      strsql = strsql & "," & x
    Next
    strsql = strsql & ") Values (" & sTema

    For Each x In Request.form
      TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
      if TipoValido = False then exit for
      ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
      if ValorNulo = True then exit for
'''      s = ofv.queesx(rs(x).type)
'''      if rs(x).type = 200 then s = "'"

      s = "'"

      xdescri = rs(1)
     ' response.write xdescri & " xdescri " & "<br>" & chr(13)

      if xdescri <> "" then
	     strsql = strsql & "," & s & Request.Form(x) & s
         else
         strsql = strsql & "," & s & UCase(Request.Form(x)) & s
      end if
    Next

    if TipoValido = False or ValorNulo = True then
      sRedirect = "errortipo.asp?volver=" & vuelta & "&Columna=" & rs(x).name
      sRedirect = sRedirect & "&Dato=" & Request.Form(x)
      response.redirect sRedirect
    else
      strsql = strsql & ")"
     ' call ofv.crearconsultaEx(StrSql,cn,1,volver)


       STRSQL="PRC_EXECSQL '" & REPLACE(STRSQL,"'","''") & "'"
	  	   SET VV= ofv.crearconsultaEx(StrSql,cn,1,volver)

	  	     set errores = cn.errors
	  	     if cn.errors.count = 0 then
	  	       IF NOT ISNULL(VV(0)) THEN  IDINSERT=VV(0)
	           end if

	  		sql = "select id from clienteplan"
	  		set rx = ofv.crearconsultaEx(sql,cn,1,volver)

	  		if not rx.eof then
	  'response.write "UN: " & session("UN")
	  		do while not rx.eof
	  		puede = 0

	  		if cstr(rx(0)&"") = cstr(session("UN")&"") or cstr(rx(0)&"") = 1 then puede = 1

	  		sql = "insert into ATRMTEMAST (atrmtipost,atrmcodst,atrmcodes,atrmverst) values ('"
	  		sql = sql & "S','" & IDINSERT & "','" & rx(0) & "','" & puede & "')"
	  		'response.write sql & "<br>" & "puede: " & puede & "<br>" & "UN: " & session("UN")
	  		call ofv.crearconsultaEx(sql,cn,1,parametros)

	  		rx.movenext
	  		loop

	  		end if

		call ofv.cerrarconsulta(rx)


      nErrores = cn.errors.count

      if nErrores = 0 then
        response.redirect vuelta
      else
        response.write ofv.ErrorDB(cn,vuelta)
      end if
    end if
    call ofv.cerrarconsulta(rs)
  else

  end if

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

if nErrores = 0 then response.write sHTML

%>