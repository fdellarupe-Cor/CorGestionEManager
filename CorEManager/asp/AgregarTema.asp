<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  vuelta = request.querystring("Volver")
  bAgregar = request.querystring("Agregar")

  on error resume next

  set cn = ofv.conectar(ofv.strconn2)

  nErrores = 0

  if bAgregar = "True" then
    strsql = "Select * From Tema"

 '' response.write strsql & " strsql -0- " & "<br>"

    set rs = ofv.crearconsulta2(strsql,cn)

    strsql = "Insert Into Tema ("
    n = 0
    For Each x In Request.form
      if n = 0 then
        strsql = strsql & x
      else
        strsql = strsql & "," & x
      end if

      n = n + 1

    Next
    strsql = strsql & ") Values ("

 '' response.write strsql & " strsql -1- " & "<br>"

    n = 0
    For Each x In Request.form
      TipoValido = ofv.ValidarTipo(rs(x),Request.Form(x))
      if TipoValido = False then exit for
      ValorNulo = ofv.EsNulo(rs(x),Request.Form(x))
      if ValorNulo = True then exit for

''''      s = ofv.queesx(rs(x).type)
''''      if rs(x).type = 200 then s = "'"

       s = "'"


  xdescri = rs(1)
  'response.write xdescri & " xdescri " & "<br>" & chr(13)

      if n = 0 then
         if xdescri <> "" then
           strsql = strsql & s & Request.Form(x) & s
         else
           strsql = strsql & s & UCase(Request.Form(x)) & s
         end if
      else
         if xdescri <> "" then
           strsql = strsql & "," & s & Request.Form(x) & s
         else
           strsql = strsql & "," & s & UCase(Request.Form(x)) & s
         end if
     end if

     n = n + 1

    Next

    if not TipoValido or ValorNulo then
      response.write ofv.ErrorTipo(vuelta,Request.Form(x),rs(nx).name)
      if err.number <> 0 then response.write ofv.ErrorTipo(vuelta,"","")
    else
      strsql = strsql & ")"
 ''   response.write strsql & " strsql -2- " & "<br>"
      'call ofv.crearconsultaEx(StrSql,cn,1,parametros)



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
		sql = sql & "T','" & IDINSERT & "','" & rx(0) & "','" & puede & "')"
		'response.write strsql & "<br>" & "puede: " & puede & "<br>" & "UN: " & session("UN")
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

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

if nErrores = 0 then response.write sHTML

%>