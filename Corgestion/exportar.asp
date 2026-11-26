
<%

For i=1 To request.form.count
	RESPONSE.WRITE  request.form.key(i)
Next
scon1 = "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

Set Cn = Server.CreateObject("ADODB.Connection")
Cn.open scon1


Set RS = Cn.Execute("sp_cg_Utilidades @Modo='listar_tablas'")
response.write "<form action=exportar.asp?todas=S method=post ><input value='seleccionar todas las tablas' type=submit></form>"
response.write "<form action=exportar.asp?todas=N method=post ><input value='No seleccionar todas las tablas' type=submit></form>"
response.write "<form action=exportar2.asp method=post TARGET=_NEW><input value=exportar type=submit><br>"
C=1
'RESPONSE.WRITE "<TABLE BORDER><TR><TD>TODOS<input type=checkbox  name=TODOS></TD></TR><TR>"
RESPONSE.WRITE "<TABLE BORDER><TR>"

todas=request.querystring("todas")

If todas="S" Then
	check="checked"
Else
	check=""
End If

Do While Not rs.eof

	If C <7 Then
		response.write "<TD>" & rs(0) & "<input type=checkbox " & check &" name="&rs(0) & "></TD>"
		C=C+1
	Else
		response.write "<TD>" & rs(0) & "<input type=checkbox   " & check &" name="&rs(0) & "></TD>"
		RESPONSE.WRITE "</TR><TR>"
		C=1
	End If
	rs.movenext
	'exit do
Loop
response.write "</TABLE></form>"



stablas="DOMINIOS"


atablas=Split(stablas,"#")


'for it=0 to -1 'ubound(atablas)
For Zi=1 To request.form.count
	TABLA=request.form.key(Zi)
	strsql="select * from " & TABLA '& " order by id"
	'RESPONSE.WRITE STRSQL
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	scol=""

	delim=""
	For i=0 To rs.fields.count-1
		scol=scol & delim & rs(i).name
		delim=","
	Next
	'  RESPONSE.WRITE "truncate table  "  & atablas(it)  & CHR(13)
	'  RESPONSE.WRITE "GO" & CHR(13)
	RESPONSE.WRITE "set identity_insert "  & TABLA & "  on" & Chr(13)
	RESPONSE.WRITE "GO" & Chr(13)
	Do While (Not rs.EOF)
		sdata=""
		delim=""
		sql=Chr(13) & "insert into " & TABLA  & "("
		sql=sql & scol & ") values ("
		For i=0 To rs.fields.count-1

			' response.write rs(i).type  &"  " & rs(i).name  & chr(13)
			If  IsNull(rs(i)) Then
				sdata=sdata & delim & "NULL"
			Else

				If rs(i).type=11 Then ' tipo bit
					vNuevoValor=rs(i)
					If UCase(vNuevoValor) = "FALSE" Then vNuevoValor = "0"
					If UCase(vNuevoValor) = "TRUE" Then vNuevoValor = "1"

					If UCase(vNuevoValor) = "FALSO" Then vNuevoValor = "0"
					If UCase(vNuevoValor) = "VERDADERO" Then vNuevoValor = "1"
					sdata=sdata & delim & "'" & vNuevoValor & "'"
				ElseIf rs(i).type=131 Then ' tipo decimal
					sdata=sdata & delim & "'" & Replace(rs(i)&"",",",".") & "'"
				Else
					sdata=sdata & delim & "'" & Replace(rs(i) & "","'","''") & "'"
				End If
			End If
			delim=","
		Next
		sql=sql &sdata & ")" & Chr(13) & "go" & Chr(13)
		rs.movenext
		response.write sql
		' exit do
	Loop
	Call ofv.cerrarconsulta(rs)
	RESPONSE.WRITE "set identity_insert "  & TABLA & "   OFF" & Chr(13)
	RESPONSE.WRITE "GO" & Chr(13)

	response.write Chr(13)
	response.write Chr(13)
	response.write Chr(13)
	response.write Chr(13)

Next

%>