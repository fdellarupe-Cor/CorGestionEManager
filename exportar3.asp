<!-- #INCLUDE FILE="coremanager/IncFile/CorConect.inc" -->
<%


Set cn = ofv.conectar(ofv.strconn2)




stablas="SYSDEFDSKTOP#SYSDEFDSKTOPDET"


atablas=Split(stablas,"#")


For it=0 To UBound(atablas)
	'for Zi=1 to request.form.count
	'TABLA=request.form.key(Zi)
	TABLA=atablas(it)
	strsql="select * from " & TABLA
	' RESPONSE.WRITE STRSQL
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	
	scol=""
	
	delim=""
	For i=0 To rs.fields.count-1
		If i=0 Then col= rs(i).name
		scol=scol & delim & rs(i).name
		delim=","
	Next
	rs.close
	If tabla="SYSDEFDSKTOP" Then
		strsql="select * from " & TABLA & " where objid=38 order by " & col
	Else
		strsql="select * from " & TABLA & " where codset=38 order by " & col
	End If
	
	' response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	'  RESPONSE.WRITE "truncate table  "  & TABLA  & CHR(13)
	' RESPONSE.WRITE "GO" & CHR(13)
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
					sdata=sdata & delim & "'" & Trim(Replace(rs(i) & "","'","''")) & "'"
				End If
			End If
			delim=","
		Next
		sql=sql &sdata & ")" & Chr(13) & "go" & Chr(13)
		rs.movenext
		response.write sql
		response.flush
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