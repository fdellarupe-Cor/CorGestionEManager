<!-- #INCLUDE FILE="coremanager/IncFile/CorConect.inc" -->
<%


Set cn = ofv.conectar("Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=A_parametros_gp;user id=sa;password=bb4515")

Set cn2 = ofv.conectar(ofv.strconn0)


dim idInicio,idcliente,ecosite



idInicio=1000 ' id comienzo eco

idcliente=5   ' id cliente plan santa fe

ecosite = 10  ' tabla ecosite /site



stablas="eco"


atablas=Split(stablas,"#")

RESPONSE.WRITE "delete from eco where eco >  "  & idInicio & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)
RESPONSE.WRITE "delete from ecosite where eco >  " & idInicio & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)

RESPONSE.WRITE "delete from ecorel where ecodepend >  " & idInicio & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)

RESPONSE.WRITE "delete from ecorelsite where codrel >  " & idInicio & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)

RESPONSE.WRITE "delete from ecoperfil where codsector >  " & idInicio & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)


RESPONSE.WRITE "truncate table  tipoperfil" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)


RESPONSE.WRITE "truncate table Niveljerar" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)


RESPONSE.WRITE "truncate table  Segmento" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)


RESPONSE.WRITE "truncate table jerarquia " & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)



For it=0 To UBound(atablas)
	'for Zi=1 to request.form.count
	'TABLA=request.form.key(Zi)
	TABLA=atablas(it)
	strsql="select distinct " & idInicio & " + id as eco, descripción as descri,idsegmento as ecoseg,idjerarquia as ecojerar, 0 as codinter ,'B' as nodo , 3 as estado , 1 as ecosite , " & idcliente  & " as ecoun ,idniveles as econiv from AySSantaCruzfinal "

  strsql=strsql & "  where id  is not null"
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


	' response.write strsql

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

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

   sql="select distinct " & idInicio & " + id as eco from AySSantaCruzfinal where id  is not null"


   Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)


   do while not rs.eof
   response.write "insert into ecosite (eco,ecosite) values(" & rs(0) & "," & ecosite & ")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop

RESPONSE.WRITE "set identity_insert  ecorel  on" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)

 sql=" select distinct " & idInicio & " + iddependencia as ecomaster, " & idInicio & " + id as ecodepend from AySSantaCruzfinal  where iddependencia  is not null"
'response.write sql

   Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)


   do while not rs.eof
   response.write "insert into ecorel (codrel,ecomaster,ecodepend) values(" & rs(1) & "," & rs(0) & "," & rs(1) &")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop
RESPONSE.WRITE "set identity_insert  ecorel  OFF" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)


 sql=" select distinct " & idInicio & " + iddependencia as ecomaster, " & idInicio & " + id as ecodepend from AySSantaCruzfinal  where iddependencia  is not null"


   Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)


   do while not rs.eof
   response.write "insert into ecorelsite (codrel,site) values(" & rs(1) & "," &  ecosite & ")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop

    sql=" select idperfiles ,id + " & idInicio & " ,1 from AySSantaCruzfinal  where iddependencia  is not null"


      Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)


      do while not rs.eof
      response.write "insert into ecoperfil (codsector,codperfil,dotacion) values('" & rs(1)  &"','" & rs(0) & "',1)" & Chr(13)
      RESPONSE.WRITE "GO" & Chr(13)


      rs.movenext
   loop



'RESPONSE.WRITE "set identity_insert  tipoperfil  on" & Chr(13)
'RESPONSE.WRITE "GO" & Chr(13)

 sql="select * from tipoperfil"


   Set rs = ofv.crearconsultaEx(sql,cn2,1,parametros)


   do while not rs.eof
   response.write "insert into tipoperfil (tipoperfil,descri,tipoun,estado) values('" & rs(0) & "','" & rs(1) & "','" & rs(2)  &"'," & rs("estado") &")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop
'RESPONSE.WRITE "set identity_insert  tipoperfil  OFF" & Chr(13)
'RESPONSE.WRITE "GO" & Chr(13)



RESPONSE.WRITE "set identity_insert  Niveljerar  on" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)

 sql=" select * from Niveljerar"


   Set rs = ofv.crearconsultaEx(sql,cn2,1,parametros)


   do while not rs.eof
   response.write "insert into Niveljerar (codnivel,descri,grafica,estado) values('" & rs(0) & "','" & rs(1) & "','" & rs(2)  & "'," & rs("estado") & ")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop
RESPONSE.WRITE "set identity_insert  Niveljerar  OFF" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)



RESPONSE.WRITE "set identity_insert  Segmento  on" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)

 sql=" select * from Segmento"


   Set rs = ofv.crearconsultaEx(sql,cn2,1,parametros)


   do while not rs.eof
   response.write "insert into Segmento (codseg,descri,estado) values('" & rs(0) & "','" & rs(1)  & "'," & rs("estado") & ")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop
RESPONSE.WRITE "set identity_insert  Segmento  OFF" & Chr(13)
RESPONSE.WRITE "GO" & Chr(13)


'RESPONSE.WRITE "set identity_insert  Jerarquia  on" & Chr(13)
'RESPONSE.WRITE "GO" & Chr(13)

 sql=" select * from Jerarquia"


   Set rs = ofv.crearconsultaEx(sql,cn2,1,parametros)


   do while not rs.eof
   response.write "insert into Jerarquia (codjerar,descri,escala,orden,estado) values('" & rs(0) & "','" & rs(1)   & "','" & rs(2)  & "','" & rs(3)  & "'," & rs("estado") & ")" & Chr(13)
   RESPONSE.WRITE "GO" & Chr(13)


   rs.movenext
   loop
'RESPONSE.WRITE "set identity_insert  Jerarquia  OFF" & Chr(13)
'RESPONSE.WRITE "GO" & Chr(13)
	response.write Chr(13)
	response.write Chr(13)
	response.write Chr(13)
	response.write Chr(13)

Next

%>