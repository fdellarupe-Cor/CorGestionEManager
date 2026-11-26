<!-- #INCLUDE FILE="coremanager/IncFile/CorConect.inc" -->
<%

for i=1 to request.form.count
RESPONSE.WRITE  request.form.key(i)
NEXT

 set cn = ofv.conectar(ofv.strconn5)


sql="SELECT CAST(table_name as varchar) as tabla FROM INFORMATION_SCHEMA.TABLES where table_type='BASE TABLE' order by tabla"

set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
response.write "<form action=exportar.asp?todas=S method=post ><input value='seleccionar todas las tablas' type=submit></form>"
response.write "<form action=exportar.asp?todas=N method=post ><input value='No seleccionar todas las tablas' type=submit></form>"
response.write "<form action=exportar2wkf.asp method=post TARGET=_NEW><input value=exportar type=submit><br>"
C=1
'RESPONSE.WRITE "<TABLE BORDER><TR><TD>TODOS<input type=checkbox  name=TODOS></TD></TR><TR>"
RESPONSE.WRITE "<TABLE BORDER><TR>"

todas=request.querystring("todas")

if todas="S" then
  check="checked"
else
  check=""
end if

do while not rs.eof

IF C <7 THEN
response.write "<TD>" & rs(0) & "<input type=checkbox " & check &" name="&rs(0) & "></TD>"
C=C+1
ELSE
response.write "<TD>" & rs(0) & "<input type=checkbox   " & check &" name="&rs(0) & "></TD>"
RESPONSE.WRITE "</TR><TR>"
C=1
END IF
rs.movenext
'exit do
loop
response.write "</TABLE></form>"



 stablas="DOMINIOS"


 atablas=split(stablas,"#")


 'for it=0 to -1 'ubound(atablas)
for Zi=1 to request.form.count
TABLA=request.form.key(Zi)
 strsql="select * from " & TABLA '& " order by id"
 'RESPONSE.WRITE STRSQL
 set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   scol=""

   delim=""
   for i=0 to rs.fields.count-1
   scol=scol & delim & rs(i).name
   delim=","
   next
 '  RESPONSE.WRITE "truncate table  "  & atablas(it)  & CHR(13)
 '  RESPONSE.WRITE "GO" & CHR(13)
   RESPONSE.WRITE "set identity_insert "  & TABLA & "  on" & CHR(13)
   RESPONSE.WRITE "GO" & CHR(13)
   Do While (Not rs.EOF)
   sdata=""
   delim=""
   sql=chr(13) & "insert into " & TABLA  & "("
   sql=sql & scol & ") values ("
   for i=0 to rs.fields.count-1

   ' response.write rs(i).type  &"  " & rs(i).name  & chr(13)
   IF  ISNULL(rs(i)) THEN
    sdata=sdata & delim & "NULL"
   ELSE

   if rs(i).type=11 then ' tipo bit
      vNuevoValor=rs(i)
     if UCase(vNuevoValor) = "FALSE" then vNuevoValor = "0"
     if UCase(vNuevoValor) = "TRUE" then vNuevoValor = "1"

     if UCase(vNuevoValor) = "FALSO" then vNuevoValor = "0"
     if UCase(vNuevoValor) = "VERDADERO" then vNuevoValor = "1"
     sdata=sdata & delim & "'" & vNuevoValor & "'"
   elseif rs(i).type=131 then ' tipo decimal
      sdata=sdata & delim & "'" & replace(rs(i)&"",",",".") & "'"
   else
      sdata=sdata & delim & "'" & replace(rs(i) & "","'","''") & "'"
   end if
   END IF
   delim=","
   next
   sql=sql &sdata & ")" & chr(13) & "go" & chr(13)
   rs.movenext
   response.write sql
  ' exit do
   loop
  call ofv.cerrarconsulta(rs)
  RESPONSE.WRITE "set identity_insert "  & TABLA & "   OFF" & CHR(13)
  RESPONSE.WRITE "GO" & CHR(13)

  response.write chr(13)
  response.write chr(13)
  response.write chr(13)
  response.write chr(13)

next

%>