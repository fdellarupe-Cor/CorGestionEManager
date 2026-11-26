<!-- #INCLUDE FILE="coremanager/IncFile/CorConect.inc" -->
<%


 set cn = ofv.conectar(ofv.strconn0)




 stablas="SYSDEFDSKTOP#SYSDEFDSKTOPDET"


 'atablas=split(stablas,"#")


' for it=0 to ubound(atablas)
for Zi=1 to request.form.count
TABLA=request.form.key(Zi)
'TABLA=atablas(it)
 strsql="select * from " & TABLA
' RESPONSE.WRITE STRSQL
 set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   scol=""

   delim=""
   for i=0 to rs.fields.count-1
   if i=0 then col= rs(i).name
   if rs(i).name <> "fulltext_timestamp" then
   scol=scol & delim & rs(i).name
   delim=","
   end if
   next
   rs.close
   strsql="select * from " & TABLA & " where codemp=1 order by " & col
  ' response.write strsql
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   RESPONSE.WRITE "truncate table  "  & TABLA  & CHR(13)
   RESPONSE.WRITE "GO" & CHR(13)
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
    if rs(i).name <> "fulltext_timestamp" then sdata=sdata & delim & "'" & trim(replace(rs(i) & "","'","''")) & "'"
   end if
   END IF
   delim=","
   next
   sql=sql &sdata & ")" & chr(13) & "go" & chr(13)
   rs.movenext
   response.write sql
   response.flush
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