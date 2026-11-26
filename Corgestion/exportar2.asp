
<%


scon1 = "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

Set Cn = Server.CreateObject("ADODB.Connection")
Cn.open scon1




 stablas="SYSDEFDSKTOP#SYSDEFDSKTOPDET"


 'atablas=split(stablas,"#")


' for it=0 to ubound(atablas)
for Zi=1 to request.form.count
TABLA=request.form.key(Zi)
'TABLA=atablas(it)
 strsql="select *  from " & TABLA
' RESPONSE.WRITE STRSQL
Set RS = Cn.Execute(strsql)

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
   strsql="select * from " & TABLA
   Set RS = Cn.Execute(strsql)
  ' RESPONSE.WRITE "truncate table  "  & TABLA  & CHR(13)
  ' RESPONSE.WRITE "GO" & CHR(13)
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
      IF rs(i).NAME="IDUSR" THEN
      sdata=sdata & delim & "'" & trim(replace(49 & "","'","''")) & "'"
      ELSE
      sdata=sdata & delim & "'" & trim(replace(rs(i) & "","'","''")) & "'"
      END IF
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
  'call ofv.cerrarconsulta(rs)
  RESPONSE.WRITE "set identity_insert "  & TABLA & "   OFF" & CHR(13)
  RESPONSE.WRITE "GO" & CHR(13)

  response.write chr(13)
  response.write chr(13)
  response.write chr(13)
  response.write chr(13)

next

%>