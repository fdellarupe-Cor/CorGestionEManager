


<form action="pruebagrupos.asp" method=post>
<br>
raiz ldap (ejemplo < LDAP://COR01 >): <input type=text value='<LDAP://COR01>' name=ldap>
<br>
<br>
usuario conexion ldap : <input type=text name=usernx>
<br>
<br>
pwd  usuario conexion ldap : <input type=password name=pwdx>
<br>

<br>
<input type=submit value=consultar>
</form>
<%
pwdx=request.form("pwdx")
usernx=request.form("usernx")
ldap=request.form("ldap")
'response.write  usernx & "     "  & pwdx

if pwdx <> "" then
xxx=buscarG2(usernx,pwdx,ldap)
end if






function buscarG(usernx,passx,sldap)
SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE
     '   response.write userN & "----" & pass & "---" & strBase
        connaX.OPEN

         strAttributes = "name,description,mail,groupType"

        ' strFilter = "(&((&(objectCategory=group)(&(groupType:1.2.840.113556.1.4.803:=8)(!(groupType:1.2.840.113556.1.4.803:=-2147483648))))) (name=*) )"
           strFilter =  	"(&(objectCategory=group)(name=*usu*))"

         strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)
'response.write rsax.eof & " 2222222222222"
 response.write "<table border>"
 response.write  "<tr><td><B>GRUPO</B></td><TD><B>MAIL</B></TD><TD><B>groupType</B></TD></TR>"
   do while not  rsax.eof


     response.write "<tr>"

      response.write "<td>" & RSAX("name") & "</td>"
      'response.write "<td>" &  RSAX("description")  & "</td>"

     response.write "<td>" & rsax("mail")  & "</td>"
     response.write "<td>" & rsax(3)  & "</td>"
      response.write "</tr>"

rsax.movenext
   loop
 response.write "<table>"
   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

end function


function buscarG2(usernx,passx,sldap)
SET connaX = server.CREATEOBJECT("ADODB.Connection")


       connaX.provider = "ADSDSOObject"
       connaX.Properties("User ID") = usernx
       connaX.Properties("Password") = passx
       connaX.Properties("Encrypt Password") = TRUE

        connaX.OPEN
        strQuery="SELECT Name FROM '" & sldap & "' WHERE objectCategory='group' AND   (Name='master*' or name='desarrollador*') "

         'strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"

'response.write strQuery
   SET RSAX = connaX.EXECUTE(strQuery)
'response.write rsax.eof & " 2222222222222"
 response.write "<table border>"
 response.write  "<tr><td><B>GRUPO</B></td><TD><B>MAIL</B></TD><TD><B>groupType</B></TD></TR>"
   do while not  rsax.eof


     response.write "<tr>"

      response.write "<td>" & RSAX("name") & "</td>"
      'response.write "<td>" &  RSAX("description")  & "</td>"

    ' response.write "<td>" & rsax("mail")  & "</td>"
     'response.write "<td>" & rsax(3)  & "</td>"
      response.write "</tr>"

rsax.movenext
   loop
 response.write "<table>"
   RSAX.CLOSE
   SET RSAX = NOTHING
   connaX.CLOSE
   SET connaX = NOTHING

end function



%>