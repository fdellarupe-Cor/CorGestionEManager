<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->


<form action="infouser.asp" method=post>
usuario  a buscar (id): <input type=text name=usr>
<input type=submit value=consultar>
</form>
<%
usr=request.form("usr")
'session("USRPRV")  = "BPN"
session("SGU_grupos")=""
session("SegGrpLst")=""
IF usr <> "" THEN
xx=getinfouser(usr)

call getGroupsUser(USR)
END IF
set cn=ofv.conectar(ofv.strconn2)
SQL="SELECT * FROM GRUPOS order by descripcion"

set XFrs = ofv.crearconsultaEx(SQL,CN,1,parametros)
STG="<U>LISTA DE GRUPOS EN EMANAGER</U><BR><BR><table border ><tr><td align=center bgcolor=#f8f9f9 ><I>ID EMANAGER</i></td><td bgcolor=#f8f9f9 align=center><i> MAPEO AD </i></td><td bgcolor=#f8f9f9 align=center><i>DESCRIPCION</i></td></tr>"
DO WHILE NOT XFrs.EOF
 STG="<tr>" & STG & "<td>" & XFrs("ID") & "</td><td>"  & XFrs("MAPGRP") & " </td><td>" & XFrs("descripcion") & "</td></tr>"
 XFrs.MOVENEXT
LOOP
IF usr <> "" THEN
 response.write STG & "</table><br><br> "
END IF
 Session("SegGrpLst")    = session("SGU_grupos")


IF usr <> "" THEN CALL  ChkGrp()
'RESPONSE.WRITE session("SGU_grupos")

DOMC= REQUEST.servervariables("LOGON_USER")
DOMC=SPLIT(DOMC,"\")

if ubound(DOMC) >=0 THEN
SESSION("DOMINIOX")=DOMC(0)
ELSE
SESSION("DOMINIOX")=""
END IF


'response.write session("dominioX") & "222"
   XXX=ValDominio(session("dominioX"))

'RESPONSE. WRITE "RAIZ: " & REPLACE(REPLACE(session("ldap"),"<","*"),">","*")  & "<BR>"
'RESPONSE. WRITE "USRLDAP: " &session("usrldap") & "<BR>"
'RESPONSE. WRITE "PWDLDAP: " &  session("pwdldap") & "<BR>"

%>
<% IF USR <> "" THEN%>
<U>INFORMACION DEL USUARIO</U><BR><BR> <TABLE BORDER>
<TR>
<TD bgcolor=#f8f9f9><i>USR ID </i></TD><TD><%=session("SGU_userid")%></TD>
</tr>
<tr>
<TD bgcolor=#f8f9f9><i>NOMBRE</i></TD><TD><%=session("SGU_nombre")%> </TD>
</tr>
<tr>
<TD bgcolor=#f8f9f9><i>MAIL </i></TD><TD><%=session("SGU_mail")%></TD>
</tr>
<tr>
<TD bgcolor=#f8f9f9><i>DESCRI SECTOR </i></TD><TD><%=session("SGU_descrisector")%> </TD>
</tr>
<tr>
<TD bgcolor=#f8f9f9><i>SECTOR </i></TD><TD><%=session("SGU_idcostcenter")%></TD>
</tr>
<tr>
<TD bgcolor=#f8f9f9><i>DESCRI PUESTO </i></TD><TD><%=session("SGU_descripuesto")%> </TD>
</tr>
<tr>
<TD bgcolor=#f8f9f9><i>COD PUESTO </i></TD><TD><%=session("SGU_codigopuesto")%> </TD>

</TR>

<tr>
<TD bgcolor=#f8f9f9><i>GRUPOS </i></TD><TD> <%=session("SegGrpLst")%> </TD>

</TR>
<TABLE>

<% END IF

IF usr <> "" THEN
 RESPONSE.WRITE "<BR><BR><U>GRUPOS AD DEL USUARIO</U><BR><BR>"
 RESPONSE.WRITE replace(session("SGU_grupos"),chr(9),"<br>")
END IF
%>

<%

function ChkGrp()

'    SESSION("USREXPIR") = 50

  SESSION("USREXPIR") = "0"

  set XFCN = ofv.conectar(ofv.strconn0)

  XFStrSql = "Select id,sessexp,descripcion From Grupos where "
   Session("SegGrp")    = "DEF"
  session("SegGrpLst") = Session("SegGrp") & chr(9) & session("SegGrpLst")
  session("SegGrpdsc") = Session("SegGrp") & chr(9) & session("SegGrpLst")


  XMFTGRP = split(session("SegGrpLst"),chr(9))
  XMFSEP  = ""

  if session("USRPRV")  = "I" then
  XFStrSql = XFStrSql & " id in ("
  else
  XFStrSql = XFStrSql & " mapgrp in ("
  end if

  for XMFI = 0 to ubound(XMFTGRP)
       IF (XMFTGRP(XMFI) & "" <> "") THEN
        XFStrSql  = XFStrSql  & XMFSEP & "'" & replace(XMFTGRP(XMFI) & "","'","''") & "'"
        XMFSEP  = ","
      END IF
  next

  XFStrSql = XFStrSql & ") "
  XFStrSql = XFStrSql & " Order By id"
 ' response.write XFStrSql
  set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

  if  not XFrs.eof then
      XMFSEP               = ""
      session("SegGrpLst") = ""
      session("SegGrpdsc") = ""
      do until XFRS.EOF

         if XFrs(1) <> "D" and SESSION("USREXPIR") <> "D" then
            if (cint("0" & XFrs(1)) * 60) > cint(SESSION("USREXPIR")) then
               SESSION("USREXPIR") = cint("0" & XFrs(1)) * 60
            end if
          else
           SESSION("USREXPIR") = "D"
          end if

         session("SegGrpLst") = session("SegGrpLst") & XMFSEP & XFRS(0)
         session("SegGrpdsc") = session("SegGrpdsc") & XMFSEP & XFRS(2)
         XMFSEP  = chr(9)
      XFRS.MOVENEXT
      LOOP

      IF  session("SegGrpLst") = "" THEN session("SegGrpLst") = Session("SegGrp")

  else
      session("SegGrpLst") = Session("SegGrp")
      session("SegGrpdsc") = Session("SegGrp")
  end if
  call ofv.cerrarconsulta(XFrs)
  call ofv.cerrarconn(XFCN)



end function

%>