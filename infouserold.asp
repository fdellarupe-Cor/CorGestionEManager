<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->


<form action="infouser.asp" method=post>
usuario id : <input type=text name=usr>
<input type=submit value=consultar>
</form>
<%
usr=request.form("usr")

if usr <> "" then
xx=getinfouser(usr)
session("SGU_grupos")=""
session("SegGrpLst")=""

call getGroupsUser(USR)


'CALL SendMail("prueba", session("SGU_mail"), session("SGU_mail"),"prueba", session("MailComponent"),"SI",0, "NO")

end if
set cn=ofv.conectar(ofv.strconn2)

response.write cn.ConnectionTimeout

 Session("SegGrpLst")    = session("SGU_grupos")


CALL  ChkGrp()
'RESPONSE.WRITE session("SegGrpLst")

DOMC= REQUEST.servervariables("LOGON_USER")
DOMC=SPLIT(DOMC,"\")

if ubound(DOMC) >=0 THEN
SESSION("DOMINIOX")=DOMC(0)
ELSE
SESSION("DOMINIOX")=""
END IF


'response.write session("dominioX") & "222"
 XXX=ValDominio(session("dominioX"))

RESPONSE. WRITE "RAIZ: " & REPLACE(REPLACE(session("ldap"),"<","*"),">","*")  & "<BR>"


%>
<% IF USR <> "" THEN%>
<TABLE BORDER>
<TR>
<TD>USR ID </TD><TD><%=session("SGU_userid")%></TD>
</tr>
<tr>
<TD>NOMBRE </TD><TD><%=session("SGU_nombre")%> </TD>
</tr>
<tr>
<TD>MAIL </TD><TD><%=session("SGU_mail")%></TD>
</tr>
<tr>
<TD>DESCRI SECTOR </TD><TD><%=session("SGU_descrisector")%> </TD>
</tr>
<tr>
<TD>SECTOR </TD><TD><%=session("SGU_idcostcenter")%></TD>
</tr>
<tr>
<TD>DESCRI PUESTO </TD><TD><%=session("SGU_descripuesto")%> </TD>
</tr>
<tr>
<TD>COD PUESTO </TD><TD><%=session("SGU_codigopuesto")%> </TD>

</TR>

<tr>
<TD>GRUPOS </TD><TD> <%=session("SegGrpLst")%> </TD>

</TR>
<TABLE>

<% END IF %>

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
'  response.write XFStrSql
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