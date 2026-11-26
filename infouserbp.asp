<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->


<form action="infouser.asp" method=post>
usuario id : <input type=text name=usr>
<input type=submit value=consultar>
</form>
<%
usr=request.form("usr")

if usr <> "" then
xx=getinfoUserADBP(usr)
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


Function getinfoUserADBP(legajo)
    on error resume next
	Set cnDOM = ofv.conectar(ofv.strconn0)

	Strsqlq = "Select * From dominios WHERE RAIZLDAP IS NOT NULL and DESCRIPCION <> 'DEFAULT' order by id"
	' RESPONSE.WRITE Strsqlq
	Set rsdom = ofv.crearconsultaEx(StrSqlq, cnDOM, 1, parametros)
	' session("SGU_nombre")=""
	Do While (Not rsdom.eof)  Or (session("SGU_nombre") <> "NO")
		sldap =rsdom("raizldap")

		usernx =  rsdom("usrldap")
		passx = CStr(ofv.getCript("D",rsdom("pwdldap")&""))'rsdom("pwdldap") 'ofv.GetCript("D", CSTR(rsdom("pwdldap")))

		Set connaX = server.CreateObject("ADODB.Connection")


		connaX.provider = "ADSDSOObject"
		connaX.Properties("User ID") = usernx
		connaX.Properties("Password") = passx
		connaX.Properties("Encrypt Password") = True
		' response.write userN & "----" & pass & "---" & strBase
		connaX.OPEN

		strAttributes = "sAMAccountName,cn,title,mail,memberOf,name,DisplayName"

		strFilter = "(&(objectCategory=person)(objectClass=user)" & "(samAccountName=" & legajo &"))"


		strQuery = sldap & ";" & strFilter & ";" & strAttributes & ";subtree"
		'response.write strQuery
		Set RSAX = connaX.Execute(strQuery)

       If connaX.Errors.Count > 0 Then
	         ' Enumerate Errors collection and display
	         ' properties of each Error object.
	         For Each errLoop In connaX.Errors
	            strError = "Error #" & errLoop.Number & "<br>" & _
	               "   " & errLoop.Description & "<br>" & _
	               "   (Source: " & errLoop.Source & ")" & "<br>" & _
	               "   (SQL State: " & errLoop.SQLState & ")" & "<br>" & _
	               "   (NativeError: " & errLoop.NativeError & ")" & "<br>"
	            If errLoop.HelpFile = "" Then
	               strError = strError & _
	                  "   No Help file available" & _
	                  "<br><br>"
	            Else
	               strError = strError & _
	                  "   (HelpFile: " & errLoop.HelpFile & ")" & "<br>" & _
	                  "   (HelpContext: " & errLoop.HelpContext & ")" & _
	                  "<br><br>"
	            End If

	            Response.Write("<p>" & strError & "</p>")
	         Next
          End If

		If RSAX.EOF Then
			session("SGU_userid") = Legajo
			session("SGU_nombre") = "NO"
			session("SGU_mail") = "NO"
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = "NO"
			session("SGU_codigopuesto") = "NO"
			session("SGU_site")=1

		Else

			session("SGU_userid") = Legajo
			session("SGU_nombre") = RSAX("DisplayName")
			session("SGU_mail") =  RSAX("mail")
			session("SGU_descrisector") = "NO"
			session("SGU_idcostcenter") = "0"
			session("SGU_descripuesto") = RSAX("title")
			session("SGU_codigopuesto") = "NO"
			session("SGU_site")=1
			Exit Do
		End If

		RSAX.CLOSE
		Set RSAX = Nothing
		connaX.CLOSE
		Set connaX = Nothing
		rsdom.MOVENEXT
	Loop
	Call ofv.cerrarconsulta(rsdom)
	Call ofv.cerrarconn(cnDOM)


End Function

%>