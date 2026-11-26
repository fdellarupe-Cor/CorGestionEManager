 <!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
 <%


 ok = ofv.CheckUsuario()
 If ok <> "Y" Then
 	uv = ofv.uv
 	If uv = "NO" Then ok = ofv.ValidarUsuario
 	shtml = ok
 Else
 	response.expires = 0
 	shtml = ofv.FormHeader(SESSION("FormN"))

     If (SESSION("UN") & "") = "" Then


 	  xxxx = NEWEMPRE("NO")


 	  SEMPRE = ATREMP(0, 0)
 	  SESSION("UN") = SEMPRE & ""

 	  If (SESSION("UN") & "") = "" Then SESSION("UN") = 1


 	End If
	Set cn = ofv.conectar(ofv.strconn2)

  VCOLOR   = request.querystring("VCOLOR")
  VVUELT   = request.querystring("VVUELT") & " "

  IF VVUELT <> "" THEN
     VVUELT = REPLACE(VVUELT,"_","&")
  ELSE
     VVUELT = "CHGCOLOR.ASP"
  END IF

  id=request.querystring("id")
  CODFAM=request.querystring("CODFAM")


sql="SELECT  B.PRTY3 AS MS3,  B.PRTY6 AS MS6 FROM SYSDEFDSKTOPDET A INNER JOIN SYSDEFDSKTOPDET B ON A.PRTCOD = B.PRTCOD AND B.CODSET = 0 "
sql= sql & " WHERE "
sql=  sql & "A.CODSET = " & SESSION("VSET") & " AND B.PRTY1 = '" & CODFAM &"' and b.prty5 = 'SI' "
sql= sql &  "and a.OBJID=0" & id

 set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
'response.write sql
 item = ""
 if not rs.eof then item = rs(1) & "/" & rs(0) & "</br>"

 call ofv.cerrarconsulta(rs)

SHTML =SHTML & "<CENTER>"
SHTML =SHTML & "<table CLASS=AAA CELLSPACING=2 CELLPADDING=10 ><tr>"
SHTML =SHTML & "<TD VALIGN=MIDDLE ALIGN=left >" & item

SHTML =SHTML & "    <FORM ID=FRMCOL ACTION='" & VVUELT & "' TARGET=_self METHOD=POST >"
'SHTML =SHTML & "    <table ><tr><td><input  id=xsi NAME=xsi type=text  value='#" & VCOLOR &"' ></td><td id=celcol bgcolor=#'"& VCOLOR &"'>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
'SHTML =SHTML & "   </td></tr><tr><td ALIGN=left ><div CLASS=selcolor></div></td></tr>"

'SHTML =SHTML & "    <table ><tr><td><input  id=xsi NAME=xsi type=text  value='#" & VCOLOR &"' ></td><td id=celcol bgcolor=#'"& VCOLOR &"'>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table>"
SHTML =SHTML & "   </td></tr><tr><td ALIGN=left ><div CLASS=selcolor></div></td></tr>"
SHTML =SHTML & "    <tr><td><table ><tr><td><input  id=xsi NAME=xsi type=text  value='#" & VCOLOR &"' ></td><td id=celcol style=border-radius:5px;width:30px;height:30px; bgcolor=#'"& VCOLOR &"'>&nbsp;&nbsp;&nbsp;&nbsp;</td></tr></table></td></tr>"
 SHTML =SHTML & "   <input id=xs2 NAME=xs2 type=HIDDEN value=' ' size=40 style='font-family:tahoma;font-size:8pt;font-color:#336699;border-style:solid;border-width:1;border-color:#dddddd;text-align:center;background-color:#ffffff;' >"

 SHTML =SHTML & "   "

 SHTML =SHTML & "<tr><td><input class=bietx type=SUBMIT value='ACEPTAR' size=40 ></td></tr>"

SHTML =SHTML & "</tr>    </FORM>"
SHTML =SHTML & "</td>"



SHTML =SHTML & "</tR>"


SHTML =SHTML & "</TABLE></center>"
sHTML = sHTML & s & ofv.botoneraESSX("0","","","","","","","","","","","","")
sHTML = sHTML   & "</body></html>"
Call ofv.cerrarconn(cn)
end if


response.write shtml
%>
