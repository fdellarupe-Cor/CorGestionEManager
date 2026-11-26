

<TITLE></TITLE>
<body bgcolor="#FFFFFF">
<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->
<%

mySQL=request.form("sql")
'mySQL=UCASE(mySQL)
'On Error Resume Next

set cn = ofv.conectar(ofv.strconn1)


Set rstemp = ofv.crearconsultaEx(mySQL,cn,1,parametros)
If rstemp.eof then
    response.write "No records matched<br>"
    response.write mySQL & "<br>So cannot make table..."
    Call CloseAll
    response.end
end if

response.write "<table border='1'><tr>"
'Put Headings On The Table of Field Names
for each whatever in rstemp.fields
    response.write "<td><b>" & whatever.name & "</B></TD>"
next
response.write "</tr>"

do while not rstemp.eof

Response.Write("<TR>")

for I = 0 to rstemp.fields.count-1
 Response.Write("<TD>" & rstemp(i)  & "&nbsp;</TD>")
next
Response.Write("</TR>")
rstemp.movenext
loop
Response.Write("</TABLE>")




Call CloseAll


SUB CloseALL
    rstemp.close
    set rstemp=nothing
    cn.close
    set cn=nothing
END SUB
%>
</body></html>









