<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

txt = request.querystring("txt")

set cn = ofv.conectar(ofv.strconn2)

If session("UN") <> 1 Then

	sql="select b.eco,b.descri from eco b "
	sql = sql & "where b.descri like '" & txt & "%' and b.estado=3 and b.ecoun=0" & session("UN") & " and b.eco not in "
	sql = sql & "(select codsector from objsector where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "') "
	sql = sql & "order by b.descri"

Else

	XEMPRE = Split(session("EmpHbl"),Chr(9))
	DELIM=""
	idempre = ""

	For i=0 To UBound(XEMPRE)

		idempre = idempre & DELIM & XEMPRE(i)
		DELIM = ","

	Next

	If idempre = "" Then idempre = session("CLIEMP")

	sql="select b.eco,b.descri from eco b "
	sql = sql & "where b.descri like '" & txt & "%' and b.estado=3 and b.eco not in "
	sql = sql & "(select codsector from objsector where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "') "
	sql = sql & "order by b.descri"

End If

set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

sTodos = ""
sPos = 1
sIcon = "<font style='font-size:13pt;'><span class='la la-bank'></span></font>"

do while not rs.eof

	if sPos <= 15 then
		sTodos = sTodos & "<li class='l' data-id='" & rs("eco") & "' data-pos='" & sPos & "'>" & sIcon & "&nbsp;" & rs("descri") & "</li>"
	else
		sTodos = sTodos & "<li class='l' data-id='" & rs("eco") & "' data-pos='" & sPos & "' style='display:none;'>" & sIcon & "&nbsp;" & rs("descri") & "</li>"
	end if

	sPos = sPos + 1

rs.movenext
loop

if (sPos - 1) < 15 then
	sLista = "<ul id='listados' data-posAct='" & (sPos - 1) & "' data-max='" & (sPos - 1) & "'>" & sTodos & "</ul>"
else
	sLista = "<ul id='listados' data-posAct='15' data-max='" & (sPos - 1) & "'>" & sTodos & "</ul>"
end if

response.write sLista
%>
