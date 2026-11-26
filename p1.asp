<html>
<body>
<iframe src="http://server-40/corv48unif/COREManager/ARCHIVOS/Testing V48unif.xlsx" style="width:600px; height:450px;" frameborder="1"></iframe>
</body>
</html>
<%
'Response.ContentType = "application/application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
'response.redirect "http://server-40/corv48unif/COREManager/ARCHIVOS/Testing V48unif.xlsx"
XLDFINI = clng(date) - CLng(CDate("01-01-1900"))

response.write XLDFINI
%>