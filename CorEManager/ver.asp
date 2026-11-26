<%
 strconn="Provider=SQLOLEDB;Data Source=server-20;Initial Catalog=DBCORDAT2V34des;"
Set conn = Server.CreateObject("adodb.connection")
 conn.Open strconn, "sadmin" , "manager"

 set rs = conn.execute("select * from usuarios SELECT @@rowcount AS RowsReturned")

response.write clng(cdate("31/12/2002"))
%>