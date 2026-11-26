<%

set conn=Server.CreateObject("ADODB.Connection")

conn.open "Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=a_datgp_final;user id=sa;pwd=bb4515"


set rs= conn.execute("select * from usuarios")


do while not rs.eof

response. write rs(0) & "<br>"
rs.movenext
loop

rs.close
conn.close


%>