<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

sOCesque = Request.querystring("Codobjeto")
nCodigo = Request.querystring("Codproced")
 sTipo = Request.querystring("Tipobjeto")
xx=request.form("xfname")

response.write xx


 Set cn = OFV.conectar(OFV.strconn2)
st=split(xx,"///")


for i=0 to ubound(st)-1

linea =split(st(i),"###")

sql="update OBJPROCURSODET   set "
sql=sql & " posx=" & replace(linea(1),"px","") & ",posy="  &  replace(linea(2),"px","")
sql =sql &  " where codobjeto='" & sOCesque &"' and codaccion='" &  linea(0) & "'"
'RESPONSE.WRITE SQL & "<br>"
call ofv.crearconsultaEx(sql,cn,1,volver)
next
response.redirect "CursogramaAdmin.asp?op=SI&Codobjeto="&sOCesque & "&Tipobjeto=" & sTipo & "&Codproced=" & nCodigo
%>

