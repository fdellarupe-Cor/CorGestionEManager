<!-- #INCLUDE FILE="coress20/IncFile/CorConect.inc" -->

<%
'response.write OFV.strconn2
Set cn = OFV.conectar(OFV.strconn2)
sql="select  * from objetos"

set rs3=sqlejec3(sql,cn, t2)

do while not rs3.eof
response.write rs3("descri") & "<br>"
rs3.movenext
loop

set rr=rs3.NextRecordset()
  t=rr(0)
response.write "<br>total de registros=" & t
function sqlejec3(sql,cn,byref t)

STRSQLS= "PRC_EXECSQLRC '" & Replace(SQL,"'","''") & "'"
'response.write STRSQLS

Set VV = cn.execute(STRSQLS)

	If cn.errors.count = 0 Then
	  '   set rr=vv.NextRecordset()
	   '  t=rr(0)
		End If

 set  sqlejec3=VV
end function

%>