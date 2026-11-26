<html>
<% op=request.querystring("op") %>

<%if op="Y" then %>
 <form action="result.asp" method="post" ID="Form1" target=_new>
<textarea name="sql" cols="100" rows="5" ID="sql"></textarea>
<br>
<input type="submit" value="Enviar SQL" ID="Submit1" NAME="Submit1">
</form>

Select   syscolumns.name as columna,sysobjects.name as tabla
from syscolumns,sysobjects where (sysobjects.id=syscolumns.id) and
(sysobjects.xtype ='U')


<br>
<br>
Select   sysobjects.name as tabla
from sysobjects where sysobjects.xtype ='U'
<% end if%>
</html>