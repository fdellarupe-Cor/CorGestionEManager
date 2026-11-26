<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

  XXXX = GENLOGEXT("EDIDET")
TipObjeto=request.querystring("TipObjeto")
CodObjeto=request.querystring("CodObjeto")
%>

<html>
<frameset rows="35,*" FRAMESPACING=0  FRAMEBORDER=1 BORDER=0 >
<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC="Edselect.asp?tipobjeto=<%=tipobjeto%>&codobjeto=<%=codobjeto%>" NAME="Select" SCROLLING=no>
<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME="Transfer" sRC="nulo.asp" SCROLLING=auto>
</frameset>
</body>
</html>