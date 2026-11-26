<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires = 0

randomize

sID = request.querystring("ID")
sitem = request.querystring("sitem")
stitem = request.querystring("stitem")
stesq = request.querystring("stesq")
CantRegMoverxx = request.querystring("CantRegMover")
sNow = ofv.convertircar(now," ","a") & ofv.convertircar(rnd,",","b")
sNow = ofv.convertircar(sNow,"-","c")
sNow = ofv.convertircar(sNow,":","d")
SUNI = sNow

%>

<html><frameset cols="200,*">
<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC="OBJtreeSubTema.asp?Tipo=ESQUEMAS&ID=<%=sID%>&sitem=<%=sitem%>&stitem=<%=stitem%>&sticono=xxx&stnivel=1&CantRegMover=<%=CantRegMoverxx%>&stesq=<%=stesq%>" NAME="ftemas" SCROLLING=auto>
<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME="fesquemas" sRC="nulo.asp" SCROLLING=auto>
</frameset></body></html>