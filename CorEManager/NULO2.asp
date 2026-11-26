<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires = 0

If BrowserName() = "IE" OR BrowserName() = "CHROME" then
  Xbody = "bgcolor=#FFFFFF onload='uu();Init();'"
else
  Xbody = "bgcolor=#ccbbaa"
end if

sHTML = "<html><Title>COR EManager</title>" & ofv.formstyler()
sHTML = sHTML & "<script language=JavaScript>"
sHTML = sHTML & "function Init () "
sHTML = sHTML & "   { ff(); setTimeout (" & chr(34) & "Init()" & chr(34) & ", 350); }"
sHTML = sHTML & "function uu()"
sHTML = sHTML & " {  xx = 0; }"
sHTML = sHTML & "function ff() "
sHTML = sHTML & "  {   document.getElementById('im1').style.left='-' + xx + 'px'; "
sHTML = sHTML & "  if (xx==450) xx=0; else xx = xx + 50; } "
sHTML = sHTML & "</script>"
sHTML = sHTML & " <BODY " & Xbody & " bacKground='../imagenes/headfon.jpg' ><center>"




sImagen = "<DIV id=im2 name=im2 Style='position:absolute;clip:rect(0px,50px, 30px, 0px)'><img id=im1 name=im1 src=imagenes/coressrot.jpg border=0 style='position:relative;left:-0px'></DIV>"
Celda = ofv.GenCelda("","","center","","","","",sImagen,"si")
Fila =  ofv.GenRow("","","","","","","",Celda,"si")

sHTML = sHTML & "<br><br><br><br><br><br><br>" & ofv.GenTabla("","","","","","4","","0","0","","",Fila,"si")



sInput = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
sInput = sInput & "<font face=verdana color=#000088 size=2>"
sInput = sInput & "<b>COR <font face=helvetica color=#f75500 size=4>"
sInput = sInput & "<i>e</i></font>Solution Suite V4.8</b></font>"

IF SESSION("EMPMUN") = "S" THEN
   sInput = sInput & "<br><br>"
   sInput = sInput & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
   sInput = sInput & "<font face=verdana color=#993333 size=3>"
   sInput = sInput & "<b>Corporate</b></font>"
END IF


Celda = ofv.GenCelda("","","center","","","","",sInput,"si")
Fila =  ofv.GenRow("","","","","","","",Celda,"si")



sHTML = sHTML & "<br><br>" & ofv.GenTabla("","","","","","4","","0","0","","",Fila,"si")
sHTML = sHTML & "</center></BODY></html>"

response.write shtml

%>