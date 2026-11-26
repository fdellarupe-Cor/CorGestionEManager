<style type="text/css">
table.mystyle
{
border-width: 1 1 2px 2px;
border-spacing: 0;
border-collapse: collapse;
border-style: solid;
border-color:#99ccff;
}

td.mystyle
{
border-style:none;
border-width:1;
border-color:#888888;
font-family:arial;
font-size:10px;
color:#6699CC;
font-weight:bold;

}
</style>

<%

  response.expires=0

  sHTML = ""
   response.write "<HTML><body>"

   response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br>"
   response.write "<table  width=70% style='border-style:SOLID;border-width:1;border-color:#888888;'>"
   response.write "<tr ><td ALIGN=CENTER><h5><B>NOMBRE</B></h5></td><td ALIGN=CENTER><h5><B>&nbsp;&nbsp;TAMAÑO&nbsp;&nbsp;&nbsp;</B></h5></TD></TR>"
    AFiles=split(session("Sfilename"),"#zzzz#")
    ATam=split(session("Stamano"),"#zzzz#")
    for z=0 to ubound(afiles)-1
      response.write "<tr >"
      response.write "<td class=mystyle>" & UCASE(afiles(z)&"") &"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td><td class=mystyle ALIGN=right>" & ATam(z)& " kb</td>"
      response.write "<tr>"

   next

response.write "</table></center>"


  response.write "</BODY></HTML>"

  %>