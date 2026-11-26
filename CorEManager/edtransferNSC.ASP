<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

Xarchivo = request.querystring("archivo")
XarchivoV = Xarchivo



set fff = server.createobject("scripting.filesystemobject")
xarchivo = ofv.convertircar(xarchivo,chr(34)," ")
Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)

xxx = ""

do while fff1.AtEndOfstream <> true 
  xxx = xxx & fff1.readline
loop 

fff1.close

  nn = 0
  i = 1
  nx = 80
    redim preserve x(0)
  do until i > len(xxx)
    redim preserve x(nn)
    x(nn) = mid(xxx,i,nx) 
    nn = nn + 1
    i = i + nx
    if nx < 80 then exit do
    if (i + nx)  > len(xxx) then
      nx = (len(xxx) - i) + 1
      if mid(xxx,nx + i -1 ,1) <> "}" then nx = nx -1
    end if
  loop 

xa = Request.ServerVariables("SERVER_NAME")
xb = Request.ServerVariables("SCRIPT_NAME")
xe = instr(Ucase(xb),"COREMANAGER")
xd = xa & mid(xb,1,xe-1)

'classid="clsid:B088D12E-6140-427C-99AE-19325DF8650F"
'codebase="http://server-a/CorEditNsc/CorEditNSC.ocx#version=1,0,0,13"
'classid="clsid:AAC33A2F-2640-42E7-AC28-1B3A750BC678"
'codebase="http://server-a/Coredit/Coredit.ocx#version=2,0,24">
%>

<html><meta HTTP-EQUIV="expires" CONTENT="0" >
<embed type="application/x-eskerplus" id="rtf" name="rtf"
classid="clsid:B088D12E-6140-427C-99AE-19325DF8650F"
codebase="http://<% response.write trim(xd) %>CorEditNsc/CorEditNSC.ocx#version=1,0,0,13"
width=710 height=400 ></embed>

<script language=javascript>

function guardar(arc) {
var wdatos;
wdatos= rtf.TomarStr();
document.archivo.datos.value = wdatos;
document.archivo.nombre.value = arc;
document.archivo.submit();

}
</script>

<body bgcolor=#ccbbaa language=javascript onLoad="Init()">
<script language=javascript>
    function Init ()
    {
        if (window ["rtf"])
        {
            // the activex control is activated
            abrir();           
 
        }
        else
        {
            // the activex control is not activated yet, retry later...
            setTimeout ("Init()", 100);
        }
    }
</script>

<br><center>
<form name=ientra  method=post>
<% 

response.write "<input name=idatos type=hidden value='"
for nn = 0 to Ubound(x)
response.write x(nn)
 xn = xn + 1
 next 
response.write "' maxlength=64000 >"
%>
</form>
<form name=Actual method=post>
<input type=button value=Actualizar onclick=guardar(<% response.write "'" & ofv.convertircar(trim(xarchivo),"\","_") & "'" %>); >
</form>
<form name=archivo action=edserverx.asp?archivo=<% response.write trim(XarchivoV) %>  method=post>
<input name=nombre type=hidden value=x >
<input name=datos type=hidden value=x maxlength=64000 >
</form>



<script language=javascript>
function abrir() {

rtf.AbrirStr(document.ientra.idatos.value);
}
</script>
</body>
</html>
