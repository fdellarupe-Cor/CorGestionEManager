<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

  response.expires=0
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  Xarchivo = request.querystring("archivo")
  XarchivoV = Xarchivo


  TipObjeto = request.querystring("TipObjeto")
  CodObjeto = request.querystring("CodObjeto")
  Xopc = request.querystring("opc")

  set cn = ofv.conectar(ofv.strconn2)

  select case xopc
    case "1"
      Strsql = "Select ARCHIVO,titulo From ObjMain where CodObjeto=" & CodObjeto & " and TipObjeto='"
      Strsql = Strsql & TipObjeto & "'"
    case "2"
      Strsql = "Select ARCHIVO,titulo From ObjAlcance where CodObjeto=" & CodObjeto & " and TipObjeto='"
      Strsql = Strsql & TipObjeto & "'"
    case "3"
      Strsql = "Select ARCHIVO,titulo From ObjConsid where CodObjeto=" & CodObjeto & " and TipObjeto='"
      Strsql = Strsql & TipObjeto & "'"
  end select

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     TITO=rs(1)
  end if
 '' call ofv.cerrarconsulta(rs)





  set fff = server.createobject("scripting.filesystemobject")
  xarchivo = ofv.convertircar(xarchivo,chr(34)," ")
  xxx = ""
  if fff.FileExists(trim(Xarchivo)) then
  Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)



  do while fff1.AtEndOfstream <> true
    xxx = xxx & fff1.readline
  loop

  fff1.close

  end if

   if xxx="" then ' esto es un rtf vacio
     xxx="{\rtf1\ansi\deff0\deftab720{\fonttbl{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}{\f2\froman Times New Roman;}}"
     xxx=xxx & "{\colortbl\red0\green0\blue0;}"
     xxx=xxx & "\deflang11274\pard\plain\f2\fs20 "
     xxx=xxx & "\par }"
  end if

  nn = 0
  i = 1
  nx = 80

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

'<body bgcolor=#ccbbaa>
'<html>

response.write ofv.formbody

%>


<script language=vbscript >

sub guardar(arc)

wdatos = rtf.TomarStr()
WTxtVacio=rtf.TomarTexto()

document.archivo.datos.value = wdatos
document.archivo.nombre.value = arc
document.archivo.TxtVacio.value=WTxtVacio
document.archivo.submit()

end sub

</script>

<center>
<table  style='filter:shadow(color=#dddddd,direction=125);' cellspacing=5 ><tr><td align=center >
<table cellspacing=5 ><tr><td class=vt><big><% response.write TITO %></big></td></tr></table>


<OBJECT STYLE={border-style:outset;border-width:3;} id="rtf"

CLASSID="CLSID:AAC33A2F-2640-42E7-AC28-1B3A750BC678"
codebase="http://<% response.write trim(xd) %>Coredit/CorEditPlus.CAB#version=2,0,0,6">
</OBJECT>
<script language=javascript>
 rtf.NumBulletIndent=25;
 rtf.Numindent=25;
</script>
<%

if bModificar then
  ' response.write "<form name=Actual method=post>"
  ' response.write "<input class=bt type=button value=ACTUALIZAR "
  ' response.write " onclick=guardar(" &  chr(34) & trim(xarchivo) & chr(34) & ") >"
  ' response.write " </form> "

   cnbot = "<input class=bt type=button value=ACTUALIZAR onclick=guardar(" &  chr(34) & trim(xarchivo) & chr(34) & ") >"
   Cons = "<td align=center width=1% >" & cnbot & "</td>"
   Cons = "<form name=Actual method=post>" & Cons & " </form> "

   fill = space(70)

   fill = replace(fill," ","&nbsp;")

   fill = "<td align=right class=vt align=right width=100% >" &  fill & "<big>Cor&nbsp;<font color=#f7aa44 size=3><i>e</i></font>-Manager&nbsp;</big>&reg;&nbsp;</td>"

   sHTML = "<div  id=odiv name=odiv style='filter:shadow(color=" & Session("APLSHW") & ",direction=125);'  >"
   sHTML = sHTML & "<table style='position:relative;top:+4;filter:shadow(color=#dddddd,direction=125);' width=100% style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ><tr>"
   sHTML = sHTML & Cons & fill & "<td>&nbsp;</td>"
   sHTML = sHTML & "</tr>"
   sHTML = sHTML & "</table>"
   sHTML = sHTML & "</div>"
   response.write sHTML

end if

%>

<form name=archivo action=edserverx.asp?Codobjeto=<% response.write trim(Codobjeto) %>&Tipobjeto=<% response.write trim(Tipobjeto) %>&opc=<% response.write trim(xopc) %>&archivo=<% response.write trim(XarchivoV) %> method=post>
<input name=nombre type=hidden value=x >
<input name=TxtVacio type=hidden  >
<input name=datos type=hidden value=x maxlength=64000 >
</form>
</td></tr></table>
</center>

<script language=vbscript>
dim info
dim p

<%

response.write " xarchivo =" & chr(34) &  xarchivo & chr(34) & " " &  chr(13)
xn = 0

for nn = 0 to Ubound(x)
response.write " redim preserve pepe(" & xn & ")" & chr(13)
response.write "pepe(" & xn & ")  =  " & chr(34) & replace(x(nn),chr(34),"'") & chr(34) & " " & chr(13)
 xn = xn + 1
 next

%>

for ii = 0 to ubound(pepe)
info = info & pepe(ii)
next
info=info & chr(34)
rtf.AbrirStr(info)
</script></body></html>