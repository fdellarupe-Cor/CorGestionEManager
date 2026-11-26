<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  TARGET = request.querystring("TARGET")
'response.write request.querystring("DATOS")
  if   TARGET = "" then   TARGET = request.querystring("DATOS")

 ' IF LEN(TARGET) > 2 THEN

   ' TARGET = RIGHT(TARGET,LEN(TARGET)-1)
   ' TARGET = LEFT(TARGET,LEN(TARGET)-1)

    TARGET = trim(target)

'  ELSE

'    TARGET = "SELECCION VACIA"

'  END IF



  IF TARGET <> "" THEN
     SESSION("GLTARGET") = TARGET
  ELSE
     VMOVER = SESSION("GLTARGET")
  END IF






 ' sHTML = ofv.MENUHeader("","")

' sHTML = replace(sHTML,"<html>","<html><title>Consultas al Glosario</title>")

  sHTML = "<html><title>Consultas al Glosario</title>"
  sHTML = sHTML & "<STYLE>"
  sHTML = sHTML & " input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;}"
  sHTML = sHTML & " td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE",Session("CMFPNTIT"),0) & ";}"
  sHTML = sHTML & " td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
  sHTML = sHTML & " td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
  sHTML = sHTML & " td.ttt{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
  sHTML = sHTML & "</STYLE>"
  sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"">"
  sHTML = sHTML & "<body>"



  'Encabezados

    Celda = ofv.GenCelda("","TTR COLSPAN=2 ","","","","","","CONSULTAS AL GLOSARIO ","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ofv.GenCelda("","ttt","","","","","","BUSQUEDA:&nbsp;","si")

    Celda = Celda & ofv.gencelda("","ttt","","80%","","","","<BIG>" & TARGET & "</BIG>","si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",FILA,"si")

  Formulario = ""





  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From GLOSARIO WHERE KEYWORD LIKE '%" & TARGET & "%' Order By KEYWORD"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi


      Celda = ofv.GenCelda("","","","","","","","<font face=corflaticon size=3 color=#F7AA22 >4</font>","si")


    'Tipo
    sInput = "<B><BIG>" & rs(2) & "</BIG></B>"

  ''  sinput1 = "<a class=nt title='Consultar Descripcion' href='javascript:void(0);' "
  ''  sinput1 = sinput1  & " onclick=" & chr(34) & "verglos('" & rs(0) & "'); return false;" & chr(34) & " >" & sinput & "</a>"

    Celda = Celda & ofv.gencelda("","VT","","100%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & Fila

''''''''''''''''''''''''''''''
  Strsql = "Select * From GLOSDSC WHERE CODWORD = " & rs(0) & " Order By FECVIG,HORVIG,OBJID "
  set rs0 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs0.eof then
  posi = 0
  Do While (Not rs0.EOF)
    sForm = "form" & posi


      Celda = ofv.GenCelda("","","","","","","","<font face=corflaticon size=2 color=#005500 >4</font>","si")

  GLAUTOR = ""

  Strsql = "Select USRWORD From GLOSusr WHERE CODWORD = " & rs(0) & " and CODDSC = " & RS0(0) & " AND TIPMOD = 'A' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

   ''   Strsql = "Select DESCRIPCION From USUARIOS WHERE ID= '" & RS1(0) & "' "
   ''   set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   ' '  if not rs1.eof then

   xx=getinfouser(rs1(0))


         GLAUTOR = "<B><font face=corflaticon size=3 color=#993333 >&#0128;</font></B>&nbsp;" & session("sgu_nombre") & "&nbsp;-&nbsp;"

   ''   END IF

  END IF


    sInput1 = cdate(RS0(2) + clng(cdate("01-01-1900")))

    s1d = datepart("d",sInput1)
    s1m = datepart("m",sInput1)
    s1a = datepart("yyyy",sInput1)

    sInput1 = s1d & "/" & s1m & "/" & s1a

    sInput2 = GETGLBHMS(RS0(3))

    IF LEN(GLBSM) = 1 THEN GLBSM = "0" & GLBSM
    IF LEN(GLBSS) = 1 THEN GLBSS = "0" & GLBSS

    sInput2 = GLBSH & ":" & GLBSM & ":" & GLBSS

    sInput = GLAUTOR & "&nbsp;&nbsp;" & sInput1 & "&nbsp;" & sInput2
    Celda = Celda & ofv.gencelda("","uT bgcolor=#f0f0f0 style='border-style:solid;border-width:1;border-color:#888888;' ","right","100%","","","",sInput,"si")

    Fila = ""

   if  SESSION("GLOSAUTOR")  = "S" then

    Fila = ofv.genrow("","","","","","","",Celda,"si")

   end if


      Celda = ofv.GenCelda("","","","","","",""," ","si")


    'Tipo
    sInput = "<p style='text-align:justify;' >" & REPLACE(rs0(4),CHR(13),"<BR>") & "</p><br>"


    Celda = Celda & ofv.gencelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","","100%","","","",sInput,"si")


    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & Fila

    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")
    Formulario = Formulario & Fila

    posi = posi + 1
    rs0.Movenext
  Loop

  end if
''''''''''''''''''''''''''''''

    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")
    Formulario = Formulario & Fila


    posi = posi + 1
    rs.Movenext
  Loop

 ELSE


      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")



    sInput = "<BR><BR>NO SE ENCONTRARON ELEMENTOS PARA LA SELECCION"
    Celda = Celda & ofv.gencelda("","VT","","100%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & Fila

 END IF


    sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","1","2","","",Formulario,"si")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>