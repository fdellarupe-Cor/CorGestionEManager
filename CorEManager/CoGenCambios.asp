<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
CONT = 0
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  sHTML = "<html><title>Consultas cambios</title>"
  sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css""><STYLE>"
  sHTML = sHTML & "<STYLE>"
  sHTML = sHTML & " input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;}"
  sHTML = sHTML & " td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;line-height:10pt;border-radius:5px;" & XspEx("#EEEEEE",Session("CMFPNTIT"),0) & ";}"
  sHTML = sHTML & " td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
  sHTML = sHTML & " td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
  sHTML = sHTML & " td.ttt{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
  sHTML = sHTML & "</STYLE>"
  sHTML = sHTML & "<body>"

  sHTML = sHTML & "<br><center>"

  sImagen4si = "<FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT>"

  'Encabezado sector

    sInput = "Cambios efectuados al documento"
    Celda  = ofv.GenCelda("","ttr","center","100%","","","",sInput,"si")
    Fila   = fila & ofv.GenRow("","","","","25","","",Celda,"si")

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT * FROM OBJgencam where codobjeto = " & nCodigo & " order by secuen "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
	CONT = CONT +1
    sInput = "&nbsp;"
    Celda  = ofv.GenCelda("","UT","","100%","","","",sInput,"si")
    Fila   = fila & ofv.GenRow("","","","","20","","",Celda,"si")

    sInput = sImagen4si & "&nbsp;" & rs(3)
    Celda  = ofv.GenCelda("","UT style='text-align:justify;' ","","100%","","","",sInput,"si")

    Fila   = fila & ofv.GenRow("","","","","","","",Celda,"si")



    rs.movenext
  loop

    sInput = "&nbsp;"
    Celda  = ofv.GenCelda("","UT style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#993333;' ","","100%","","","",sInput,"si")
    Fila   = fila & ofv.GenRow("","","","","20","","",Celda,"si")

    if CONT = 0 then
	    	    sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes cambios al documento agregados.</div>"
	    	else
	    	    sHTML = SHTML & ofv.GenTabla("","","","95%","","0","","-2","-1","","",Fila,"si")
	    	    'sHTML = sHTML & tabla
	end if

  ''''sHTML = sHTML & ofv.GenTabla("","","","95%","","0","","-2","-1","","",Fila,"si") reemplazo por validacion en caso de no tener registros

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>