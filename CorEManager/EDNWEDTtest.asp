<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  concod  = request.querystring("concod")


  sHTML = ofv.FormHeader(session("FormN"))
  vuelta = ofv.vueltaasp2("../")


  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " and concodigo = '" & concod & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     XNNWFILE = "&nbsp;"

  else
     XNNWFILE = NWEDTVER(concod,OPC)
  end if

  'Encabezados
  Celda = ofv.GenCelda("","ttt","","","","","",RS(3),"si")
  Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")

  response.write shtml

  shtml=""


%>

<STYLE>
    P {margin-top :0px;margin-bottom:0 px;}
    BLOCKQUOTE { line-height:4mm }
    
    .btn     		{ width: 22px; height: 22px; border: 1px solid buttonface; margin: 0; padding: 0; }
    .btnOver 		{ width: 22px; height: 22px; border: 1px outset; }
    .btnDown 		{ border:1px groove orange  }
    .btnDisabled  	{ border:none}
    

</STYLE>


   <%=XNNWFILE%>


<%


  Agregar = ""
  Buscar  = ""
  Celda   = ""
  test    = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",test,Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

function NWEDTVER(XNNWCOD,OPC)

  XNVER = "N"
  XNSEC = 0

  'Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
  'XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
  XferTarget0=""
  XNABRIR = XferTarget0 & XNNWCOD & ".CSE"

  'set fff = server.createobject("scripting.filesystemobject")

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
  '  xxxfl = ""

  '  if fff.FileExists(trim(Xarchivo)) then
  '     Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
  '     do while fff1.AtEndOfstream <> true
  '        xxxfl = xxxfl & fff1.readline
  '     loop
  '     fff1.close
  '  end if 

  '  NWEDTVER = xxxfl
     NWEDTVER = getReadfile(xarchivo,"")
end function

%>