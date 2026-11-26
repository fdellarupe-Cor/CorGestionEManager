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
  OPC     = request.querystring("OPC")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'shtml=replace(shtml,"p{font-family:verdana;color:#888888;font-size:7pt;font-weight:500;}","p{font-family:verdana;color:#888888;font-size:7pt;}")
  shtml=replace(shtml,"p{font-family:verdana;color:#888888;font-size:7pt;font-weight:500;}","")
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

  XNWNOMD = RS(3)



 sHTML = sHTML &  "   <div id='summernote'>" & XNNWFILE & "</div>"
 sHTML = sHTML &  "<button id='guardarBtn' class='btn-success mt-3'>Guardar</button>"

sHTML = sHTML &  "<form  id='MSForm' name='MSForm' style='display:none;' method=post enctype='multipart/form-data'"
sHTML = sHTML &  "   action=""asp/EDNWEDTDET2.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&concod=" & concod &"&OPC=S"" >"
sHTML = sHTML &  "    <textarea style='xdisplay:none;' id='txt1' name='txt1' cols=100 rows=10></textarea>"
sHTML = sHTML &  "  </form>"
  Agregar = ""
  Buscar  = ""

  sAccion = "EDNWEDT.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Celda = ofv.BotonVolver(sAccion)

  'Boton test
  sAccion = "EDNWEDTTEST.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&concod=" & concod
  test    = ofv.BotonTest(sAccion)






 '''xxxx= ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
    xxxx= ofv.botoneraESSX("0","","","","","","","","",test,Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  '--------------------------------------------------------------------------------
    sHTML = sHTML & "</span></div>"

    Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

    'Estructura del Indice

    sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & xxxx
  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

function NWEDTVER(XNNWCOD,OPC)

  XNVER = "N"
  XNSEC = 0


     xxxfl= getReadfile(XNNWCOD & ".CSE","")
    NWEDTVER = xxxfl

end function

%>