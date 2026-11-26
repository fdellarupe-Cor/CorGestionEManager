<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else



  session("sqlcons")=""
  session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

  seg = request.querystring("seg")

 ofv.ObtenerAtributos seg,""
 bAgregar   = ofv.AAgregar
 bModificar = ofv.AModificar
 bSuprimir = ofv.AEliminar




base=request.querystring("BASE")

    IF base <> "" THEN
        SESSION("BASE")=base
    END IF


  XXXX = NEWEMPRE("NO")

  session("ASP")="SUBCONSULTA.ASP"

  sEMPRE = request.querystring("EMPRE")
  IF SEMPRE = "" THEN
     SEMPRE = ATREMP(0,0)
  ELSE
     XXXX = NEWEMPRE(SEMPRE)
  END IF

  sUN = SEMPRE

  if sUN = "" then sUN = 1
  session("UN") = sUN



  set cn = ofv.conectar(ofv.strconn2)



  sNombren = sNombre


  sHTML = ofv.FormHeader(session("FormN")) & chr(13)
   shtml= shtml &  "<script src='incfile/mascara.js' type='text/javascript' language='JavaScript'></script>" & chr(13)

  sImagenM = "<B><FONT FACE=""WEBDINGS"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""WEBDINGS"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI = "<B><FONT FACE=""WINGDINGS"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""WEBDINGS"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=4>x</FONT></B>"


  Fila    = ""
  Celda   = ""
  sHref   = ""



  Fila = ""

  formx      = ""
  Formulario = ""

  Fila0 = ""

  sInput = "<h4>""ABM FORMULARIOS""</h4>" & "<HR>"

  'sInput = "<h3>Panel de filtros para la consulta Basica</h3>"
  Celda  = ofv.GenCelda("","istt colspan=4  xstyle='border-style:solid;border-width:1;border-color:#dddddd;' ","","","","","",sInput,"si")
  Fila   = ofv.genrow("","","","","30","","",Celda,"si")

  Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
  Celda  = Celda & ofv.GenCelda("","vti colspan=4 ","","70%","","","","&nbsp;","si")
  Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

'''   fila0 = ofv.GenTabla("", "aaa XSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")
   fila0 = fila
  '' SHTML=SHTML &  fila0

   if bAgregar  then
    Celda      =ofv.GenCelda("","istt","","25%","","","","ALTA&nbsp;DE&nbsp;FORMULARIO","si")
    sInput     = "&nbsp;" '<input type='text' class='dt' name='BUSCAR' id='textot' xonkeyup=mascara(this,'-',patron,true) size='20' style='border-style:solid;border-width:1;border-color:#666666;' >"
    Celda  =  CELDA & ofv.GenCelda("","istt","","20%","","","",sInput,"si")
    sInput     ="&nbsp;&nbsp;&nbsp;&nbsp;<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('ALTAFORM').submit();"">"
    Celda      =Celda & ofv.GenCelda("","istt","","25%","","","",sInput,"si")
   ' Celda      =Celda & ofv.GenCelda("","istt","","25%","","","","&nbsp;","si")
   Celda      = Celda & ofv.GenCelda("","istt","","10%","","","","&nbsp;","si")
    Fila1       =  ofv.GenRow("","","","","30","","",Celda,"si")
   ' formx    = Fila
    sTARGET  = ""
    sAccion    = "EDITARDOC.asp?acc=PF&pedido=ALTA&seg=" & seg &"&docmodal=CL"
    FILA1=  ofv.GenForm("ALTAFORM","",sAccion,"","post","",fila1,"si")
   ' SHTML=SHTML & ofv.GenTabla("", "aaa border' ", "", "50%", "", "0", "", "-2", "0", "", "", Fila, "si")

   end if

  if  bModificar then

   Celda      = ofv.GenCelda("","istt","","25%","","","","<b>MODIFICAR&nbsp;FORMULARIO&nbsp;(*)</b>","si")
   sInput     = "<input type='text' class='dt' name='BUSCAR' id='textot' xonkeyup=mascara(this,'-',patron,true) size='20' style='border-style:solid;border-width:1;border-color:#666666;' >"
   Celda      = Celda & ofv.GenCelda("","istt","","20%","","","",sInput,"si")
   sinput=  "&nbsp;&nbsp;&nbsp;&nbsp;<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform2').submit();"">"
   Celda      = Celda & ofv.GenCelda("","istt","","25%","","","",sInput,"si")
   Celda      = Celda & ofv.GenCelda("","istt","","10%","","","","&nbsp;","si")
   Fila2       =   ofv.GenRow("","","","","30","","",Celda,"si")

  ' formx    = Fila

   sAccion="EDITARDOC.asp?acc=PF&pedido=MODI&seg=" & seg & "&docmodal=CL"
   sTarget = ""
   Fila2 =   ofv.GenForm("Subconsultaform2", "", sAccion, sTarget, "post", "", fila2, "si")
 '  SHTML=SHTML & "<br>"& ofv.GenTabla("", "aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "50%", "", "0", "", "-2", "0", "", "", Fila, "si")
  end if



   if   bSuprimir then

        Celda      = ofv.GenCelda("","istt","","25%","","","","<b>BAJA&nbsp;DE&nbsp;FORMULARIO&nbsp;&nbsp;(*)&nbsp;&nbsp;&nbsp;&nbsp;</b>","si")
	    sInput     = "<input type='text' class='dt' name='BUSCAR' id='textot' xonkeyup=mascara(this,'-',patron,true) size='20' style='border-style:solid;border-width:1;border-color:#666666;' >" '& "&nbsp;&nbsp;&nbsp;&nbsp;<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform3').submit();"">"
	     Celda      = Celda & ofv.GenCelda("","istt","","20%","","","",sInput,"si")
	    sinput=  "&nbsp;&nbsp;&nbsp;&nbsp;<img src='IMAGENES/LUPA.GIF'  onclick=""document.getElementById('Subconsultaform3').submit();"">"
        Celda      = Celda & ofv.GenCelda("","istt","","25%","","","",sInput,"si")
	    'Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
	    Celda      = Celda & ofv.GenCelda("","istt","","10%","","","","&nbsp;","si")
	    Fila3       =   ofv.GenRow("","","","","30","","",Celda,"si")

	    'formx    = Fila3

	    sAccion="EDITARDOC.asp?acc=PF&pedido=BAJA&seg=" & seg & "&docmodal=CL"
	    sTarget = ""
	    Fila3 =  ofv.GenForm("Subconsultaform3", "", sAccion, sTarget, "post", "", FILA3, "si")
 '  SHTML=SHTML & "<br>"& ofv.GenTabla("", "aaa sSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "48%", "", "0", "", "-2", "0", "", "", Fila, "si")
  end if

	sInput     = "<BR><BR>(*) Importante: el número de Formulario debe integrarse de forma completa, es decir, Número más dígito verificador. Ejemplo: 4460-4."
        Celda      = ofv.GenCelda("","istt colspan=4 style='text-align:justify;' ","","25%","","","",sInput,"si")
        Fila4      = ofv.GenRow("","","","","30","","",Celda,"si")


   TABLA= ofv.GenTabla("", "aaa  XSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "",Fila0 &  Fila1 & fila2 & fila3 & fila4, "si")

   celDA=ofv.GenCelda("","istt","","80%","","","",TABLA,"si")
   TEXTO="<FONT TITLE='EL FORMULARIO EDITABLE ULTIMA VERSIÓN <BR> debes solicitarlo al buzón formularios@claridadgalicia.com.ar' SIZE=8 FACE=corflaticon STYLE='CURSOR: hand;border-style:solid;' >s<FONT>"
   CELDA=CELDA &  ofv.GenCelda("","istt ALIGN=CENTER ","","20%","","","",TEXTO,"si")
   Fila       =   ofv.GenRow("","","","","30","","",Celda,"si")
   TABLA1= ofv.GenTabla("", "aaa  STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "58%", "", "0", "", "0", "10", "", "", Fila, "si")

   SHTML=SHTML & "<center><br><br><br>" & TABLA1 & "</center>"
END IF
sHTML = sHTML & "<script>var element=document.getElementById('MTH'); if (typeof(element) != 'undefined' && element != null) {element.style.display='none';}</script></body></html>"
RESPONSE.WRITE SHTML
%>