<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  APLCHECK()

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False


  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  IDTXT = request.querystring("IDTXT")
  vuelve = request.querystring("vuelve")
  vuelve = replace(vuelve,"***","_")
  volver = "_Vuelve=" & replace(vuelve,"_","***")
  addqry = replace(vuelve,"_","&")

  sform="modtxt"

  wposi = 0
  wpossal = 0

  vuelta = "../edestrdetTXT.asp?IDTXT=" & IDTXT & volver
  dim jsHashPefiles
  jsHashPefiles="hashperfiles"

  dim jsHashsectores
  jsHashsectores="hashsectores"
SESSION("XGUITEMCD")="NO"
SESSION("FBUSR")="NO"
    set cn = ofv.conectar(ofv.strconn2)



  eststt = " STYLE='font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;' "
  estst1 = " STYLE='font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;' "
  estst2 = " STYLE='font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#88bbee;' "
  estst3 = " STYLE='font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#77aadd; ' "
  estst4 = " STYLE='font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#6699cc; ' "
  estst5 = " STYLE='font-family:tahoma,arial;color:#eeeeee;font-size:12pt;background-color:#5588bb;' "
  esttxt = " STYLE='font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#ffffff;border-style:outset;border-width:2;' "

'f7eedd

  sHTML = ofv.FormHeader(session("FormN"))
  if session("APLMDCLI") =  "BCRx" or session("APLMDCLI") = "BGAL" OR session("APLMDCLI") = "BANELCO" then
  shtml=shtml & "<script type='text/javascript' src='hashtable.js'></script>"
  shtml=shtml & "<script type='text/javascript'>"&armaperfiles(jsHashPefiles,cn,nCodigo)& armasectores(jsHashsectores,cn,nCodigo)&"</script>"
  end if

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

response.write sHTML

  sHTML = ""

  if session("APLMDCLI") =  "BCRx" or session("APLMDCLI") = "BGALx" OR session("APLMDCLI") = "BANELCO" then
    salvar=" <button id='bt1'  class=bietx onclick=" & chr(34) & "insdato('spanpf');" & chr(34) & ">Guardar</button>"
   ' elseIF ISNUMERIC(SESSION("XGUITEMCD")) THEN
 end if

  response.write "<center>"

  response.write "<TABLE class='aaa' width=60%  marginwidth=0  cellspacing=0  cellpadding=5 >"



    'null
     sicK = "<i class='las la-file-alt' style='font-size:12pt' ></i>&nbsp;&nbsp;"
     Celda =  ofv.GenCelda("","ttT valign=top  ","","","","","",sicK & "<b>Modificacion de Texto de contenido</b>&nbsp;&nbsp;&nbsp;&nbsp;" ,"si")
     Celda =  celda & ofv.GenCelda("","xxx","","","","","", salvar,"si")

     Fila = ofv.GenRow("","","","","","","",celda,"si")
     response.write Fila




  strsql = "SELECT * from Objestdet where objid = " & IDTXT
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if  not rs.eof then

      wasotxt = "N"
      stylob  = ""


         if not isnull(rs(32))  then
            if  clng(rs(32)) > 0 then

                Strsql = "SELECT * From OBJLNK where objid = 0" & rs(32)
                set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsx.eof then
                   wasotxt  = replace(" " & rsx(14) & ""," ","&nbsp;")
                   IF wasotxt = "" THEN wasotxt = "N"
                end if
                call ofv.cerrarconsulta(rsx)

            end if
          end if

       IF wasotxt = "S" THEN

          stylob = "UT valign=middle align=center style='border-style:solid;border-width:1;border-color:#990000;' "

       end if


      'Texto
	textoe= " " &  trim(rs(8) &"")
	'response.write "222222222222222" & textoe

      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & IDTXT & "&Tabla=Objestdet&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


      end if

	if  session("APLMDCLI") <>  "BCRx" AND session("APLMDCLI") <> "BGAxL" AND session("APLMDCLI") <> "BANELCO" then

    sInput = "<font class=at><textarea class=dt name=" & rs(8).name & "  id=" & rs(8).name & " rows=20 cols=120 "
    if bModificar and wasotxt = "N" then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 3990) {alert('El texto es demasiado"
       sInput = sInput & " extenso. ' ); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" & trim(rs(8)&"") & "</textarea></font>"
        sInput = ofv.GenForm(sForm,"",sAccion,"","post","",sInput,"si")

      Celda =  ofv.GenCelda("",stylob,"","","","","",sInput,"si")
	else
    sInput = "<font class=at><textarea style='display:none;' class=dt name=" & rs(8).name & "  id=" & rs(8).name & " rows=10 cols=60 "
    if bModificar and wasotxt = "N" then
       'sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
      ' sInput = sInput & " extenso'); return false;} else {"
      ' sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    '  sInput = sInput & " onblur=asignar('"& sform & "');"

    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" & textoe & "</textarea></font>"
        sInput = ofv.GenForm(sForm,"",sAccion,"","post","",sInput,"si")
      'textoe=replace(textoe," ","&nbsp;")
      Celda =  ofv.GenCelda("",stylob & " ","","","","","",sInput & AGREGAREDITOR(REPLACE(textoe & "","''","'")),"si")

    end if








      Fila  = ofv.GenRow("","","","","","","",Celda,"si")

      wstextoLNK = rs(9) & ""

       if  isnumeric(wstextoLNK) and wstextoLNK <> "0" then
           sImagenP = "<font face='corflaticon' size=2 color=#005500 >&#0252;</font>"
       else
           sImagenP = "<font face='corflaticon' size=2 color=#f7aa22 >&#0252;</font>"
       end if

       sHref = "javascript:void(0);"

      sID = "edestrdetLnk.asp?"
      sID = sID & "IDTXT=" & IDTXT & "&Tabla=Objproced&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHref = sID

       sinputk = "<a class='ft1o'  href='" & sHref
       sinputk = sinputk & "' title='Actualizar link a otro documento' >" & sImagenP & "</a>"
       slink   = "&nbsp;" & sinputk

      Celda = ofv.GenCelda("","ft","left","5%","","","",sLink,"si")
 '     fila  = fila & ofv.GenRow("","aaa   ","","","","","",celda,"si")



            response.write Fila
'XXBSQ1= GenBuscaUsrEx("N","N","N","n","N","N","N","S","N","N","","","","","","","","","","")
      sID = "edestrdetTXTARMA.asp?"
      sID = sID & "IDTXT=" & IDTXT & "&Tabla=Objproced&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


   Whtml =   "<FORM NAME='FSEC' TARGET=WK4 ACTION='" & SID & "' METHOD='POST' >"

  Whtml = Whtml & "</FORM>"
 Whtml = Whtml & "<DIV id='EXBSUSR' class='DINDM'  STYLE='display:zblock;z-index:99;margin-top:1;margin-left:1;width:650;"
 Whtml = Whtml & "'></DIV>"

 Whtml = Whtml & "<script language=javascript >"
 Whtml = Whtml & " function clactfrm(XGU0,XGU1,XGU2,XGU3,XGU4) { " & chr(13)
 Whtml = Whtml & "    document.getElementById(XGUMASTER).value=XGU0; " & chr(13)
 Whtml = Whtml & "    document.getElementById(XGUFILTRO).value=XGU1; " & chr(13)
 Whtml = Whtml & "    document.getElementById(XGUITEMCD).value=XGU2; " & chr(13)
 Whtml = Whtml & "    document.getElementById(XGUCODIGO).value=XGU3; " & chr(13)
 Whtml = Whtml & "    document.getElementById(XGUCODDSC).value=XGU4; " & chr(13)
 Whtml = Whtml & "  wid=document.getElementById(XGUITEMCD).value;wid2=document.getElementById(XGUCODIGO).value;} " & chr(13)

 Whtml = Whtml & "</script>"
  XXBSQ1=Whtml


CELDA =ofv.GenCelda("","vt","","60%","","","",XXBSQ1   ,"si")


RESPONSE.WRITE ofv.GenRow("","class=xxxx id=filaPerf  style='display:none;' ","","","","","",CELDA  ,"si")




end if


response.write ofv.GenFTabla()
  response.write "</center>"
  ' RESPONSE.WRITE" <button id='bt2'  onclick=""asignar('"&sform&"');"">actualizar</button>"

 'if ISNUMERIC(SESSION("XGUITEMCD")) and ISNUMERIC(SESSION("FBUSR"))  then

'  RESPONSE.WRITE" <button id='bt1'  onclick=""insdato('spansec');"">ACEPTAR</button>"
' end if
  if  SESSION("FBUSR") ="NO" AND SESSION("XGUITEMCD")="NO" AND FALSE THEN
    RESPONSE.WRITE  "<SCRIPT>"
    RESPONSE.WRITE  "var ele = document.getElementById('bt1');"
    RESPONSE.WRITE  "ele.style.visibility = 'hidden';"
    RESPONSE.WRITE   "</SCRIPT>"
  END IF

  Agregar = ""
  Buscar = ""

  'Boton Volver
  sAccion   = "edestrdet.asp?" & addqry
  sRetornar = ofv.BotonVolver(sAccion)





  sHTML = ""
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmover,cantregmostrar,"no","no","","","","",Agregar,Buscar,sRetornar)

response.write sHTML
ss="<script> var wid=0;var wid2=0;"

ss=ss & "function insdato(classx){"

ss=ss & " if( wid==0 && (wid2=='NO' || wid2==0) ) { var es=document.getElementById('idzz#') ;if (es  != null){es.outerHTML=''; }}"
ss=ss & " var e=document.getElementById('foo')  ;"
ss= ss & "var ele=getElementsByClassName(e,'spansec');"

ss= ss & "     for(i=0;i< ele.length;i++)"
ss= ss & "     {if(ele[i].id==''){ele[i].outerHTML='';}"

ss= ss & "     }"


ss= ss & "var ele=getElementsByClassName(e,'spanpf');"

ss= ss & "     for(i=0;i< ele.length;i++)"
ss= ss & "     {if(ele[i].id==''){ele[i].outerHTML='';}"

ss= ss & "     }"


ss=ss & " var es=document.getElementById('idzz#');if (es  != null){ "
ss=ss & " if (wid2 =='NO') {es.id='idsec' + wid  ;"
ss=ss & "es.name='idsec' + wid  ;"
ss=ss & "es.setAttribute('className', 'spansec') ;}"
SS=SS & "else { if (wid2 !=0){"
SS=SS & " es.id='pf' + wid2  ;es.name='idsec' + wid  ; es.setAttribute('className', 'spanpf') ;}"
SS=SS & "}"

SS=SS & "}"
ss=ss & " var tt=e.innerHTML;tt=tt.replace(/<P[^>]*>/gi, '');tt=tt.replace(/<\/P>/gi, '<br/>');tt=tt.replace(/&nbsp;/gi,' ');"

ss=ss & " var e2=document.getElementById('dettexto');  e2.innerText=tt ;"
ss=ss & "var e2=document.getElementById('" & sform &"');e2.submit();"
ss=ss & "}"



ss=ss & "</script>"

response.write ss
FUNCTION AGREGAREDITOR (texto)
dim scedit
'texto=replace(texto," " ,"&nbsp;")
scedit="<div  unselectable='on' style='width:720; background-color:#f9f9f9;border-style:SOLID;border-width:2;border-color:'" & session("CMFMNBOT") & ";>"

scedit=scedit & "<table style=""margin-top:0px;margin-bottom:0px;"" cellpadding=0 cellspacing=1 width=92%  >"
if session("APLMDCLI") =  "BCRx" or session("APLMDCLI") = "BGAL" OR session("APLMDCLI") = "BANELCO" then
scedit=scedit & "<tr>"
'scedit=scedit & "<td CLASS=TTT width=50% >&nbsp;</td>"
scedit=scedit & "<td>  <img  src=""imgeditor/mn_createlink.gif"" id=E1 onclick=ak_ocultar(""filaPerf"");''"
scedit=scedit &   " title=""crear SECTOR/PERFIL"">"
scedit=scedit &"  </img></td>"



scedit=scedit & "</tr>"
end if
scedit=scedit & "</table>"

scedit=scedit & "<div id=foo onsKeyUp='_onKeyUp()' oncliqck='getCaretCharacterOffsetWithin(this);'  onfgocus='getCaretCharacterOffsetWithin(this)'  ognblur=asignar('"& sform & "');  name=foo class=nzzz contenteditable align=left style=""height:150; width:700;background-color:white;"
scedit=scedit & " padding:3; border-style:SOLID;border-width:1;border-color:#aaaaaa;overflow=auto;"">"
scedit=scedit & texto
 scedit=scedit & " </div>"
scedit=scedit & "<br style=""font-size:5pt;"" >"
scedit=scedit & "</div>"

AGREGAREDITOR=scedit
END FUNCTION



%>
<script>

var caretOffset=0 ;





function getCookie(cname)
{
var name = cname + "=";
var ca = document.cookie.split(';');
for(var i=0; i<ca.length; i++)
  {

 //var c = ca[i].trim();
  var c = ca[i];
  if (c.indexOf(name)==0) return c.substring(name.length,c.length);
  }
return "";
}


var editor='foo';
var tr;
function asignar(f)

	{
	      alert('10');
	      change_span('&nbsp;',hashperfiles,'spanpf');
	      change_span('&nbsp;',hashsectores,'spansec');
	      var tt=document.getElementById(editor).innerHTML;
//alert(caretOffset);
//	     tt=tt.replace(/<P[^>]*>/gi, '');
        tt=tt.replace(/<\/P>/gi, '<br/>');
 //        tt=tt.replace( /"/g, '');
        // tt=tt.replace(/&nbsp;/gi,' ');
  	      document.getElementById("dettexto").value=tt;
  	   if (document.getElementById('dettexto').value.length==0) {document.getElementById('dettexto').value=' ';};
        //  BreakItUp();

        var e2=document.getElementById(f); e2.submit();
	 // return false;
	}
function getElementsByClassName(node, classname) {
    var a = [];
    var re = new RegExp('(^| )'+classname+'( |$)');
    var els = node.getElementsByTagName("*");
    for(var i=0,j=els.length; i<j; i++)
        if(re.test(els[i].className))a.push(els[i]);
    return a;
}
	function ak_ocultar(obj)
			{
            document.getElementById('foo').focus();


			var t= document.getElementById(obj);
			//style={display:none;}
			if (t.id=='filaPerf') { ak_inserttags('&nbsp;<span class=classx name=idseczz id=idzz#>X','</span>&nbsp;');}
			if (t.style.display=='none')
			{
			t.style.display='';
			if (t.id=='filaPerf') {
            var t= document.getElementById('FSEC');

            t.submit();
             }
			}
			else
			{
			t.style.display='none';

			}
	}


function _onKeyUp()
{
if (window.event.keyCode == 13) {
var sel = document.selection.createRange();
sel.pasteHTML('<br>');
sel.moveEnd();
}
return true;
}

function validar(e) {
  tecla = (document.all) ? e.keyCode : e.which;
  if (tecla==13) alert ('Has pulsado enter');
}

function ak_inserttags(str1, str2) {
          //  alert(str2);
		  	document.getElementById('foo').focus(); alert('0');
		  	var tr  = document.selection.createRange().text;     alert('1');
        var tr2 = window.getSelection.createRange().text;    alert('2');
        alert(document.getElementById('foo').innerHTML);
			if (str2 != ""){tr.pasteHTML(str1 + tr.text + str2);}
			else {tr.pasteHTML(str1);}

		  	tr.select();
		  	document.getElementById('foo').focus();

	}

</script>
<%
response.write  "</body><script>change_span('',"& jsHashPefiles &",'spanpf');change_span('',"& jsHashsectores &",'spansec');</script></html>"




  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




end if



%>