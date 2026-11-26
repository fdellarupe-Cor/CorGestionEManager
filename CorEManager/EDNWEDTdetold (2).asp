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

''  Celda = ofv.GenCelda("","ttt","","","","","",RS(3),"si")
 '' Fila  = ofv.GenRow("","","","","","","",Celda,"si")
''  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")

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

<script src="incfile/jscolor.js" type="text/javascript"></script>
<script language="Javascript1.2">

var editor='foo';
var subdirectory = '';


	function asignar()
	{
  	  document.getElementById("txt1").value=document.getElementById(editor).innerHTML;
          BreakItUp();

	  return false;
	}


function BreakItUp()
{
  //Set the limit for field size.
  var FormLimit = 102399;

  //Get the value of the large input object.
  var TempVar = new String;
  TempVar = document.getElementById("txt1").value;

  //If the length of the object is greater than the limit, break it
  //into multiple objects.
  if (TempVar.length > FormLimit)
  {
    document.getElementById("txt1").value = TempVar.substr(0, FormLimit);
    TempVar = TempVar.substr(FormLimit);

    while (TempVar.length > 0)
    {
      var objTEXTAREA = document.createElement("TEXTAREA");
      objTEXTAREA.name = "txt1";
      objTEXTAREA.value = TempVar.substr(0, FormLimit);
      objTEXTAREA.style.display = 'none';
      document.MSForm.appendChild(objTEXTAREA);

      TempVar = TempVar.substr(FormLimit);
    }
  }
}



	function ak_inserttags(str1, str2) {
		  	document.getElementById(editor).focus();
		  	var tr = document.selection.createRange();
			if (str2 != ""){tr.pasteHTML(str1 + tr.text + str2);}
			else {tr.pasteHTML(str1);}

		  	tr.select();
		  	document.getElementById(editor).focus();
	}


	function ak_underline()
		{
			document.getElementById(editor).focus();
			document.execCommand('Underline', false, null);
		}
		function ak_strikethrough()
		{
			document.getElementById(editor).focus();
			document.execCommand('StrikeThrough', false, null);
			document.getElementById(editor).focus();
		}
	function ak_superscript()
	{
		document.getElementById(editor).focus();
		if (document.queryCommandValue('SuperScript')) {
			document.execCommand('RemoveFormat');
		} else {
			document.execCommand('SuperScript');
		}

	}

	function ak_subscript()
	{
		document.getElementById(editor).focus();
		if (document.queryCommandValue('SubScript')) {
			document.execCommand('RemoveFormat');
		} else {
			document.execCommand('SubScript');
		}

	}

	function ak_bold66(){
		var texto = document.getElementById("foo");
		if (document.selection)
		{
			var str=document.selection.createRange().text;
			//var str=window.getSelection().toString();
			seleccion = document.selection.createRange();
			//seleccion = window.getSelection();
			seleccion.pasteHTML( "\<B\>" + str + "\<\/B\>");
			texto.focus();

				/*
				var str = document.getElementById('LINKTXT').innerHTML;"
				var texto = document.getElementById("foo");


				var position=window.getSelection().getRangeAt(0).startOffset;" 'guard


				var out= [str.slice(0, position), '||||$', str.slice(position)].join(


				position=window.getSelection().getRangeAt(0).endOffset +5 ;" 'guardam


				var out= [out.slice(0, position), '$||||', out.slice(position)].join(


				document.getElementById(xnomdato).value = out;" 'asignamos el string


				 */
		}

	}


	function chagecolor()
	{

	var bgcolor = document.getElementById('bgcolor');

	bgcolor.style.backgroundColor ='#ffffff';
	}

   function setTextColor(picker) {
		highlight( picker.toString());


	}


	function setTextColorf(picker) {
			ak_forecolor( picker.toString());


	}

	function ak_bold()
		{
			//document.getElementById("iframename").contentWindow.document
			//  document.getElementById(editor).focus();
			  document.execCommand('bold');
			//  document.getElementById(editor).focus();



	}
	   /*var oDoc, sDefTxt;


	 function initDoc() {
        oDoc = document.getElementById(editor);
        sDefTxt = oDoc.innerHTML;
       /* if (document.compForm.switchMode.checked) {
            setDocMode(true);
        }
    }
*/
    function formatDoc(sCmd, sValue) {
            document.execCommand(sCmd, false, null);
           // oDoc.focus();

    }


	function ak_italic()
		{
			document.getElementById(editor).focus();
			document.execCommand('Italic', false, null);
			document.getElementById(editor).focus();
	}

	function ak_image()
		{
			var myimage = showModalDialog(subdirectory + "insertimage.html", "", "dialogHeight:230px; dialogWidth: 420px; scroll: no; status: no; help: no;");
			if (myimage){
				ak_inserttags(myimage, "");
			}
		}


	function ak_table()
	{
		var mytable = showModalDialog(subdirectory + "inserttable.html", "", "dialogHeight:160px; dialogWidth: 325px; scroll: no; status: no; help: no;");
		if (mytable){
			ak_inserttags(mytable, "");
		}
	}

	function ak_inserthorizontalrule()
		{
			document.getElementById(editor).focus();
			document.execCommand('InsertHorizontalRule', false, null);
	}


	function ak_backcolor()
	{
	//alert('1');
		//var mycolor = document.getElementById('bn_ak_backcolor').value;
		var mycolor = window.open(subdirectory + "insertcolors2.asp", "", "dialogHeight:230px; dialogWidth: 250px; scroll: yes; status: no; help: no;");
		mycolor=document.getElementById('t').value;
		//alert(document.getElementById('t').value);
		document.getElementById(editor).focus();
	//	mycolor="#66CC33";
	   //document.execCommand("HiliteColor", false, mycolor);
		//document.execCommand('BackColor', true, mycolor);
		highlight(mycolor);
	}


function makeEditableAndHighlight(colour) {
    var sel = window.getSelection();
    var range = null;
    if (sel.rangeCount && sel.getRangeAt) {
        range = sel.getRangeAt(0);
    }
    document.designMode = "on";
    if (range) {
        sel.removeAllRanges();
        sel.addRange(range);
    }
    // Use HiliteColor since some browsers apply BackColor to the whole block
    if (!document.execCommand("HiliteColor", false, colour)) {
        document.execCommand("BackColor", false, colour);
    }
    document.designMode = "off";
}

function highlight(colour) {
    var range, sel;
    if (window.getSelection) {
        // IE9 and non-IE
        try {
            if (!document.execCommand("BackColor", false, colour)) {
                makeEditableAndHighlight(colour);
            }
        } catch (ex) {
            makeEditableAndHighlight(colour)
        }
    } else if (document.selection && document.selection.createRange) {
        // IE <= 8 case
        range = document.selection.createRange();
        range.execCommand("BackColor", false, colour);
    }
}

	function ak_forecolor(mycolor)
	{   document.getElementById(editor).focus();
		document.execCommand('ForeColor', false, mycolor);


	}

	function ak_fontname()
	{
		document.getElementById(editor).focus();
		var name = document.getElementById('bn_ak_fontname').value;
		document.execCommand('FontName', false, name);
	}

	function ak_fontsize()

	{

		document.getElementById(editor).focus();
		var size = document.getElementById('bn_ak_fontsize').value;

		document.execCommand('FontSize', false, size);
	}

	function ak_Undo()
	{

		document.getElementById(editor).focus();
		document.execCommand('Undo');
	}


	function ak_Redo()
	{

		document.getElementById(editor).focus();
		document.execCommand('Redo');
	}

	function ak_insertunorderedlist()
		{
			document.getElementById(editor).focus();
		//	document.getElementById('bn_ak_insertorderedlist').className='btn';
			document.execCommand('insertUnorderedList', false, null);
		}

	function ak_insertorderedlist(e)
	{      // alert(document.getElementById(e.id).className =='btnDisabled');

		document.getElementById(editor).focus();
	//	if (document.getElementById(e.id).className =='btnDisabled')
	//	{
	//	 document.getElementById(e.id).className='btnDown';

	//	}
	//	else
	//	{
	//	document.getElementById(e.id).className ='btnDisabled';
	//	}
		document.execCommand('insertOrderedList', false, null);
	}

	function ak_justifyleft()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyright').className='btn';
		//document.getElementById('bn_ak_justifycenter').className='btn';
		//document.getElementById('bn_ak_justifyfull').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=false;
		//document.getElementById('bn_ak_justifyright').disabled=false;
		//document.getElementById('bn_ak_justifycenter').disabled=false;
		//document.getElementById('bn_ak_justifyleft').disabled=true;
		document.execCommand('justifyLeft', false, null);
	}

	function ak_justifycenter()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyright').className='btn';
		//document.getElementById('bn_ak_justifyleft').className='btn';
		//document.getElementById('bn_ak_justifyfull').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=false;
		//document.getElementById('bn_ak_justifyright').disabled=false;
		//document.getElementById('bn_ak_justifycenter').disabled=true;
		//document.getElementById('bn_ak_justifyleft').disabled=false;
		document.execCommand('justifyCenter', false, null);
	}

	function ak_justifyright()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyleft').className='btn';
		//document.getElementById('bn_ak_justifycenter').className='btn';
		//document.getElementById('bn_ak_justifyfull').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=false;
		//document.getElementById('bn_ak_justifyright').disabled=true;
		//document.getElementById('bn_ak_justifycenter').disabled=false;
		//document.getElementById('bn_ak_justifyleft').disabled=false;
		document.execCommand('justifyRight', false, null);
	}

	function ak_justifyfull()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyleft').className='btn';
		//document.getElementById('bn_ak_justifycenter').className='btn';
		//document.getElementById('bn_ak_justifyright').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=true;
		//document.getElementById('bn_ak_justifyright').disabled=false;
		//document.getElementById('bn_ak_justifycenter').disabled=false;
		//document.getElementById('bn_ak_justifyleft').disabled=false;
		document.execCommand('JustifyFull', false, null);
	}

	function ak_indent()
	{
	document.getElementById(editor).focus();
	document.execCommand('Indent', false, null);
	document.getElementById(editor).focus();
	}

	function ak_outdent()
	{
	document.getElementById(editor).focus();
	document.execCommand('Outdent', false, null);
	document.getElementById(editor).focus();
	}


	function ak_createlink()
		{
		document.getElementById(editor).focus();
		document.execCommand('CreateLink', true, null);
		document.getElementById(editor).focus();
		}

	function ak_unlink()
	{
	myEditor.focus();
	myEditor.document.execCommand('Unlink', false, null);
	}

	function doInit(){alert(document.all.length);
	// Ensure that all document elements except the content editable DIV are unselectable.
	for (i=0; i<document.all.length; i++)
	 document.all(i).unselectable = "on";
	oDiv.unselectable = "off";
	// Clear any text in the Document window and set the focus.
	oDiv.innerHTML="";
	oDiv.focus();
}


</script>



<div  style="width:720; background-color:#f9f9f9;border-style:SOLID;border-width:2;" >

<table style="margin-top:0px;margin-bottom:0px;" cellpadding=0 cellspacing=1 width=92%  >
<tr>
<td CLASS=TTT width=50% >&nbsp;<%=XNWNOMD %>&nbsp;</td>


<td>  <img class="btnDisabled" src="imgeditor/mn_insertorderedlist.gif" id=E1 onclick='ak_insertorderedlist(this);'
    title="Ordered List">
  </img></td>

<td>  <img src="imgeditor/mn_insertunorderedlist.gif" onclick='ak_insertunorderedlist();'
    title="Unordered List">
  </img></td>

<td>  <img src="imgeditor/mn_outindent.gif"  onclick='ak_outdent();'
     title="Outdent">
  </img></td>

<td> <img src="imgeditor/mn_indent.gif" onclick='ak_indent();'
    title="Indent">
  </img></td>

<td> <img  src="imgeditor/mn_justifyright.gif" onclick='ak_justifyright();'
   title="Left Justify">
  </img></td>

<td> <img  src="imgeditor/mn_justifyleft.gif" onclick='ak_justifyleft();'
   title="Right Justify">
  </img></td>

<td> <img  src="imgeditor/mn_justifycenter.gif" onclick='ak_justifycenter();'
   title="Center Justify">
  </img></td>



<td> </td>


<td> <img  src="imgeditor/mn_inserthorizontalrule.gif" onclick='ak_inserthorizontalrule();'
       title="Insert Horizontal Rule">
  </img></td>

<td>
  <input type="image" id=bgcolor2 src="imgeditor/mn_forecolor.gif"  class="jscolor {onFineChange:'setTextColorf(this)'}">

  </td>


<td>
<input type="image" id=bgcolor src="imgeditor/mn_backcolor.gif"  class="jscolor {onFineChange:'setTextColor(this)'}">

</td>


<td> <img  src="imgeditor/undo.gif" onclick='ak_Undo();'
          title="Undo">
  </img></td>

<td>  <img  src="imgeditor/redo.gif" onclick='ak_Redo();'
           title="Redo">
  </img></td>

</tr>
</table>

<table style="margin-top:0px;margin-bottom:0px;" cellpadding=0 cellspacing=1 width=92% >
<tr>
<td width=50% >&nbsp;</td>
<!--
<td><img  src="imgeditor/mn_bold.gif" onclick="ak_bold();" title="Bold"> </img></td>

-->
<!--  http://jsfiddle.net/Q65Qt/1/  -->
<!--
<td><input type="button" class="" value="B" onmousedown="formatDoc('bold'); return false"/></td>

 -->

<td><img  src="imgeditor/mn_italic.gif" onclick='ak_italic();'
    title="Italic">
  </img></td>

<td><img  src="imgeditor/mn_underline.gif"  onclick='ak_underline();'
   title="Underline">
  </img></td>

<td><img   src="imgeditor/mn_strikethrough.gif"  onclick='ak_strikethrough() ;'
    title="Strike-through">
  </img></td>

<td>  <img src="imgeditor/mn_superscript.gif"  onclick='ak_superscript();'
    title="Superscript">
  </img></td>

<td>  <img src="imgeditor/mn_subscript.gif"  onclick='ak_subscript();'
    title="Subscript">
  </img></td>

<td>
<select id="bn_ak_fontname" class="optfont" onchange="ak_fontname();">
	<option value="Arial " selected>Arial</option>
	<option value="Courier New">Courier New</option>
	<option value="Georgia">Georgian</option>
	<option value="Tahoma">Tahoma</option>
	<option value="Times New Roman">Times New Roman</option>
	<option value="Verdana" >Verdana</option>
	<option value="impact" >Impact</option>
</select>



</td>
<td>
 <select id="bn_ak_fontsize" class="optfont" onchange="ak_fontsize();">


 <%
 for xf=1 to 22

 response.write "<option value="&xf&">"&xf&"</option>"
 next

 %>

 </select>
</td>
<td>
<% if  bmodificar then
  response.write "<input VALUE='ACTUALIZAR' CLASS=BT type=button onclick='asignar();document.MSForm.submit();' >"
  end if
%>
  <form  id='MSForm' name='MSForm' style="display:none;" method=post
   action="<% IF SESSION("FILETIPO")="R" THEN RESPONSE.WRITE session("DirUpload") & "/" %>asp/EDNWEDTDET2.asp?Codobjeto=<%=nCodigo%>&Tipobjeto=<%=sTipo%>&concod=<%=concod%>&OPC=S" >
    <textarea style="display:none;" id="txt1" name="txt1" cols=100 rows=10></textarea>
  </form>
</td>
</tr>
</table>

<div id="foo" onkeyuph="chagecolor();" name="foo" class="nzzz" contenteditable="true" align="left" style="z-index:99999;height:400; width:700;background-color:white;padding:3; border-style:SOLID;border-width:1;border-color:#aaaaaa;overflow:auto;">
   <%=XNNWFILE%>
  </div>
<br style="font-size:5pt;" >
</div>
<script languaje="javascript">
  //ak_fontsize();
  //ak_fontname();
</script>
<%


 ''' sHTML = sHTML & "<div style='border-style:solid;border-width:2;' >" & xxxx & "</div>"



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

 ' Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
 ' XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
 ' XferTarget0 = "\\server-a\FILESERVER\CarpetasCMD\"

'  XNABRIR = XferTarget0 & XNNWCOD & ".CSE"

'  set fff = server.createobject("scripting.filesystemobject")

'    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
'    xxxfl = ""

'    IF OPC = "S" THEN

   ''''    XNPSRFL = REQUEST.FORM("txt1")
'       XNPSRFL = ""

'       For I = 1 To Request.Form("txt1").Count
'           XNPSRFL = XNPSRFL & Request.Form("txt1")(I)
'       Next
'

'       Set MyFile = fff.CreateTextFile(xarchivo)
'       MyFile.Write XNPSRFL
'       MyFile.Close

'       xxxfl = XNPSRFL
'
'    ELSE

   ' if fff.FileExists(trim(Xarchivo)) then
    '   Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
   '    do while fff1.AtEndOfstream <> true
   '       xxxfl = xxxfl & fff1.readline
    '   loop
    '   fff1.close

       xxxfl= getReadfile(XNNWCOD & ".CSE","")
   ' end if
     'RESPONSE.WRITE xxxfl

'    END IF

    NWEDTVER = xxxfl

end function

%>