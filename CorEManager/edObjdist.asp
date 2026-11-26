<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sestado = 1
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")


  aplcheck()
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  bModificaresp = ofv.AModificar
'  RESPONSE.WRITE Session("Form2") & "  " & session("Nomodificar")
  if session("Nomodificar") = "S" then

   'if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False
      bModificar = False
  ELSE
   if session("Nomodificar") = "N" AND LEFT(Session("Form2"),4) ="TEMA"   THEN bModificar = False
  end if


  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sVolver = request.querystring("volver")
  sCod    = request.querystring("cod")
  vuelta  = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=ddd>"

'-------------------------------------------------

  set cn  = ofv.conectar(ofv.strconn2)

  Strsql = "SELECT * FROM OBJdistPrty  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' "
'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  rs.eof then


     Strsql = "SELECT * FROM OBJdistPrty  "
	 Strsql = Strsql & " Where codobjeto=0"
	 Strsql = Strsql & " And tipobjeto = 'DEFDST' "

      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if rs.eof then
      Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
      Strsql = Strsql & nCodigo & ",'" & stipo & "','COREMANAGER','Activacion de <ESQNOR>',"
       if  session("APLMDCLI")="LACAJA" then


         Strsql = Strsql & "'Se publica <ESQNOR>',1,'NO','7')"

       elseif session("APLMDCLI")="CFA"  THEN
       Strsql = Strsql & "'Se publica <ESQNOR>',1,'NO','7')"
       ELSE
       Strsql = Strsql & "'<ESQNOR> esta vigente a partir de la fecha',1,'NO','7')"
      END IF

      else
      Strsql = "insert into OBJdistPrty  (codobjeto,tipobjeto,emisor,asunto,texto,importancia,aviso,dias) values ("
	   Strsql = Strsql & nCodigo & ",'" & stipo & "','COREMANAGER','"&rs("asunto")& "',"

       Strsql = Strsql & "'"&rs("texto")&"',"&rs("importancia")&",'"&rs("aviso")&"','"&rs("dias")&"')"
       'response.write strsql
      end if

      call ofv.crearconsultaEx(StrSql,cn,1,parametros)


      Strsql = "SELECT * FROM OBJdistPrty  "
      Strsql = Strsql & " Where codobjeto=" & nCodigo
      Strsql = Strsql & " And tipobjeto = '" & stipo & "' "
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if

  set cnx = ofv.conectar(ofv.strconn0)

  Celda = ofv.gencelda("","ttt colspan=2","","","","","","Opciones&nbsp;de&nbsp;distribuci&oacute;n","si")
  Fila  = ofv.GenRow("","","","","","","",Celda,"si")

  sxopc3 = ""

  strsql = "Select descripcion From Usuarios  where id = '" & rs(3) & "' "
  set rsx = ofv.crearconsultaEx(StrSql,cnx,1,volver)
  if  not  rsx.eof then
      sxopc3 = rsx(0)
  end if
  call ofv.cerrarconsulta(rsx)
  call ofv.cerrarconn(cnx)

  sImagenL = "<FONT FACE=corflaticon COLOR=#AA0000 SIZE=2 >&#0157;</FONT>"
  sImagenG = "<FONT FACE=CorFlatIcon COLOR=#005500 SIZE=3 >&#0253;</FONT>"

  sForm  = "distx"

  if bModificar  then
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJdistPrty&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

       '' if SESSION("MAILOPC") = "HTML" then
       ''      sInput = ofv.GenerarInput("","Actualizar","submit","","10","bt on","")
       ''       formu54 =  ofv.GenForm(sForm,"",sAccion,"","post onsubmit="" return adddato(this);""","",sInput,"si")
       ''      Celda = ofv.GenCelda("","aaa colspan=4","","10%","","","",formu54,"si")
       ''      Fila12 = ofv.genrow("","","","","","","",Celda,"si")
       ''    end if
       '' else
       '' sAccion = ""

  end if


  scomK  = "1" & chr(9) & "2" & chr(9) & "3" & chr(9) & "4" & chr(9) & "5" & chr(9) & "6" & chr(9) & "7" & chr(9) & "8" & chr(9) & "9" & chr(9) & "10" & chr(9)
  scomK  = scomK & "11" & chr(9) & "12" & chr(9) & "13" & chr(9) & "14" & chr(9) & "15" & chr(9) & "16" & chr(9) & "17" & chr(9) & "18" & chr(9) & "19" & chr(9)
  scomK  = scomK & "20" & chr(9) & "21" & chr(9) & "22" & chr(9) & "23" & chr(9) & "24" & chr(9) & "25" & chr(9) & "26" & chr(9) & "27" & chr(9) & "28" & chr(9)
  scomK  = scomK & "29" & chr(9) & "30" & chr(9) & "31"
  scomD  = scomK
  sxopc2 = ofv.gencomboFX(sForm ,rs(8).name,rs(8),scomK,scomD," ")

  scomK  = "0" & chr(9) & "1" & chr(9) & "2"
  scomD  = "BAJA" & chr(9) & "NORMAL" & chr(9) & "ALTA"
  sxopc4 = ofv.gencomboFX(sForm ,rs(6).name,rs(6),scomK,scomD," ")

  scomK  = "NO" & chr(9) & "SI"
  scomD  = scomK
  sxopc5 = ofv.gencomboFX(sForm ,rs(7).name,rs(7),scomK,scomD," ")


  sxopcx  = "&nbsp;GLOBAL"
  sxopcxi = sImagenG

  Strsql = "SELECT * FROM OBJdist  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if  not rsx.eof then
      sxopcx  = "&nbsp;LISTA DE DISTRIBUCION"
      sxopcxi = sImagenL
  end if
  call ofv.cerrarconsulta(rsx)

    celdas = ""

    celdas = celdas & ofv.GenCelda("","tt","","20%","","","","Emisor","si")
    sInput = sxopc3
    celdas = celdas & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    fila   = fila & ofv.GenRow("","","","","20","","",celdas,"si")

    celdas = ""
    celdas = celdas & ofv.GenCelda("","tt","","","","","","Importancia","si")
    sInput = sxopc4
    celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""
    celdas = celdas & ofv.GenCelda("","tt","","","","","","Vigencia (dias)","si")
    sInput = sxopc2
    celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""
    celdas = celdas & ofv.GenCelda("","tt","","","","","","Aviso","si")
    sInput = sxopc5
    celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""
    celdas = celdas & ofv.GenCelda("","tt","","","","","","Destino","si")
    sInput = sxopcxi & "<a class=ttt href='edObjDistDet.asp?" & request.querystring & "'  title='Actualizar lista de distribucion.' >" & sxopcx & "</a>"
    celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""
    celdas = ofv.GenCelda("","tt","","","","","","Asunto","si")
    sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"90","dt","")

    celdas = celdas & ofv.GenCelda("","uti","","","","","",sInput,"si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ttt colspan=2 ","","","","","","Mensaje:","si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")

    sInput = " onkeypress=" & chr(34)
    sInput = sInput & " if (this.value.length > 1950) {alert('El texto es demasiado"
    sInput = sInput & " extenso'); return false;}" & chr(34)

'response.write  rs(5).name
     if SESSION("MAILOPC") = "HTML" then
      sInput = "<textarea style='display:none;' name=" & rs(5).name  & "  id=" & rs(5).name   & sInput &  " cols=80 rows=4>" & rs(5)  & "</textarea>"
      seditor=AGREGAREDITOR(REPLACE(rs(5) & "","''","'"))
    else
    seditor=""
    sInput = sInput & " onblur=" & chr(34) & "document." & sform   & ".submit();" & chr(34)
    sInput = "<textarea  name=" & rs(5).name  & " id=" & rs(5).name   &  sInput &  " cols=80 rows=4>" & rs(5)  & "</textarea>"
    end if

   'sInput =   ofv.GenForm(sForm & "xsx2","",sAccion,"","post","",sInput,"si'")

    celdas = ofv.GenCelda("","ut colspan=2","center","","","","",sInput & seditor,"si")
    fila   = fila & ofv.GenRow("","","","","10","","",celdas,"si")



   if bModificar  then
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJdistPrty&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      '  response.write "222"

   else
      sAccion =""
    end if


         if SESSION("MAILOPC") = "HTML" then
             boton = ofv.GenerarInput("","Actualizar","submit","","10","bt on","")
      	     Celdab = ofv.GenCelda("","aaa colspan=4","","10%","","","",boton,"si")
      	     Filab = fila & ofv.genrow("","","","","","","",Celdab,"si")
          end if
           ntabla= ofv.GenTabla("","aaa","","80%","","0","0","1","0","","", filab ,"si")
           sformt =  ofv.GenForm(sForm,"",sAccion,"","post onsubmit="" return adddato(this);""","",ntabla,"si")


  sHTML = sHTML & "<br><center>"
  sHTML = sHTML  & sformt








  Celda = ofv.gencelda("","ttt colspan=4","","","","","","Lista&nbsp;de&nbsp;distribuci&oacute;n","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","1%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","ttr colspan=3","","98%","","","","Destinatario","si")

  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  Strsql = "SELECT * FROM OBJdist  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

  do until rs.eof

    sForm = "dist" & rs(0)

   if bModificar  then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objdist&volver="
      sHref = sHref  & vuelta & "?Cod=" & sCod
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Usuario",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'destinatario

    sTabla = "Usuario"

    sBusca = rs(4)

    if Ucase(sBusca) = "SECT-" then

       sGrpUsr = "S/D"


       Strsql = "SELECT * FROM eco"
       Strsql = Strsql & " Where eco = 0" & rs(5)

       set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#993333><b>Sector: </b></font>" & rx(1)
       call ofv.cerrarconsulta(rx)


       sinput = ofv.gencelda("","VT  colspan=3 ","","100%","","","",sGrpUsr,"si")


    elseif Ucase(sBusca) = "TODOS" then

       sGrpUsr = "S/D"


       Strsql = "SELECT * FROM CLIENTEPLAN "
       Strsql = Strsql & " Where ID = 0" & rs(5)

       set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#336699><b>Empresa: </b></font>" & rx(1)
       call ofv.cerrarconsulta(rx)


       sinput = ofv.gencelda("","VT  ","","100%","","","",sGrpUsr,"si")


    else

       sGrpUsr = "S/D"
       set cx = ofv.conectar(ofv.strconn0)

       Strsql = "SELECT * FROM " & sTabla & "s"
       Strsql = Strsql & " Where ID = '" & sBusca & "' "

       set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#005500><b>" & sTabla & ": </b></font>" & rx(0) & " - " & rx(1)
       call ofv.cerrarconsulta(rx)
       call ofv.cerrarconn(cx)

       sinput = ofv.gencelda("","VT  colspan=3 ","","100%","","","",sGrpUsr,"si")

    end if

    Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

   rs.movenext
   loop
    else


       Celda  = ofv.gencelda("","vt  colspan=4  ","","100%","","","","&nbsp;","si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    end if

    Celda = ofv.gencelda("","vt","","2%","","","","&nbsp;","si")
    celda = celda & ofv.gencelda("","vt","","80%","","","","&nbsp;","si")
    celda = celda & ofv.gencelda("","vt","","4%","","","","&nbsp;","si")
    celda = celda & ofv.gencelda("","vt","","4%","","","","&nbsp;","si")

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

'    sHTML = sHTML & ofv.GenTabla("","","","98%","","0","0","-2","-1","","",fila,"si")

    Fila = ""

   if bModificar  then

       celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")
       AgregarA = "Agregardist.asp?Codobjeto=" & nCodigo & "&ta=U"
       AgregarA = AgregarA & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
       Agregar = ofv.GenerarInput("bac1","Usuario","submit","","10","bt","")
       Agregar = ofv.GenCelda("","","right","1%","","","",Agregar,"si")
       Agregar = ofv.GenForm("Agregar","",AgregarA,"","post","",Agregar,"si")

       AgregarA = "Agregardist.asp?Codobjeto=" & nCodigo & "&ta=S"
       AgregarA = AgregarA & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
       Agregar2 = ofv.GenerarInput("bac1","Sector","submit","","10","bt","")
       Agregar2 = ofv.GenCelda("","","right","1%","","","",Agregar2,"si")
       Agregar2 = ofv.GenForm("Agregar2","",AgregarA,"","post","",Agregar2,"si")

    IF SESSION("EMPMUN") = "S" then

       AgregarA =  "Agregardist.asp?Codobjeto=" & nCodigo & "&ta=E"
       AgregarA = AgregarA & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
       Agregar3 = ofv.GenerarInput("bac1","Empresa","submit","","10","bt","")
       Agregar3 = ofv.GenCelda("","","right","1%","","","",Agregar3,"si")
       Agregar3 = ofv.GenForm("Agregar2","",AgregarA,"","post","",Agregar3,"si")
    else
       Agregar3 = ofv.GenCelda("","","right","1%","","","","&nbsp;","si")
    end if

       Fila = fila & ofv.GenRow("","","","","","","",celda  & Agregar3 & Agregar  & Agregar2,"si")


   end if



  sHTML = sHTML & "</center>"


'  sHTML = sHTML & ofv.GenTabla("","","","10","","0","0","1","0","","",fila,"si")


  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)



  Agregar  = ""
  Retornar = ""



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")
response.write  sHTML

%>
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
 function setTextColor(picker) {
		highlight( picker.toString());


	}


	function setTextColorf(picker) {
			ak_forecolor( picker.toString());


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

		function ak_bold(){
			var texto = document.getElementById("foo");
			if (document.selection)
			{
				var str=document.selection.createRange().text;
				seleccion = document.selection.createRange();
				seleccion.pasteHTML( "\<B\>" + str + "\<\/B\>");
				texto.focus();
			}
	}

function replaceAll( text, busca, reemplaza ){
	  while (text.toString().indexOf(busca) != -1)
	      text = text.toString().replace(busca,reemplaza);
	  return text;
	}

	function ak_boldw()
		{

			//document.getElementById(editor).focus();

			  document.execCommand('bold');
			//document.getElementById(editor).focus();

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


	function ak_ocultar()
		{
		var t= document.getElementById("texto");
		//style={display:none;}
   // alert(t);
		if (t.style.display=='')
		{
		t.style.display='none';
		}
		else
		{
		t.style.display='';
		}
	}


	function ak_inserttags(str1, str2) {
	          //  alert(str2);
			  	document.getElementById('foo').focus();
			  	var tr = document.selection.createRange();
				if (str2 != ""){tr.pasteHTML(str1 + tr.text + str2);}
				else {tr.pasteHTML(str1);}

			  	tr.select();
			  	document.getElementById('foo').focus();

		}




function ak_createlinkDoc()
		{


		ak_inserttags('##ESQNOR','##');

		}
	function ak_createlinka()
		{
		document.getElementById(editor).focus();

		var surl=prompt("ingrese codigo doc")

		var n=surl.split("-");

		if (n.length==3)
		{
		var a = document.createElement('a');
	    a.innerHTML ="<u>Ir a documento..</u>";
	    var sActcon = 'Consultar.asp?DOCMODAL=CL&Codigo='+ n[1] +'&Tipo=' + n[2];
            sActcon = sActcon +  '&Datos=*smenu=zzzzz*lista=3';
	    surl= '"' + sActcon + '"';
		//	a.textContent = 'click me';
			a.href = '#';
			a.setAttribute('className', 'istt');
			//alert('AbrirCons(\'' + surl + '\',"")');
		    a.setAttribute('onclick', 'AbrirCons(' + surl + ',"'+ n[2] + '")');


        document.getElementById(editor).appendChild(a);

		}
		}

		function ak_createlink()
				{
				document.getElementById(editor).focus();

				var surl=prompt("ingrese url")

				//var n=surl.split("-");

			//	if (n.length==0)
			//	{
				var a = document.createElement('a');
			    a.innerHTML ="<u>Link Externo..</u>";

					a.href = "'" + surl + "' ";
					a.setAttribute('className', 'istt');
					a.setAttribute('target', '_new');
			;


		        document.getElementById(editor).appendChild(a);

		}



function adddato(f)
	{
	var d= document.getElementById("foo").innerHTML ;



   // alert(d);
    var tt=d;

    tt=tt.replace(/<P[^>]*>/gi, '');
    tt=tt.replace(/<\/P>/gi, '<br/>');
    tt=tt.replace(/&nbsp;/gi,' ');
    var d= document.getElementById("foo")
   // f.texto.value=tt;
    tt=tt.replace(/##ESQNOR##/gi,'<ESQNOR>')
    var ft= document.getElementById("texto");
    ft.value=replaceAll(tt,"'","");

	if (ft.value.length > 3000)
	{
	alert('El texto es demasiado extenso');
	return false;
	}

return true;
	}

	function ak_unlink()
	{
	myEditor.focus();
	myEditor.document.execCommand('Unlink', false, null);
	}
</script>
<%
  sHTML =  "</body></html>"
end if

response.write sHTML


FUNCTION AGREGAREDITOR (texto)
dim scedit

scedit="<div unselectable='on' style='width:720; background-color:#f9f9f9;border-style:SOLID;border-radius:5px;border-width:2;border-color:'" & session("CMFMNBOT") & ";>"

scedit=scedit & "<table style=""margin-top:0px;margin-bottom:0px;"" cellpadding=0 cellspacing=1 width=92%  >"
scedit=scedit & "<tr>"
'scedit=scedit & "<td CLASS=TTT width=50% >&nbsp;</td>"
scedit=scedit & "<td>  <img  src=""imgeditor/mn_createlink.gif"" id=E1 onclick='ak_createlinkDoc(this);'"
scedit=scedit &   " title=""crear link a Documento"">"
scedit=scedit &"  </img></td>"



scedit=scedit & "<td>  <img class=""btnDisabled"" src=""imgeditor/mn_insertorderedlist.gif"" id=E1 onclick='ak_insertorderedlist(this);'"
scedit=scedit &   " title=""Ordered List"">"
scedit=scedit &"  </img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_insertunorderedlist.gif"" onclick='ak_insertunorderedlist();'"
scedit=scedit & "   title=""Unordered List"">"
scedit=scedit &  "</img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_outindent.gif""  onclick='ak_outdent();'"
scedit=scedit & "     title=""Outdent"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img src=""imgeditor/mn_indent.gif"" onclick='ak_indent();'"
scedit=scedit & "    title=""Indent"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img  src=""imgeditor/mn_justifyright.gif"" onclick='ak_justifyright();'"
scedit=scedit & "   title=""Left Justify"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img  src=""imgeditor/mn_justifyleft.gif"" onclick='ak_justifyleft();'"
scedit=scedit & "   title=""Right Justify"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img  src=""imgeditor/mn_justifycenter.gif"" onclick='ak_justifycenter();'"
scedit=scedit & "   title=""Center Justify"">"
scedit=scedit & "  </img></td>"



scedit=scedit & "<td> </td>"


scedit=scedit & "<td> <img  src=""imgeditor/mn_inserthorizontalrule.gif"" onclick='ak_inserthorizontalrule();'"
scedit=scedit &"       title=""Insert Horizontal Rule"">"
scedit=scedit &"  </img></td>"

scedit=scedit & "<td> <button id=bgcolor style=""border: none;height:20px;width: 20px;background: url(imgeditor/mn_backcolor.gif)"" class=""jscolor {valueElement:'valueSpan',styleElement:'styleSpan',onFineChange:'setTextColor(this)'}""></button> </td>"


scedit=scedit & "<td>  <button id=bgcolor style=""border: none;height:20px;width: 20px;background: url(imgeditor/mn_forecolor.gif)"" class=""jscolor {valueElement:'valueSpan',styleElement:'styleSpan',onFineChange:'setTextColorf(this)'}""></button></tD>"



scedit=scedit & "<td> <img  src=""imgeditor/undo.gif"" onclick='ak_Undo();'"
scedit=scedit & "          title=""Undo"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>  <img  src=""imgeditor/redo.gif"" onclick='ak_Redo();'"
scedit=scedit & "           title=""Redo"">"
scedit=scedit & "  </img></td>"



'scedit=scedit & "</tr>"
'scedit=scedit & "</table>"

'scedit=scedit & "<table style=""margin-top:0px;margin-bottom:0px;"" cellpadding=0 cellspacing=1 width=92% >"
'scedit=scedit & "<tr>"
'scedit=scedit & "<td width=50% >&nbsp;</td>"

'scedit=scedit & "<td><img  src=""imgeditor/mn_bold.gif"" onclick='ak_bold();' title=""Bold""> </img></td>"

scedit=scedit & "<td><img  src=""imgeditor/mn_italic.gif"" onclick='ak_italic();'"
scedit=scedit & "    title=""Italic"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td><img  src=""imgeditor/mn_underline.gif""  onclick='ak_underline();'"
scedit=scedit & "   title=""Underline"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td><img   src=""imgeditor/mn_strikethrough.gif""  onclick='ak_strikethrough() ;'"
scedit=scedit & "    title=""Strike-through"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_superscript.gif""  onclick='ak_superscript();'"
scedit=scedit & "    title=""Superscript"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_subscript.gif""  onclick='ak_subscript();'"
scedit=scedit & "    title=""Subscript"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>"
scedit=scedit & "<select id=""bn_ak_fontname"" class=""optfont"" onchange=""ak_fontname();"">"
scedit=scedit & "	<option value=""Arial"" selected>Arial</option>"
scedit=scedit & "	<option value=""Courier New"">Courier New</option>"
scedit=scedit & "	<option value=""Georgia"">Georgian</option>"
scedit=scedit & "	<option value=""Tahoma"">Tahoma</option>"
scedit=scedit & "	<option value=""Times New Roman"">Times New Roman</option>"
scedit=scedit & "	<option value=""Verdana"" >Verdana</option>"
scedit=scedit & "	<option value=""impact"" >Impact</option>"
scedit=scedit & "</select>"
scedit=scedit & "</td>"
scedit=scedit & "<td>"
scedit=scedit & " <select id=""bn_ak_fontsize"" class=""optfont"" onchange=""ak_fontsize();"">"



 for xf=1 to 22

 scedit=scedit & "<option value="&xf&">"&xf&"</option>"
 next



 scedit=scedit & "</select>"
scedit=scedit & "</td>"
scedit=scedit & "<td >"
 scedit=scedit & "  <img src=""imgeditor/mn_info.gif""   onclick='ak_ocultar();'"
 scedit=scedit & "    title=""Subscript"">"
 scedit=scedit & "  </img>"


'
scedit=scedit & "</td>"
scedit=scedit & "</tr>"
scedit=scedit & "</table>"

scedit=scedit & "<div id=foo name=foo  class=nzzz contenteditable align=left style=""height:150; width:700;background-color:white;"
scedit=scedit & " padding:3; border-style:SOLID;border-width:1;border-color:#aaaaaa;overflow=auto;"">"
texto=replace(texto ,"<ESQNOR>","##ESQNOR##")
  scedit=scedit & texto
 scedit=scedit & " </div>"
scedit=scedit & "<br style=""font-size:5pt;"" >"
scedit=scedit & "</div><span id=valueSpan style='display:none'><span id=styleSpan style='display:none' >Color preview</span>"

AGREGAREDITOR=scedit
END FUNCTION
%>