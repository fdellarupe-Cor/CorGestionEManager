<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!DOCTYPE>
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

Response.ContentType = "text/html"
Response.AddHeader "Content-Type", "text/html;charset=UTF-8"
Response.CodePage = 65001
Response.CharSet = "UTF-8"

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
     Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  sHTML = ofv.FormHeader(session("FormN"))

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

'----------------------------------------------

 shtml = shtml & CorPnlFnc("N","")


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

%>
<!-- #INCLUDE FILE="IncFile/CorVinculadorPL.inc" -->
<%


  IF WKSIGUE = "S" THEN

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  if sTipo = "CON" or sTipo = "ESQ" Then
    sCalificador = "R"
    sSector = "Sectores responsables"
  else
    sCalificador = "O"
    sSector = "Sectores responsables"
  end if

IF session("EMPMUN") = "S" then

  strsql = "SELECT a.*, c.descri,b.cliente FROM  OBJSECTOR a "
  strsql = strsql & " inner join ECO c on (a.codsector=c.eco) "
  strsql = strsql & " inner join clienteplan b on (b.ID=c.ecoUN) "
  strsql = strsql & " where a.codsector = c.eco "
  strsql = strsql & " and c.estado=3 and a.codobjeto=" & nCodigo
  strsql = strsql & " And (calificador = 'O' or calificador = 'C' or calificador = 'O') "
  strsql = strsql & " order by b.cliente"

else

  Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
  Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and codobjeto=" & nCodigo
  Strsql = Strsql & " And (calificador = '" & sCalificador & "' or calificador = 'C' or calificador = 'O') "
  Strsql = Strsql & " order by eco.descri"

end if

 '	strsql = "select * from eco"
sIcon = "<font style='font-size:13pt;'><span class='la la-bank'></span></font>"
	'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sCargados = ""
  sPosX = 1

  do until rs.eof

	if sPosX <= 5 then
  	    sCargados = sCargados & "<li class='c' data-id='" & rs("codsector") & "' data-pos='" & sPosX & "'>" & sIcon & "&nbsp;" & rs("descri") & "</li>"
  	else
  		sCargados = sCargados & "<li class='c' data-id='" & rs("codsector") & "' data-pos='" & sPosX & "' style='display:none;'>" & sIcon & "&nbsp;" & rs("descri") & "</li>"
  	end if

	  sPosX = sPosX + 1

  rs.movenext
  loop

end if
'''''''''''''''''

		If session("UN") <> 1 Then

			'sql="select  eco,descri from eco where estado=3 and ecoun=0" & session("UN") & " order by descri"

			sql="select b.eco,b.descri from eco b "
			sql = sql & "where b.estado=3 and b.ecoun=0" & session("UN") & " and b.eco not in "
			sql = sql & "(select codsector from objsector where codobjeto=" & session("NCliente") & " and tipobjeto='" & session("TCliente") & "') "
			sql = sql & "order by b.descri"

		Else

			XEMPRE = Split(session("EmpHbl"),Chr(9))

			DELIM=""
			idempre = ""

			For i=0 To UBound(XEMPRE)

				idempre = idempre & DELIM & XEMPRE(i)
				DELIM = ","

			Next

			If idempre = "" Then idempre = session("CLIEMP")

			'sql="select  eco,descri from eco where estado=3 and ecoun in (" & idempre & ") order by descri"

			sql="select b.eco,b.descri from eco b "
			sql = sql & "where b.estado=3 and b.eco not in "
			sql = sql & "(select codsector from objsector where codobjeto=" & nCodigo & " and tipobjeto='" & sTipo & "') "
			sql = sql & "order by b.descri"

		End If
		'response.write sql
		set rx = ofv.crearconsultaEx(sql,cn,1,parametros)

		  sTodos = ""
		  sPos = 1

		  do until rx.eof

			if sPos <= 15 then
		  		sTodos = sTodos & "<li class='l' data-id='" & rx("eco") & "' data-pos='" & sPos & "'>" & sIcon & "&nbsp;" & rx("descri") & "</li>"
		  	else
		  		sTodos = sTodos & "<li class='l' data-id='" & rx("eco") & "' data-pos='" & sPos & "' style='display:none;'>" & sIcon & "&nbsp;" & rx("descri") & "</li>"
		  	end if

			sPos = sPos + 1

		  rx.movenext
 		 loop

	  'DIV LISTADO
	  sHTML = sHTML &  "<div id='listadosCont'>"
	  sHTML = sHTML &  "<div id='listadosTit'>"
	  sHTML = sHTML &  "<table id='tabPag'>"
	  sHTML = sHTML &  "<tr>"
	  sHTML = sHTML &  "<td id='td1'><span class='la la-chevron-circle-left' style='font-size: 25pt;'></span></td>"
	  sHTML = sHTML &  "<td id='td2'>TODOS</td>"
	  sHTML = sHTML &  "<td id='td3'><span class='la la-chevron-circle-right' style='font-size: 25pt;'></span></td>"
	  sHTML = sHTML &  "<td id='td4'><span class='la la-search' style='font-size: 25pt;'></span></td>"
	  sHTML = sHTML &  "</tr>"
	  sHTML = sHTML &  "</table>"
	  sHTML = sHTML &  "</div>"
	  sHTML = sHTML &  "<div id='busCont' style='display:none;'><input id='buscador' name='buscador' class='busc' type='text' ></div>"
	  sHTML = sHTML &  "<ul id='listados' data-posAct='15' data-max='" & (sPos - 1) & "' style='display:none;'>" & sTodos & "</ul>"
	  sHTML = sHTML &  "</div>"

	  'DIV CARGADOS
	  sHTML = sHTML &  "<div id='cargadosCont'>"
	  sHTML = sHTML &  "<div id='cargadosTit'>"
	  sHTML = sHTML &  "<table id='tabPag1'>"
	  sHTML = sHTML &  "<tr>"
	  sHTML = sHTML &  "<td id='td5'><span class='la la-chevron-circle-left' style='font-size: 25pt;'></span></td>"
	  sHTML = sHTML &  "<td id='td6'>CARGADOS</td>"
	  sHTML = sHTML &  "<td id='td7'><span class='la la-chevron-circle-right' style='font-size: 25pt;'></span></td>"
	  sHTML = sHTML &  "</tr>"
	  sHTML = sHTML &  "</table>"
	  sHTML = sHTML &  "</div>"
	  sHTML = sHTML &  "<ul id='cargados' data-posAct='5' data-max='" & (sPosX - 1) & "'>" & sCargados & "</ul>"
	  sHTML = sHTML &  "</div>"

	  'DIV ELIMINAR
  	  sHTML = sHTML &  "<div id='eliminar'><span id='tacho' class='la la-trash-o' style='color:#000000;font-size: 190pt;'></span></div>"

  'Agregar
  'if bModificar then
  '  sAccion = "AgregarSectorEN.asp?Tabla=Objsector&volver=" & vuelta & "&Codobjeto="
  '  sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Calificador='" & sCalificador & "'"
  '  Agregar = ofv.BotonAgregar(sAccion)
  'end if

  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if




'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","") & GTCOBS()

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

shtml = shtml & "<link rel='stylesheet' href='../Coress20/IncFile/css/estilos.css'>"
shtml = shtml & "<script type='text/javascript' src='../Coress20/IncFile/js/jquery.min.js'></script>"
shtml = shtml & "<script type='text/javascript' src='../Coress20/IncFile/js/jquery-ui.js'></script>"

shtml = shtml & "<script>"

shtml = shtml & "$( function() {"

'CSS
shtml = shtml & "$('#cargadosCont').css({"
shtml = shtml & "'border': 'solid 3px #000000',"
shtml = shtml & "'border-radius': '10px',"
shtml = shtml & "'width': '400',"
shtml = shtml & "'height': '210',"
shtml = shtml & "'margin-left': '20',"
shtml = shtml & "'overflow': 'hidden',"
shtml = shtml & "'margin-top': '20',"
shtml = shtml & "'color': '#000000'"
shtml = shtml & "});"

shtml = shtml & "$('#cargadosTit, #listadosTit').css({"
shtml = shtml & "'width': '400',"
shtml = shtml & "'height': '50',"
shtml = shtml & "'font-size': '18pt',"
shtml = shtml & "'text-align': 'center'"
shtml = shtml & "});"

shtml = shtml & "$('#busCont').css({"
shtml = shtml & "'width': '400',"
shtml = shtml & "'height': '35'"
shtml = shtml & "});"

shtml = shtml & "$('#cargados, #listados').css({"
shtml = shtml & "'overflow': 'auto',"
shtml = shtml & "'width': '100%',"
shtml = shtml & "'height': '100%',"
shtml = shtml & "'padding-left': '25',"
shtml = shtml & "'font-size': '13',"
shtml = shtml & "});"

shtml = shtml & "$('#listadosCont').css({"
shtml = shtml & "'border': 'solid 3px #000000',"
shtml = shtml & "'border-radius': '10px',"
shtml = shtml & "'width': '400',"
shtml = shtml & "'height': '430',"
shtml = shtml & "'margin-right': '20',"
shtml = shtml & "'overflow': 'hidden',"
shtml = shtml & "'float': 'right',"
shtml = shtml & "'margin-top': '20',"
shtml = shtml & "'color': '#000000'"
shtml = shtml & "});"

shtml = shtml & "$('#eliminar').css({"
shtml = shtml & "'width': '400',"
shtml = shtml & "'height': '170',"
shtml = shtml & "'float': 'left',"
shtml = shtml & "'margin-left': '20',"
shtml = shtml & "'margin-top': '10',"
shtml = shtml & "'text-align': 'center'"
shtml = shtml & "});"

'encabezado de listados
shtml = shtml & "$('#tabPag').css({"
shtml = shtml & "'width': $('#listadosTit').width(),"
shtml = shtml & "'height': $('#listadosTit').height(),"
shtml = shtml & "'border-bottom': 'solid 3px black',"
shtml = shtml & "'text-align': 'center'"
shtml = shtml & "});"

'encabezado de cargados
shtml = shtml & "$('#tabPag1').css({"
shtml = shtml & "'width': $('#cargadosTit').width(),"
shtml = shtml & "'height': $('#cargadosTit').height(),"
shtml = shtml & "'border-bottom': 'solid 3px black',"
shtml = shtml & "'text-align': 'center'"
shtml = shtml & "});"

shtml = shtml & "$('#td1, #td2, #td3, #td4').css({"
shtml = shtml & "'width': $('#tabPag').width() /4,"
shtml = shtml & "'font-size': '18pt'"
shtml = shtml & "});"

shtml = shtml & "$('#td5, #td6, #td7').css({"
shtml = shtml & "'width': $('#tabPag1').width() /3,"
shtml = shtml & "'font-size': '18pt'"
shtml = shtml & "});"

'SORTABLE
shtml = shtml & "$('#listados').sortable({"
shtml = shtml & "connectWith: '#cargados',"
shtml = shtml & "start: function( event, ui ) {"
shtml = shtml & "ui.item.addClass('entra');"
shtml = shtml & "},"
shtml = shtml & "stop: function( event, ui ) {"
shtml = shtml & "ui.item.removeClass('entra');"
shtml = shtml & "}"
shtml = shtml & "});"

shtml = shtml & "$('#cargados').sortable({"
shtml = shtml & "update: function(ev,ui){ "
shtml = shtml & "ui.item.removeClass('l');"
shtml = shtml & "ui.item.addClass('c');"
shtml = shtml & "}"
shtml = shtml & "});"

'boton anterior todos
shtml = shtml & "$( '#td1' ).click(function() {"
shtml = shtml & "var minimo = 15;"
shtml = shtml & "var posActual = parseInt($('#listados').attr('data-posAct'));"
shtml = shtml & "if(posActual > minimo){"
shtml = shtml & "$('#listados').attr('data-posAct',posActual - 15);"
shtml = shtml & "var posActual = parseInt($('#listados').attr('data-posAct'));"
shtml = shtml & "}"
shtml = shtml & "$('#listados li').each(function (i) {"
shtml = shtml & "var liPos = parseInt($(this).attr('data-pos'));"
shtml = shtml & "$(this).hide();"
shtml = shtml & "if(liPos > posActual - 15 && liPos <= posActual){"
shtml = shtml & "$(this).show();"
shtml = shtml & "}"
shtml = shtml & "});"
shtml = shtml & "});"

'boton anterior cargados
shtml = shtml & "$( '#td5' ).click(function() {"
shtml = shtml & "var minimo = 5;"
shtml = shtml & "var posActual = parseInt($('#cargados').attr('data-posAct'));"
shtml = shtml & "if(posActual > minimo){"
shtml = shtml & "$('#cargados').attr('data-posAct',posActual - 5);"
shtml = shtml & "var posActual = parseInt($('#cargados').attr('data-posAct'));"
shtml = shtml & "}"
shtml = shtml & "$('#cargados li').each(function (i) {"
shtml = shtml & "var liPos = parseInt($(this).attr('data-pos'));"
shtml = shtml & "$(this).hide();"
shtml = shtml & "if(liPos > posActual - 5 && liPos <= posActual){"
shtml = shtml & "$(this).show();"
shtml = shtml & "}"
shtml = shtml & "});"
shtml = shtml & "});"

'mostrar todos los sectores para cargar
shtml = shtml & "$( '#td2' ).mouseover(function() {"
shtml = shtml & "$( '#listados' ).show();"
shtml = shtml & "});"

'boton siguiente todos
shtml = shtml & "$( '#td3' ).click(function() {"
shtml = shtml & "var cantMax = parseInt($('#listados').attr('data-max'));"
shtml = shtml & "var posActual = parseInt($('#listados').attr('data-posAct'));"
shtml = shtml & "var contador = 0;"
shtml = shtml & "if((posActual + contador) < cantMax){"
shtml = shtml & "$('#listados li').each(function (i) {"
shtml = shtml & "var liPos = parseInt($(this).attr('data-pos'));"
shtml = shtml & "$(this).hide();"
shtml = shtml & "if(liPos > posActual && liPos <= posActual + 15){"
shtml = shtml & "$(this).show();"
shtml = shtml & "contador++;"
shtml = shtml & "}"
shtml = shtml & "});"
shtml = shtml & "}"
shtml = shtml & "if((posActual + contador) < cantMax){"
shtml = shtml & "$('#listados').attr('data-posAct',posActual + contador);"
shtml = shtml & "}"
shtml = shtml & "});"

'boton siguiente cargados
shtml = shtml & "$( '#td7' ).click(function() {"
shtml = shtml & "var cantMax = parseInt($('#cargados').attr('data-max'));"
shtml = shtml & "var posActual = parseInt($('#cargados').attr('data-posAct'));"
shtml = shtml & "var contador = 0;"
shtml = shtml & "if((posActual + contador) < cantMax){"
shtml = shtml & "$('#cargados li').each(function (i) {"
shtml = shtml & "var liPos = parseInt($(this).attr('data-pos'));"
shtml = shtml & "$(this).hide();"
shtml = shtml & "if(liPos > posActual && liPos <= posActual + 5){"
shtml = shtml & "$(this).show();"
shtml = shtml & "contador++;"
shtml = shtml & "}"
shtml = shtml & "});"
shtml = shtml & "}"
shtml = shtml & "if((posActual + contador) < cantMax){"
shtml = shtml & "$('#cargados').attr('data-posAct',posActual + contador);"
shtml = shtml & "}"
shtml = shtml & "});"

'mostrar input del buscador
shtml = shtml & "$( '#td4' ).click(function() {"
shtml = shtml & "$( '#busCont' ).show();"
shtml = shtml & "$( '#buscador' ).focus();"
shtml = shtml & "});"

'buscador
shtml = shtml & "$('#buscador').keyup(function( e ) {"
shtml = shtml & "if(e.keyCode == 13) {"
shtml = shtml & "var url = 'sBuscador.asp?txt=' + $('#buscador').val();"
shtml = shtml & "$.get( url , function( data ) {"
shtml = shtml & "$('#listados').sortable( 'destroy' );"
shtml = shtml & "$('#listados').remove();"
shtml = shtml & "$('#listadosCont').append(data);"
shtml = shtml & "$('#listados').sortable({"
shtml = shtml & "connectWith: '#cargados'"
shtml = shtml & "});"
shtml = shtml & "$('#listados').css({"
shtml = shtml & "'overflow': 'auto',"
shtml = shtml & "'width': '100%',"
shtml = shtml & "'height': '100%',"
shtml = shtml & "'padding-left': '25',"
shtml = shtml & "'font-size': '13',"
shtml = shtml & "});"
shtml = shtml & "});"
shtml = shtml & "}"
shtml = shtml & "});"

'INSERT
shtml = shtml & "$( '#cargadosCont' ).droppable({"
shtml = shtml & "drop: function( event, ui ) {"
shtml = shtml & "if ($(ui.draggable).hasClass('l')) {"
shtml = shtml & "$.get('sectAcc.asp?acc=CARGA&calif=" & sCalificador & "&codsector=' + $(ui.draggable).attr('data-id'), function( data ) {"
shtml = shtml & "$(ui.draggable).attr('data-pos', $('#cargados').find('li').size());"
shtml = shtml & "$('#cargados').attr('data-max', $('#cargados').find('li').size());"
shtml = shtml & "if($(ui.draggable).attr('data-pos') > 5) {"
shtml = shtml & "$(ui.draggable).css('display','none')"
shtml = shtml & "}"
shtml = shtml & "});"
shtml = shtml & "}"
shtml = shtml & "}});"

'ELIMINAR
shtml = shtml & "$( '#eliminar' ).droppable({"
shtml = shtml & "drop: function(ev,ui){"
shtml = shtml & "if ($(ui.draggable).hasClass('c')) { "
shtml = shtml & "$.get('sectAcc.asp?acc=ELIM&codsector=' + $(ui.draggable).attr('data-id'), function( data ) {"
shtml = shtml & "});"
shtml = shtml & "$(ui.draggable).remove();"
shtml = shtml & "var url = 'sBuscador.asp';"
shtml = shtml & "$.get( url , function( data ) {"
shtml = shtml & "$('#listados').sortable( 'destroy' );"
shtml = shtml & "$('#listados').remove();"
shtml = shtml & "$('#listadosCont').append(data);"
shtml = shtml & "$('#listados').sortable({"
shtml = shtml & "connectWith: '#cargados'"
shtml = shtml & "});"
shtml = shtml & "$('#listados').css({"
shtml = shtml & "'overflow': 'auto',"
shtml = shtml & "'width': '100%',"
shtml = shtml & "'height': '100%',"
shtml = shtml & "'padding-left': '25',"
shtml = shtml & "'font-size': '13',"
shtml = shtml & "});"
shtml = shtml & "});"
shtml = shtml & "}"
shtml = shtml & "alert($('#cargados').find('li').size() - 1);"
shtml = shtml & "}"
shtml = shtml & "});"

shtml = shtml & "$('.l, .c').mouseover(function() {"
shtml = shtml & "$(this).css({"
shtml = shtml & "'font-size': '14pt'"
shtml = shtml & "});"
shtml = shtml & "});"

shtml = shtml & "$('.l, .c').mouseleave(function() {"
shtml = shtml & "$(this).css({"
shtml = shtml & "'font-size': '10pt'"
shtml = shtml & "});"
shtml = shtml & "});"

shtml = shtml & "});"

shtml = shtml & "</script>"

sHTML = sHTML &  "</body></html>"
end if



response.write sHTML

%>
