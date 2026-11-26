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
  sTipo = request.querystring("Tipobjeto")
  sTabla = request.querystring("Tabla")
  nID = request.querystring("ID")
  volver = request.querystring("volver")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------


  sHTML = sHTML & "<center>"

  sVuelve = ofv.convertircar(volver,"_","&")
  sVuelve = ofv.convertircar(sVuelve,"*","&")
  sVuelta = ofv.convertircar(sVuelve,"&","*")
  sVuelve = mid(sVuelve,4,len(sVuelve) - 3)

  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:8pt;background-color:#f7eedd;} "

  set cn = ofv.conectar(ofv.strconn2)

  if sTabla = "Objesquema2" then
     sdato = "Refesq2,rellink"
  else
     sdato = "Refesq"
  end if

  strsql = "SELECT " & sdato & " From " & sTabla & " Where objid = " & nID
''response.write strsql & "<br>"
''response.write nCodigo & " codobjeto " & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    sAccion = ""
    if bModificar then
      'Encabezado Actualizar
      sAccion = "asp/ActualizarEN.asp?Tabla=" & sTabla & "&volver=" & volver
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID="
      sAccion = sAccion & nID & "&Referencia=SI"
    end if
    sHTML = sHTML & ofv.GenForm("Actualizar","",sAccion,"","post","","","no")

    sHTML = sHTML & "<br>" & ofv.GenTabla("","","","60%","","0","0","0","0","","","","no")

''' Agregado para version 4.6 22012014

      set cn2 = ofv.conectar(ofv.strconn2)
	  scontdesc = "SD"
	  Strsql = "select * from csesquemasel where objid = '" &  nID & "' "
	  Strsql = Strsql & " and Codobjeto='" &  nCodigo & "' "
	''response.write strsql & "<br>"
	  set rsz = ofv.crearconsultaEx(StrSql,cn,1,volver)
	  if not rsz.eof then
	     scontdesc = rsz(1)
	  end if
	  call ofv.cerrarconsulta(rsz)

''' Agregado para version 4.6 22012014 FIN

    'Encabezado Referencia
    Celda = ofv.GenCelda("","ttT","","100%","","","","Texto de referencia - " & scontdesc,"si")
    sHTML = sHTML & ofv.GenRow("","","","","25","","",Celda,"si")

  if sTabla = "Objesquema2" then

      wasotxt = "N"
      stylob  = ""


         if not isnull(rs(1))  then
            if  clng(rs(1)) > 0 then

                Strsql = "SELECT * From OBJLNK where objid = 0" & rs(1)
                set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsx.eof then
                   wasotxt  = rsx(14) & ""
                   IF wasotxt = "" THEN wasotxt = "N"
                end if
                call ofv.cerrarconsulta(rsx)

            end if
          end if

       IF wasotxt = "S" THEN

          stylob = "UT  style='border-style:solid;border-width:1;border-color:#990000;' "

       end if


    'Texto de Referencia
    celda = "<font class=at><textarea class=dt name=" & rs(0).name & " rows=20 cols=100 "
    if bModificar and wasotxt = "N" then
        celda = celda & "onblur=" & chr(34)
        celda = celda & " if (this.value.length == 0) {alert('El texto es demasiado"
        celda = celda & " corto'); this.focus(); this.value=this.defaultValue;this.select(); return false;}" & chr(34)
        celda = celda & " onkeypress=" & chr(34)
        celda = celda & " if (this.value.length > 2450) {alert('El texto es demasiado"
        celda = celda & " extenso'); return false;}" & chr(34)
      else
        celda = celda & " readonly"
      end if

    celda = celda & " >" & rs(0) & "</textarea></font>"

    Celda = ofv.GenCelda("",stylob,"center","","","","",celda,"si")

  else

    'Texto de Referencia
    celda = "<font class=at><textarea class=dt name=" & sdato & " rows=20 cols=100 "
      if bModificar then
       celda = celda & "onblur=" & chr(34)
        celda = celda & " if (this.value.length == 0) {alert('El texto es demasiado"
        celda = celda & " corto'); this.focus(); this.value=this.defaultValue;this.select(); return false;}" & chr(34)
        celda = celda & " onkeypress=" & chr(34)
        celda = celda & " if (this.value.length > 2450) {alert('El texto es demasiado"
        celda = celda & " extenso'); return false;}" & chr(34)
      else
        celda = celda & " readonly"
      end if

    celda = celda & " >" & rs(0) & "</textarea></font>"

    Celda = ofv.GenCelda("","tt","center","","","","",celda,"si")

  end if



    sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenFTabla()

    sHTML = sHTML & ofv.GenTabla("","","","60%","","0","0","0","0","","","","no")
    Celda = ""
    if bModificar then
      'Encabezado Actualizar
      sInput = ofv.GenerarInput("boton","Actualizar","submit","","20","bt","")
      Celda = ofv.GenCelda("","","right","50%","","","",sInput,"si")
    end if
  end if
  Celda = Celda & ofv.GenFForm()

  'Encabezado volver
  sInput = ofv.GenerarInput("boton2","Volver","submit","","20","bt","")
  Celdan = ofv.GenCelda("","","left","50%","","","",sInput,"si")
  Form = ofv.GenForm("Retornar","",sVuelve & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo,"","post","",Celdan,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",Celda & Form,"si")

  sHTML = sHTML & ofv.GenFTabla()  & "</center>"

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML& "</body></html>"
end if

response.write sHTML

%>