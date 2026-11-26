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

  nAccion = request.querystring("Accion")
  nProced = request.querystring("Codproced")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  volver = request.querystring("volver")

''response.write nAccion & " nAccion " & chr(13)
''response.write nProced & " nProced " & chr(13)
''response.write nCodigo & " nCodigo " & chr(13)


  sVuelve = ofv.convertircar(volver,"_","&")
  sVuelve = ofv.convertircar(sVuelve,"*","&")
  sVuelta = ofv.convertircar(sVuelve,"&","*")
  sVuelve = mid(sVuelve,4,len(sVuelve) - 3)


  sHTML = ofv.FormHeader(session("FormN"))


'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  set cn2 = ofv.conectar(ofv.strconn2)

''''Para obtener la descripcion de la Accion
  ''Primero obtiene el tipo de accion de la tabla objaccion
    Strsql = "SELECT * From OBJACCION "
    Strsql = Strsql & " WHERE codproced = '" & nProced & "' "
    Strsql = Strsql & " AND objid = '" & nAccion & "' "
    set rsz = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	if not rsz.eof then
	   nTAccion  = rsz(4)
	end if
    call ofv.cerrarconsulta(rsz)

  ''Despues obtiene la descri de la accion de la tabla Tipoaccalt
    Strsql = "Select descri From Tipoaccalt "
    Strsql = Strsql & " Where tipoaccionalt = '" & nTAccion & "' "

    set rsy = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	if not rsy.eof then
	   nTdescri  = rsy(0)
	end if
    call ofv.cerrarconsulta(rsy)

    fila = ""

    'Encabezado Referencia
    Celda = ofv.GenCelda("","ttt","","100%","","","","Referencia del paso - " & nTdescri,"si")
    fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")



  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "SELECT Referencia,Referencia2,Referencia3,rellink From Objaccion Where objid = " & nAccion

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then

      wasotxt = "N"
      stylob  = ""

    IF WKAPLIC <> "PL" THEN


         if not isnull(rs(3))  then
            if  clng(rs(3)) > 0 then

                Strsql = "SELECT * From OBJLNK where objid = 0" & rs(3)
                set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsx.eof then
                   wasotxt  = rsx(14) & ""
                   IF wasotxt = "" THEN wasotxt = "N"
                end if
                call ofv.cerrarconsulta(rsx)

            end if
          end if

       IF wasotxt = "S" THEN

          stylob = "ut style='border-style:solid;border-width:1;border-color:#990000;' "

       end if


    END IF


    'Texto de Referencia

    celda = "<font class=at><textarea class=dt name=referencia rows=6 cols=100 "
    if bModificar and wasotxt = "N" then
       celda = celda & " onkeypress=" & chr(34)
       celda = celda & " if (this.value.length > 2450) {alert('El texto es demasiado"
       celda = celda & " extenso'); return false;}" & chr(34)

    else
       celda = celda & " readonly"

    end if
    celda = celda & " >" & rs(0) & "</textarea></font>"

    Celda = ofv.GenCelda("",stylob,"","","","","",celda,"si")
    fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


  IF WKAPLIC <> "PL" THEN

    'Texto de Referencia
    celda = "<font class=at><textarea  class=dt name=referencia2 rows=6 cols=100 "
    if not bModificar then
       celda = celda & " readonly"
    else
       celda = celda & " onkeypress=" & chr(34)
       celda = celda & " if (this.value.length > 2450) {alert('El texto es demasiado"
       celda = celda & " extenso'); return false;}" & chr(34)
    end if
    celda = celda & " >" & rs(1) & "</textarea></font>"

    Celda = ofv.GenCelda("","ut","","","","","",celda,"si")
    fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


    'Texto de Referencia
    celda = "<font class=at><textarea  class=dt name=referencia3 rows=6 cols=100 "
    if not bModificar then
       celda = celda & " readonly"
    else
       celda = celda & " onkeypress=" & chr(34)
       celda = celda & " if (this.value.length > 2450) {alert('El texto es demasiado"
       celda = celda & " extenso'); return false;}" & chr(34)
    end if
    celda = celda & " >" & rs(2) & "</textarea></font>"

    Celda = ofv.GenCelda("","ut","","","","","",celda,"si")
    fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


  END IF

 end if



    sAccion = ""
    if bModificar then



     'Encabezado Actualizar
      sInput = ofv.GenerarInput("boton","Actualizar","submit","","20","bt","")
      Celda = ofv.GenCelda("","","","","","","",sInput,"si")
      fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

      'Encabezado Actualizar
      sAccion = "asp/ActualizarEN.asp?Tabla=Objaccion&volver=" & volver & "&Codproced="
      sAccion = sAccion & nProced & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID="
      sAccion = sAccion & nAccion & "&Referencia=SI"

      'fila =  ofv.GenForm("fmnact","",sAccion,"","post","",fila,"si")
      tabla= ofv.GenTabla("","aaa  ","","80%","","0","0","0","0","","",fila,"si")
      formu =  ofv.GenForm("fmnact","",sAccion,"","post","",tabla,"si")


    end if


    sHTML = sHTML & "<br>" & formu' ofv.GenTabla("","aaa  ","","80%","","0","0","0","0","","",fila,"si")










'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")
  sAccion = sVuelve
   Celda   = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>