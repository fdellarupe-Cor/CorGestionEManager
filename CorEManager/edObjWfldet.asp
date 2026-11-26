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
  sVolver = request.querystring("volver")
  sCod    = request.querystring("cod")
  SACC=REQUEST.QUERYSTRING("ACC")
   Sckt=REQUEST.QUERYSTRING("ckt")
  vuelta  = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")
   IF SACC="PEND" THEN
    session("CKTTYPE")=Sckt
   END IF

   set cn = ofv.conectar(ofv.strconn2)

   sAutorizantes=" Autorizantes"
   Strsql = "SELECT tipocpt FROM tipobjeto where tipobjeto='" & sTipo & "'"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if not rs.eof then
    if rs(0)="SN" then  sAutorizantes= " desarrolladores"
  end if

  call ofv.cerrarconsulta(rs)


  WKOPCION = session("CKTTYPE")

'response.write request.querystring
  IF WKOPCION = "APRO" THEN
     WKOPCIONdsc = "OBSERVACION"
  ELSE
     WKOPCIONdsc = "AUTORIZACION"
  END IF


'RESPONSE.WRITE session("CKTTYPE")
  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------
if sTipo <> "SN" then
  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
end if
'-------------------------------------------------

'shtml = shtml & CorPanelTop()



  Celda = ofv.gencelda("","ttt colspan=3","center","","","","","<big>Circuito de " & WKOPCIONdsc & "</big>","si")
  Fila  = ofv.GenRow("","","","","","","",Celda,"si")

  Celda = ofv.gencelda("","ttt colspan=3","center","","","","","Lista de " & sAutorizantes ,"si")
  Fila  = fila & ofv.GenRow("","","","","20","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","ttr","","1%","","","","Orden","si")
  Celda = Celda & ofv.gencelda("","ttr","","98%","","","","Descripcion","si")
  Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")


  xactiva = 0
  xmanda = 0

  Strsql = "SELECT * FROM Defwfl"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

     xactiva = rs(1)


  end if

  call ofv.cerrarconsulta(rs)

  if  xactiva then xmanda = 1

''session("CKTTYPE")


  Strsql = "SELECT * FROM OBJwfldet  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' And CKTTYPE = '" & session("CKTTYPE") & "' order by orden,objid"
'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  Formulario = ""

  vinoSC = 0

  sImagenU = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>&#0128;</FONT></B>"

  do until rs.eof

    sForm = "wfl" & rs(0)

    sHref = ""

    bModificarR = bModificar

    xorden = rs(3)

      if bModificarR then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      end if

    if not isnumeric(xorden) then xorden = 0

    if xmanda and xorden > 8999 then bModificarR = 0

    if bModificarR then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver="
      sHref = sHref  & vuelta & "?Cod=" & sCod
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Usuario",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    'Orden


    if bModificarR then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm & "x","2","dtn","SI")
      formu=ofv.GenForm(sForm & "x","",sAccion,"","post","",sInput,"si")
     'sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfldet&volver=" & vuelta
     'sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    else
       sinput = ofv.gencelda("","ut","center","","","","",rs(3) & ".&nbsp;","si")
    end if
    Celda =Celda & ofv.gencelda("","ut","center","","","","",formu,"si")
   '' Celda =  sinput



    sTabla = "Usuario"
    sBusca = rs(4)
    sGrpUsr = "S/D"

   '   02/08/2021 Busca en el arbol AD codemp del usuario  -> ALE
   '10317-0004/20
     SELECT CASE session("APLMDCLI")

		  CASE "GP"

        'xxxx     = getInfoUser(sBusca)
         '
		IdCodemp=rs("codemp") & ""

		IF IdCodemp="" THEN IdCodemp=1

		xxxx=getInfoUserEX(sBusca,IdCodemp)
		 CASE ELSE
		       xxxx     = getInfoUser(sBusca)
		 END SELECT
        '-----------------------------------------------

        sGrpUsr = session("SGU_nombre")

        sGrpUsr = sImagenU & "&nbsp;" & sBusca & " - " & sGrpUsr



    Celda = celda & ofv.gencelda("","ttt ","","98%","","","",sGrpUsr,"si")

    'Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")


   rs.movenext
   loop
t = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#999999;' ","","80%","","0","0","-2","1","","",fila,"si")

  sHTML = sHTML & "<center>"
  sHTML = sHTML & t
  sHTML = sHTML & "</center>"

  if bModificar then
     sAccion = "AgregarLista.asp?Codobjeto=" & nCodigo
     sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
     Agregar = ofv.BotonAgregar(sAccion)
  end if

  if sTipo <> "SN" then
  	sAccion  = "edobjwfl.asp?upg=S&Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo
  else
	sAccion  = "crearNormativa.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&acc=ABR&DOCMODAL=WKPED&pedido=VER"
  end if
  Retornar = ofv.BotonVolver(sAccion)


if sTipo <> "SN" then
 ' sHTML = sHTML & "</center>"

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice


  	sHTML = sHTML & CorPnlIndx("N","")
  end if
'--------------------------------------------------------------------------------
 sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")


  call ofv.cerrarconsulta(rs)
'  response.write request.querystring
  call ofv.cerrarconn(cn)
  shtml=shtml & "<script> MoveBtnas();</script>"
  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>