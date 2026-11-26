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
  bModificaresp = ofv.AModificar
  if session("Nomodificar") = "S" then

     if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False

  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  vuelta  = ofv.vueltaasp2("../")

  set  cn = ofv.conectar(ofv.strconn2)

  sHTML   = ofv.FormHeader(session("FormN"))


'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


  'Encabezados

  Formulario = ""

  celdas = ""
  celdas = celdas & ofv.GenCelda("","ttr","","100%","","","","palabras de busqueda","si")
  Formulario = ofv.GenRow("","","","","20","","",celdas,"si")


 '' Formulario = Formulario & ofv.Encabezados("EMPRESA")

  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"

  XINCD = ""
  XINCS = ""

  Fila   =  Formulario

  Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF  NOT RS.EOF THEN
      nCol = 100

      if  len(wdescri) = 0 then wdescri = "Sin Descripcion" 

      if  not  isnull(rs(8)) then 
          wdescri = rs(8)
      end if 
  
      t3b = "<font class=at><textarea class=dt "
      t3b = t3b & "name=" & rs(8).name & " cols=" & nCol & " rows=15 "
      if  bModificar then
          t3b = t3b & " onblur=edGeneral.submit() "
          t3b = t3b & " onkeypress=" & chr(34)
          t3b = t3b & " if (this.value.length > 950) {alert('El texto es demasiado"
          t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
      else
          t3b = t3b & " readonly >"
      end if
      xdatos =  wdescri
      t3b    = t3b & xdatos & "</textarea></font>"
      Celda  = ofv.GenCelda("","abc colspan=2","center","100%","","","",t3b,"si")
      Fila   =  fila & ofv.GenRow("","","","","","","",Celda,"si")

      sAccion = ""
      if  bModificar then
          sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objetos&volver=" & vuelta
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      end if
      formulario= ofv.GenTabla("","aaa","","70%","","0","","2","0","","",fila,"si")
      Formulario = ofv.GenForm("edGeneral","",sAccion,"","post","",formulario,"si")


  end if

  call ofv.cerrarconsulta(rs)



    sHTML = sHTML & "<center>"
    sHTML = sHTML & formulario'ofv.GenTabla("","","","70%","","0","","2","0","","",Formulario,"si")
    sHTML = sHTML & "</center>"


  'Boton Volver
  sAccion   = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  sRetornar = ofv.BotonVolver(sAccion)

  'Botonera
  Buscar  = ""

  

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,sRetornar)


  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>

