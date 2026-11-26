<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  bModificaresp = ofv.AModificar
  if session("Nomodificar") = "S" then

     if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False

  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  SELIM=request.querystring("ELIM")
  SOBJID=request.querystring("OBJID")
  vuelta = ofv.vueltaasp2("../")

  sHTML = ofv.FormHeader(session("FormN"))

   if sfamily = "" then sfamily = "D"
   vuelta  = ofv.vueltaasp2("../") & "?family=" & sfamily
 shtml = shtml & CorPnlFnc("N","")
 xxform=""
'----------------------------------------------


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"




set cn = ofv.conectar(ofv.strconn2)


IF SELIM="TODOS" THEN
SQL= "DELETE From ObjLNK Where Ltipo = 'L' and codobjetoL = " & nCodigo
CALL ofv.crearconsultaEx(SQL,cn,1,volver)
ELSEIF SELIM="SI" THEN
SQL= "DELETE From ObjLNK Where OBJID=" & SOBJID
CALL ofv.crearconsultaEx(SQL,cn,1,volver)
END IF
  'links asociados
  fila=""
  strsql = "Select * From ObjLNK Where Ltipo = 'L' and codobjetoL = " & nCodigo
  'response.write strsql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,volver)


 celda =   ofv.GenCelda("","ttt colspan=2 ","","","","","","ESQUEMAS QUE REFERENCIAN EL DOCUMENTO","si")
 if not rs1.eof then
 ' sNombre, sHref,CLAVE

    XREF="ELIMINARLINKS.ASP?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo &"&ELIM=TODOS"
    celda =  celda & BotonEliminarlnk("Confirma la eliminacion de TODAS las Relaciones?","ELIMINAR&nbsp;TODOS", XREF,nCodigo)
    Fila  =   ofv.GenRow("","","","","","","",Celda,"si")
     do while not rs1.eof
      XREF="ELIMINARLINKS.ASP?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo &"&ELIM=SI&OBJID="& RS1(0)
      Celda = ofv.BotonEliminar("Eliminacion",XREF)
      'celda= BotonEliminarlnk("Confirma la eliminacion de la relacion?" ,"", XREF,rs1(0))
      celda= celda & ofv.GenCelda("","ut colspan=1","","","","","",LinkObj(rs1(1),rs1(2)),"si")
      Fila =  fila & ofv.GenRow("","","","","20","","",Celda,"si")
      rs1.movenext
     loop
  else
  Fila =  fila & ofv.GenRow("","","","","20","","",Celda,"si")
  end if
     tabla=ofv.GenTabla("panel","tbp","","90%","","0","","0","0","","",fila,"si")
     sHTML = sHTML & tabla & "</center></body></html>"

    sHTML = sHTML & "</center></body></html>"
  ' end if
end if

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  'cambios
   Listar = ""


  ' if bVersion = "SI" then
  '    if SESSION("SHWMOD") <> "N" then
'
  '    elseif SESSION("VIEWMOD") <> "N" then
'
  '       'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
  ''       Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")
'
  '       sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  '       sHTML = sHTML & ConsCambios(sasp)
'
  '       Celda = Listar


  '    end if


  ' end if




    '  sHTML = sHTML & GTCCALENDAR("D")


'''''''
  sHTML = sHTML & "</div>"



  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")
  sHTML = sHTML &  "<script type='text/javascript'>"
      sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
      sHTML = sHTML & "  {"
      sHTML = sHTML & "  var r=confirm(disp_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
      'sHTML = sHTML & "    alert(fr.name);"
      sHTML = sHTML & "    fr.submit();"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion!');"
      sHTML = sHTML & "    }"
    sHTML = sHTML & "  }</script>"

 if  sfamily = "D" then
      sAccion = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  else
      sAccion = "edGENKomp.asp?Codcomp=" & Ncodigo & "&tipocomp=" & stipo & "  "
  end if

  sRetornar   = ofv.BotonVolver(sAccion)
  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,sRetornar) & xxform

  call ofv.cerrarconsulta(rs1)
  call ofv.cerrarconn(cn)

response.write sHTML

%>

<%



 function LinkObj(id,tipo)
 dim rsx


  Strsql = "Select * From csobjetosRDO where codobjeto = " & id

  'response.write strsql & "<br>"
   set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rsx.eof then


      wNomL = rsx(1) & ": " & rsx(2) & " - " & rsx(3) & " - " & rsx(5) & " - " & rsx(6) & " - (" & rsx(7) & ") "
      cod=rsx(0)
    else
      wNomL = "S/D"
    end if
    call ofv.cerrarconsulta(rsx)


    sActcon = "Consultar.asp?Codigo=" & id & "&Tipo=" & tipo
    sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
   ' response.write sActcon
    sHref = "&nbsp;"


 sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & id & "'); return false;" & Chr(34)
 wNomL = "&nbsp;<u><b><i>" & wNomL & "</i></b></u>"
 sImagen = wNomL
 sLink = OFV.GenLink("", "nt title='Ver Documento.' ", sHref, "", "", sFuncion, sImagen, "si")
 LinkObj=sLink
 end function




  Function BotonEliminarlnk(titdelete,sNombre, sHref,CLAVE)

 	             wktitdelete=titdelete '"Confirma la eliminacion de TODAS las Relaciones?"
 	             if sNombre="" then
 	              sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0130;</font>"
 	              xclass="ut"
 	             else
 	              sImageELS    =sNombre
 	              xclass="ttt"
 	             end if
 	             sformobj   = "FRMOBJ" & CLAVE
 	             sAccion    = sHref
 	             formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
 	             sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

 	            BotonEliminarlnk=ofv.GenCelda("",xclass & " style='cursor:pointer;'  title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs & formobj ,"si")



End Function




%>