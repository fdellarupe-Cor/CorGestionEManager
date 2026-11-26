<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
CONT = 0
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
  sOPCION = request.querystring("OPCION")
  vuelta  = ofv.vueltaasp2("../")

  sHTML = ofv.FormHeader(session("FormN"))


'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


  'Encabezados
 ' Formulario = ofv.Encabezados("#","OBSERVACIONES")

  celdas = ofv.GenCelda("","ttr","","2%","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","ttr","","2%","","","","#","si")
  celdas = celdas & ofv.GenCelda("","ttr","","96%","","","","OBSERVACIONES","si")
  fila = ofv.GenRow("","","","","20","","",celdas,"si")

  sImagenI   = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"

  set cn = ofv.conectar(ofv.strconn2)

  IF  SOPCION = "SI" THEN
      Strsql = "insert into OBJGENCAM  (codobjeto,tipobjeto,SECUEN,DESCRI) values ("
      Strsql = Strsql & nCodigo & ",'" & stipo & "','10',' ')"
      call ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF

  Strsql = "Select * From OBJGENCAM Where codobjeto=" & nCodigo & " Order By SECUEN"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do  until rs.eof
  CONT = CONT +1
      sForm = "edP" & rs(0)

    'Eliminar
     if  bModificar then
         sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=OBJGENCAM&volver=" & vuelta
         sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
     end if

     if  sHref <> "" then
         Celda = ofv.BotonEliminar(rs(0),sHref)
     else
         Celda = ofv.GenCelda("","","","","","","","","si")
     end if

    'SECUEN
    if  bModificar then
        sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJGENCAM&volver=" & vuelta
        sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm & "X1","5","dtn","")
        sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")
    else
        sInput = ofv.GenerarInput("",rs(4),"readonly","","5","dt","")
    end if

    Celda  = Celda & ofv.GenCelda("","UT","","2%","","","",sInput,"si")

    'DESCRI
      celda3 = "<font class=at><textarea name=" & rs(3).name & " rows=4 cols=100 "
      if bModificar then

        sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJGENCAM&volver=" & vuelta
        sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

        celda3 = celda3 & " onblur=" & sform & "X2" & ".submit() "
        celda3 = celda3 & " onkeypress=" & chr(34)
        celda3 = celda3 & " if (this.value.length > 950) {alert('El texto es demasiado"
        celda3 = celda3 & " extenso'); return false;}" & chr(34)

      else
        celda3 = celda3 & " readonly"

      end if

     celda3 = celda3 & " >" & rs(3) & "</textarea></font>"

    sInput = celda3
    sInput = ofv.GenForm(sForm & "X2","",sAccion,"","post","",sInput,"si")
    Celda  = Celda & ofv.GenCelda("","UT","","96%","","","",sInput,"si")

    Fila   = fila & ofv.GenRow("","","","","","","",Celda,"si")



    rs.movenext
  loop

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","-2","-1","","",Fila,"si")
  sHTML = sHTML & formulario




  'Listar Organigrama

  xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Visualizar cambios", "NO", "Cambios")
  sAccion = "LisOrganigrama.asp"
 ''''' test  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")

  test  = ofv.GenBoton(16, "esp1Bot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")


    Agregar = ""

  'Botonera
  if  bModificar then
      sAccion = "EdgeneralCambios.asp?OPCION=SI&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo
      Agregar = ofv.BotonAgregar(sAccion)
  end if

  sAccion = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  sRetornar = ofv.BotonVolver(sAccion)



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  if CONT = 0 then
    	    sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes cambios al documento agregados.</div>"
    	else
    	    tabla = ofv.GenTabla("TCEX1","onresize='DCEXRCLC();","","80%","","0","","4","0","","",fila,"si") ''''''tabla = ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();","","80%","","0","","4","0","","",fila,"si") quito aaa del onresize CG 12-03-26
    	    sHTML = sHTML & tabla
	end if
  'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
  'response.write "123"

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML   = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","",test,Agregar,sRetornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

  sHTML = sHTML & ConsCambios(sasp)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>

