<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sOPCION = request.querystring("OPCION")
  sfamily = request.querystring("family")


  if sfamily = "" then sfamily = "D"
   vuelta  = ofv.vueltaasp2("../") & "?family=" & sfamily

  sHTML = ofv.FormHeader(session("FormN") & " - Propiedades del Mensaje")


'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


  'Encabezados
  Formulario = ""





  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0




  Strsql = "Select * From OBJINFO  where codobjeto = " & ncodigo & " and tipobjeto = '" & stipo & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

        strsql = "Insert into OBJINFO ("

        strsql = strsql & " CODOBJETO,TIPOBJETO,OBJtipo,OBJmensaje,"
        strsql = strsql & " OBJtitulo,OBJcopete,OBJimagen,OBJurl,"
        strsql = strsql & " OBJurldsc"

        strsql = strsql & " ) "
        strsql = strsql & " values ("

        strsql = strsql & "'" & ncodigo & "','" & stipo & "','1','',"
        strsql = strsql & "'TITULO DEL DOCUMENTO','','0','',"
        strsql = strsql & "''"

        strsql = strsql & ")"


        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        Strsql = "Select * From OBJINFO  where codobjeto = " & ncodigo & " and tipobjeto = '" & stipo & "' "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if


  if not rs.eof then


        VDOPCION = RS(7) & ""

        IF VDOPCION = "" THEN VDOPCION = "0"


        sxopcVD = "<option value=0 >NO</option> "
        sxopcVD = sxopcVD & "<option value=1 >IMAGEN</option> "
        sxopcVD = sxopcVD & "<option value=2 >VIDEO</option> "

        sxopcVD = replace(sxopcVD,"value=" & VDOPCION,"value=" & VDOPCION & " selected")


     fila = ""
     sForm = "formu"



      Celda =  ""
      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Codigo: ","si")
      sInput = "<b>DOC" & NCODIGO & "</b>"
      Celda = celda & ofv.GenCelda("","cc style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","","20%","","","",sInput,"si")

      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Tipo: ","si")

      scomK = "1" & chr(9) & "2"
      scomD = "SOLO INFORMATIVO" & chr(9) & "INCLUIDO LA DISTRIBUCION"

      sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
      Celda = celda & ofv.GenCelda("","uti style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","","20%","","","",sInput,"si")

      Celda =  celda & ofv.GenCelda("","tt  ","","15%","","","","Relacion: ","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")







      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","Titulo","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea style='width:100%' name=" & rs(5).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 190) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=2>" & REPLACE(rs(5) & "","''","'") & " </textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","Copete","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea style='width:100%' name=" & rs(6).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(6) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","Mensaje","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea style='width:100%' name=" & rs(4).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 950) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=6>" & REPLACE(rs(4) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda =  ofv.GenCelda("","tt  ","","","","","","Imagenes/Video: ","si")
      sInput = "<select class=at name=" & rs(7).name  & selacc &  " >" & sxopcVD & "</select>"
      Celda = celda & ofv.GenCelda("","uti style='text-indent:0;border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","","","","","",sInput,"si")


   IF (RS(7) & "") = "1" THEN
                                                                        '"document.Botonx.submit();"
      sInput = ofv.GenerarInput("","Seleccionar","BUTTON","","10","bietx","document.getElementById('Botonx').submit();")                  
      Celda = celda & ofv.GenCelda("","uti colspan=4 style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","","","","","",sInput,"si")

  ''     sAccion = "EDGENERALINFOimg.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "&family=" & sfamily
   '    FormuX = ofv.GenForm("Botonx","",sAccion,"","post","","&nbsp;","si")



   else

       Celda = celda &  ofv.GenCelda("","ut colspan=4 style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","","","","","","&nbsp;","si")

   end if

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




   IF (RS(7) & "") = "2" THEN

      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","URL DEL VIDEO:","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(9).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(9) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   END IF



    sAccion = ""
    if bModificar then

      sInput1 = "&nbsp;"
      sInput2 = ofv.GenerarInput("","Actualizar","submit","","10","bietx","")

      Celdax = ""
      Celdax = celdax & ofv.GenCelda("","ut ","","","","","",sInput1,"si")
      Celdax = celdax & ofv.GenCelda("","ut ","","","","","",sInput2,"si")

      Filax  = ofv.genrow("","","","","","","",Celdax,"si")
      sInput = ofv.GenTabla("","","","","","0","","0","1","","",Filax,"si")


      Celda = ofv.GenCelda("","uti colspan=6 style='padding:10px 0px;border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' ","center","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJINFO&volver=" & vuelta
      sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


    end if










    posi = posi + 1





  end if


       sAccionx = "EDGENERALINFOimg.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "&family=" & sfamily
       FormuX = ofv.GenForm("Botonx","",sAccionx,"","post","","&nbsp;","si")



      Formulario =  ofv.GenTabla("","aaa","","80% ","","0","","0","1","","",Fila,"si")
      Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Formulario,"si")
  'sHTML = sHTML & "<center><br>" & Formulario & "</center>" & FormuX  & Formules
  shtml = shtml & "<center><br>" & Formulario & FormuX & "</center>" '& FormuX




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  buscar = ""

  if  sfamily = "D" then
      sAccion = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  else
      sAccion = "edGENKomp.asp?Codcomp=" & Ncodigo & "&tipocomp=" & stipo & "  "
  end if

  sRetornar   = ofv.BotonVolver(sAccion)



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,sRetornar)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>