<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'  response.buffer = true

  sTitulo = request.querystring("Titulo")
  sTitulo = ofv.convertircar(sTitulo,"_"," ")

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = nCodigo

  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID

  sHTML = ofv.FormHeader(session("FormN"))

''  sHTML = "<html><body>"

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


''  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"
 '' sHTML = sHTML & "</div>"

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"


  sHTML = sHTML & "<br>"

''''''''''''-----------------------------------------------------------------------

  set cn = ofv.conectar(ofv.strconn2)



  SOPCION = REQUEST.QUERYSTRING("OPCION") & ""

  IF SOPCION = "UPD" THEN


'''session("Sfilename")


        strsql = "UPDATE OBJPRINTPNL SET IMPARCHIVO = 'PnlPrint/"  & session("Sfilename") & "' where  CODOBJETO = '" & nCodigo & "' AND TIPOBJETO = '" & sTipo & "' "
 '' RESPONSE.WRITE STRSQL & "<BR>"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)



  END IF







  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From OBJPRINTPNL where  CODOBJETO = '" & nCodigo & "' AND TIPOBJETO = '" & sTipo & "' "
 '' RESPONSE.WRITE STRSQL & "<BR>"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0

  if rs.eof then


        strsql = "Insert into OBJPRINTPNL ("
        strsql = strsql & " CODOBJETO,TIPOBJETO,IMPADMITE,IMPTIPO,IMPARCHIVO"
        strsql = strsql & " ) "
        strsql = strsql & " values ("
        strsql = strsql & "'" & nCodigo & "','" & sTipo & "','N','DA',''"
        strsql = strsql & ")"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


     Strsql = "Select * From OBJPRINTPNL where  CODOBJETO = '" & nCodigo & "' AND TIPOBJETO = '" & sTipo & "' "
 '' RESPONSE.WRITE STRSQL & "<BR>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF



  if not rs.eof then

     sADMITE = (rs(3) & "")
     simptipo = (rs(4) & "")

        sxopc3 = "<option value='D' >DIRECTO HTML</option> "
        sxopc3 = sxopc3 & "<option value='DA' >DESDE ARCHIVO</option> "
        sxopc3 = replace(sxopc3,"value='" & rs(4) & "'","value='" & rs(4) & "' selected")


        sxopc4 = "<option value='S' >SI</option> "
        sxopc4 = sxopc4 & "<option value='N' >NO</option> "
        sxopc4 = replace(sxopc4,"value='" & rs(3) & "'","value='" & rs(3) & "' selected")



     fila = ""
     sForm = "formu01"




      Celda =  ""
      sInput = "<b>CONFIGURACION DE IMPRESION</b>"
      Celda = celda & ofv.GenCelda("","TTT colspan=2 ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ""

      Celda = celda & ofv.GenCelda("","tt  ","","20%","","","","ADMITE IMPRESION: ","si")

      sInput = "<select class=at name=" & rs(3).name  & selacc &  " >" & sxopc4 & "</select>"
      Celda = celda & ofv.GenCelda("","UTI  ","","80%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ""

      Celda = celda & ofv.GenCelda("","tt  ","","","","","","TIPO DE IMPRESION: ","si")

      sInput = "<select class=at name=" & rs(4).name  & selacc &  " >" & sxopc3 & "</select>"
      Celda = celda & ofv.GenCelda("","UTI  ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


     if  (rs(4) & "") = "DA" and (rs(3) & "") = "S" then

      Celda =  ""

      Celda = celda & ofv.GenCelda("","tt  ","","","","","","ARCHIVO: ","si")

      sInput = (rs(5) & "")

      IF  sInput = "" or sinput ="PnlPrint/" THEN
          sInput = "SIN ASIGNAR"
      ELSE
          sInput = "<A CLASS='NT' HREF='" & sInput & "' TARGET='_BLANK' ><U>" & sInput & "</U></A>"
      END IF


      Celda = celda & ofv.GenCelda("","UTI  ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ""
      Strsql = "Select perfil From tipobjeto Where tipobjeto ='" & stipo & "' "
      'response.write Strsql
      set xperf= ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not xperf.eof then
       xdperfil=xperf(0)
      else
       xdperfil=""
      end if
      call ofv.cerrarconsulta(xperf)
      sAccion = "reportesgen.asp?ID=" & nCodigo & "&tipobj=" & stipo & "&GLBSOLN=ESPREP&perfil=" & xdperfil

     ''' sAccion = "CoObjRecoADDUSR.asp?Opcion=ADDUSR&Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo & "&mensaje=" & msgcod
      sInput = ofv.GenerarInput("","Generar archivo","button","","10","btsb"," disp_confirm('Confirma la genracion del archivo de impresion del documento???','Formux2'); ")

      sInputx2 = ofv.GenerarInput("formato","WORD","hidden","","10","dt","")
      formux2 = ofv.genform("Formux2","",sAccion,"printi","post","",sInputx2,"si")


      Celda = celda & ofv.GenCelda("","UTI colspan=2  ","right","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



     end if



      Celda = ofv.GenCelda("","ut colspan=4;'  ","","","","","","&nbsp;","si")
      Fila  = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then

        sInput = ofv.GenerarInput("ppp","Actualizar","BUTTON","","10","bietx","Ej_submit('" & sForm & "');")

      Celda = ofv.GenCelda("","uti ","","","","","",sInput,"si")

          sInput = "&nbsp;"

      Celda = Celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")

          sInput = "&nbsp;"


      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")


      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJPRINTPNL&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    end if

    posi = rs(0)


  else





  end if

  Formulario =  ofv.GenTabla("","aaa","","80% ","","0","","0","1","","",Fila,"si")
  sHTML = sHTML & ofv.genform(sForm,"",sAccion,"","post","",Formulario,"si") & "<br>"

  if not rs.eof then

     if  simptipo = "DA" and sadmite = "S" then

         sPath = Request.ServerVariables("PATH_INFO")
         Xruta = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT")  & left(sPath,instrRev(sPath,"/",-1) - 1)
         sPath = server.mappath(sPath)
         sPath = replace(sPath,"EDPRINTPNL.asp","",1,-1,1) & "PnlPrint"


         session("directorio") = sPath

         session("edcodobj") = nCodigo
         session("edtipobj") = sTipo

        sHTML = sHTML & "<form enctype=" & chr(34) & "multipart/form-data" & chr(34) & " action='EdPrintPnlUploadSave.asp' method=post>"

    Celda = ofv.gencelda("","tt","","","","","","Archivo de Impresion:","si")
    Celda = Celda & ofv.gencelda("","ut","","","","","","<input class=nt name=" & chr(34) & "FILE1" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt  ","","","","","","Archivo a enviar 2:","si")
    Celda = Celda & ofv.gencelda("","ut","","","","","","<input class=nt name=" & chr(34) & "FILE2" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","aaa style='display:none;' ","","","","","",Celda,"si")


    Celda = ofv.gencelda("","tt","","","","","","Archivo a enviar 3:","si")
    Celda = Celda & ofv.gencelda("","ut","","","","","","<input class=nt name=" & chr(34) & "FILE3" & chr(34) & " type=" & chr(34) & "file" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","aaa style='display:none;' ","","","","","",Celda,"si")


    Celda = ofv.gencelda("","ut","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","ut","right","","","","","<input class='btsb' type=" & chr(34) & "submit" & chr(34) & " value=" & chr(34) & "Cargar" & chr(34) & ">","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




     sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si") & "</form>"


    end if
  end if


  Agregar = ""

  buscar = ""
  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

    sHTML = sHTML & "<script type='text/javascript'>"

    sHTML = sHTML & "function Ej_submit(disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "  }"

    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "function send_confirm(send_msg,send_href,send_aid)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(send_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('ok ');"
    sHTML = sHTML & "    var fr= document.getElementById(send_aid);"
    sHTML = sHTML & "    fr.href = send_href;"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion! ');"
  '  sHTML = sHTML & "    return void(0);"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"

    sHTML = sHTML & FormuX2


''''''''''''-----------------------------------------------------------------------

  sHTML = sHTML & "<br>"

'''''''''''''''''''''''''''''
'''''''
  sHTML = sHTML & "</span></div>"



  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")





  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Function Fechador(wkfecha)


      wkfecha=clng(wkfecha)
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>