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
  OPC=Request.querystring("OPC")
  ID = nCodigo

  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID

  sHTML = ofv.FormHeader(session("FormN"))

''  sHTML = "<html><body>"
  sImagenU = "<FONT FACE=corflaticon COLOR=#AA0000 SIZE=3 >&#0128;</FONT>"

  sImagenGM = "<FONT FACE=corflaticon COLOR=#005500 SIZE=5 >&#0148;</FONT>"
 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


''  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"
 '' sHTML = sHTML & "</div>"

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"


  sHTML = sHTML & "<center><br>"

''''''''''''-----------------------------------------------------------------------

  set cn = ofv.conectar(ofv.strconn2)


  SOPCION =REQUEST.QUERYSTRING("OPCION") & ""

  IF OPC="ELIM" THEN
    OBJID=REQUEST.QUERYSTRING("ID")
    SQL="DELETE FROM novegenpar WHERE OBJID=" & OBJID
    SET RSREC= ofv.crearconsultaEx(SQL,cn,1,parametros)

    SQL=" DELETE FROM  novegenpardist WHERE CODIGO=" & OBJID
    SET RSREC= ofv.crearconsultaEx(SQL,cn,1,parametros)
      SOPCION="SI"
  END IF

 ' Strsql = "Select * From novegenpar where MSGTIPOMESS = '9' AND CONTACTO = '" & SESSION("USUARIO") & "' AND MSGOBJCOD = '" & nCodigo & "' AND MSGOBJTIP = '" & sTipo & "' "
 ' RESPONSE.WRITE  Strsql
 ' IF session("APLMDCLI")="BGAL" THEN
 ' SET RSREC= ofv.crearconsultaEx(StrSql,cn,1,parametros)
 ' IF RSREC.EOF THEN SOPCION="SI"
 ' CALL OFV.CERRARCONSULTA(RSREC)
  'END IF







  IF SOPCION = "SI" THEN


     nFecini = 0

     if nFecini < clng(cdate("01-01-1900")) then
        nFecini = clng(cdate("01-01-1900"))
     else
        nFecini = 0
     end if

        xmsgusr = "Recomiendo la lectura de este documento"
        dFechai = clng(date()) - nFecini
        dFechae = 0



        strsql = "Insert into novegenpar ("
        strsql = strsql & " fecvig,fecexp,dias,texto,"
        strsql = strsql & " importancia,tipodist,contacto,enviado,"
        strsql = strsql & " MSGVERSION,MSGTITULO,MSGCOPETE,MSGIMAGEN,"
        strsql = strsql & " MSGURL,MSGURLDSC,"
        strsql = strsql & " MSGVAENC,MSGVAFORO,MSGCODENC,MSGCODFORO,"
        strsql = strsql & " MSGTIPOMESS,MSGAPLIC,MSGOBJCOD,MSGOBJTIP"
        strsql = strsql & " ) "
        strsql = strsql & " values ("
        strsql = strsql & dFechai & "," & dFechae & ",1,'" & xmsgusr & "',"
        strsql = strsql & "1,2,'" & Session("Usuario") & "','N',"
        strsql = strsql & "2,'Recomendacion','Recomendacion de lectura','0',"
        strsql = strsql & "'1','',"
        strsql = strsql & "'N','N','0','0',"
        strsql = strsql & "'9','" & wkaplic & "','" & nCodigo & "','" & sTipo & "'"
        strsql = strsql & ")"


        call ofv.crearconsultaEx(StrSql,cn,1,parametros)



  END IF







  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From novegenpar where MSGTIPOMESS = '9' AND CONTACTO = '" & SESSION("USUARIO") & "' AND MSGOBJCOD = '" & nCodigo & "' AND MSGOBJTIP = '" & sTipo & "' "
 ' RESPONSE.WRITE STRSQL & "<BR>"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  if not rs.eof then

     senenviado = (rs(8) & "")

        sxopc3 = "<option value=1>1 DIA</option> "
        sxopc3 = sxopc3 & "<option value=2>2 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=7>7 DIAS</option> "
        sxopc3 = sxopc3 & "<option value=10>10 DIAS</option> "


     '   sxopc3 = sxopc3 & "<option value=15 >15 DIAS</option> "
     '   sxopc3 = sxopc3 & "<option value=20 >20 DIAS</option> "
     '   sxopc3 = sxopc3 & "<option value=25 >25 DIAS</option> "
     '   sxopc3 = sxopc3 & "<option value=30 >30 DIAS</option> "
     '   sxopc3 = sxopc3 & "<option value=35 >35 DIAS</option> "
     '   sxopc3 = sxopc3 & "<option value=40 >40 DIAS</option> "

        sxopc3 = replace(sxopc3,"value=" & rs(3)& ">" ,"value=" & rs(3) & " selected>")


        sxopc4 = "<option value=0 >ALTA</option> "
        sxopc4 = sxopc4 & "<option value=1 >NORMAL</option> "
        sxopc4 = sxopc4 & "<option value=2 >BAJA</option> "

        sxopc4 = replace(sxopc4,"value=" & rs(5),"value=" & rs(5) & " selected")

        sxopc5 = "<option value=1 >TODOS</option> "
        sxopc5 = sxopc5 & "<option value=2 >LISTA</option> "


        sxopc5 = replace(sxopc5,"value=" & rs(6),"value=" & rs(6) & " selected")

        VDOPCION = RS(12) & ""

        IF VDOPCION = "" THEN VDOPCION = "0"


        sxopcVD = "<option value=0 >NO</option> "
        sxopcVD = sxopcVD & "<option value=1 >IMAGEN</option> "
        sxopcVD = sxopcVD & "<option value=2 >VIDEO</option> "

        sxopcVD = replace(sxopcVD,"value=" & VDOPCION,"value=" & VDOPCION & " selected")


     fila = ""
     sForm = "formu01"




      Celda =  ofv.GenCelda("","tt  ","","","","","","Codigo: ","si")
      sInput = "<b>RGP" & rs(0) & "</b>"
      Celda = celda & ofv.GenCelda("","cc colspan=3 ","","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")


    sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0167;</FONT></B>"

    sAccion = "AdNovedadesFec.ASP?XFNAME=NO&id=" & id & "&CantRegMover=" & CantRegMoverx & "&XFSRD=0" & rs(1)

    sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
    sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

    sinputE = ""


      sInput = Fechador(rs(1)) '''' cdate(clng(rs(1)) + clng(cdate("01-01-1900")))

   '''   sInput = ofv.GenerarInput(rs(1).name,sInput,"agregar",sForm,"60","dt","")

      Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput & sinputE,"si")

      Celda = celda & ofv.GenCelda("","tt  ","","","","","","Dias de Vigencia: ","si")

      sInput = "<select class=at name=" & rs(3).name  & selacc &  " >" & sxopc3 & "</select>"

 '     sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"60","dt","")
      Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = ofv.GenerarInput(rs(9).name,"2","hidden",sForm,"1","dt","")

      Fila = fila & ofv.genrow("","","","","","","","<br>" & sInput,"si")


      'if  session("APLMDCLI") <> "BGAL" THEN
      Celda =  ofv.GenCelda("","tt colspan=4 ","","","","","","&nbsp;&nbsp;titulo","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(10).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 190) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=2>" & REPLACE(rs(10) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=4 ","center","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ofv.GenCelda("","tt colspan=4 ","","","","","","&nbsp;&nbsp;copete","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(11).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(11) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=4 ","center","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
    ' END IF



      Celda =  ofv.GenCelda("","tt colspan=4 ","","","","","","&nbsp;&nbsp;Mensaje","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(4).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 950) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=6>" & REPLACE(rs(4) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=4 ","center","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Fila = fila & ofv.genrow("","","","","","","","<br>","si")



      Celda =  ofv.GenCelda("","tt  ","","","","","","IMAGENES/VIDEO: ","si")
      sInput = "<select class=at name=" & rs(12).name  & selacc &  " >" & sxopcVD & "</select>"
      Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")


   IF (RS(12) & "") = "1" THEN

      sInput = ofv.GenerarInput("","SELECCIONAR","BUTTON","","10","btSB","document.Botonx.submit();")
      Celda = celda & ofv.GenCelda("","cc colspan=2 ","","","","","",sInput,"si")

       sAccion = "Adnovedadesimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id
       FormuX = ofv.GenForm("Botonx","",sAccion,"","post","","&nbsp;","si")



   else

       Celda = celda &  ofv.GenCelda("","cc colspan=2  ","","","","","","&nbsp;","si")

   end if

   ''''   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




   IF (RS(12) & "") = "2" THEN

      Celda =  ofv.GenCelda("","tt colspan=4 ","center","","","","","URL DEL VIDEO:","si")

   '''   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(14).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(14) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=4 ","center","70%","","","",sInput,"si")

  '''    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   END IF



      Celda = ofv.GenCelda("","ut colspan=4 style='border-bottom-style:dotted;border-bottom-width:1;border-bottom-color:#990000;'  ","","","","","","&nbsp;","si")
      Fila  = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then

       IF (RS(8) & "") = "N" THEN
          IF session("APLMDCLI") = "BGAL"	 THEN
          sInput = ofv.GenerarInput("ppp","Grabar Mensaje","BUTTON","","10","btsb","Ej_submit('" & sForm & "');")
          ELSE
           sInput = ofv.GenerarInput("ppp","Actualizar","BUTTON","","10","btsb","Ej_submit('" & sForm & "');")
          END IF
       else
         '  IF session("APLMDCLI") = "BPN" or session("APLMDCLI") = "BP" and true THEN
            nCodigo = request.querystring("Codobjeto")
            sTipo = request.querystring("Tipobjeto")
            XACCION="CoObjRECOMENDAR.asp?OPC=ELIM&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)
            FormuX2C = ofv.GenForm("Formux2C style='display:none'","",XACCION,"","post","","&nbsp;","si")
            sInput = ofv.GenerarInput("ppp","LIMPIAR","BUTTON","","10","btsb","disp_confirm('Confirma Limpiar la recomendación???','Formux2C');")
        '   ELSE
        '  sInput = "---"
        '  sInput = "Enviado"
         ' END IF
       end if

      Celda = ofv.GenCelda("","uti ","","","","","",sInput,"si")

          sInput = "&nbsp;"

      Celda = Celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")

          sInput = "&nbsp;"

      Celda = Celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")

     '  IF (RS(8) & "") = "N" THEN
        ' if session("APLMDCLI") <> "BGAL" and false THEN
          sInput = ofv.GenerarInput("","Enviar","BUTTON","","10","btsb"," disp_confirm('Confirma el envio de esta recomendacion???','Formux2'); ")
    '     END IF
'''document.Formux2.submit();

        ''''''''' sAccion = "Adnovedadesimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id

          sAccion = "CoObjRecoDESP.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo & "&mensaje=" & rs(0)
          FormuX2 = ofv.GenForm("Formux2","",sAccion,"","post","","&nbsp;","si")


    '   else
    '      sInput = "Enviado"
    '   end if

      Celda = Celda & ofv.GenCelda("","vti ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")


      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=NOVEGENPAR&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    end if
  ''  Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = rs(0)


  else



      sInput = "No ha recomendado este documento anteriormente, desea agregar un mensaje de recomendacion??"
      Celda = ofv.GenCelda("","uti colspan=4 ","","70%","","","",sInput,"si")

      Fila = ofv.genrow("","","","","30","","",Celda,"si")

      Celda = ofv.GenCelda("","ut colspan=4 style='border-bottom-style:dotted;border-bottom-width:1;border-bottom-color:#990000;'  ","","","","","","&nbsp;","si")
      Fila  = fila & ofv.genrow("","","","","","","",Celda,"si")

      sInput = ofv.GenerarInput("","Agregar","submit","","10","btsb","")
      Celda = ofv.GenCelda("","uti colspan=4","","10%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")
      sAccion = "CoObjRecomendar.asp?Opcion=SI&Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo

    ''  Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


  end if

  Formulario =  ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;' ","","95% ","","0","","0","1","","",Fila,"si")
  sHTML = sHTML & ofv.genform(sForm,"",sAccion,"","post","",Formulario,"si") & "<br>"

 '''' sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;' ","","95% ","","0","","0","1","","",Formulario,"si") & "<br>"

  if posi > 0 then

     msgcod = posi

      Formulario = ""

     Strsql = "Select * From novegenpardist where codigo = '" & posi & "' "
'  RESPONSE.WRITE STRSQL & "<BR>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     posi = 0
     do until rs.eof

       sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=novegenpardist&volver=" & vuelta
       sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"

   ''    sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
    ''   sHref = sHref & "&Tabla=novegenpardist&volver=" & vuelta & "&Eliminar=NSNC"

       'if senenviado <> "S" then
          Celda = ofv.BotonEliminar("Usuario",sHref)
      '' else
         ' Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
     '  end if


      sBusca = rs(2)

    if Ucase(sBusca) = "SECT-" then

       sGrpUsr = "S/D"


       Strsql = "SELECT * FROM eco"
       Strsql = Strsql & " Where eco = 0" & rs(3)

       set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#993333><b>Sector: </b></font>" & rx(1)
       call ofv.cerrarconsulta(rx)


       sinput = sGrpUsr


    elseif Ucase(sBusca) = "TODOS" then

       sGrpUsr = "S/D"


       Strsql = "SELECT * FROM CLIENTEPLAN"
       Strsql = Strsql & " Where ID = 0" & rs(3)
       set rx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rx.eof then sGrpUsr = "<font color=#336699><b>EMPRESA: </b></font>" & rx(1)
       call ofv.cerrarconsulta(rx)


       sinput = sGrpUsr


    else

       sGrpUsr = "S/D"
       set cx = ofv.conectar(ofv.strconn0)

        IF mid(sBusca,1,4)="GRM-" THEN
	    sgrupoM=mid( sBusca  ,5,len(sBusca)  )
	    xxx=ucase(getMailDelGrupo(sgrupoM))
	   ' RESPONSE.WRITE sBusca & "<BR>"
	    sGrpUsr = sImagenGM & "&nbsp;&nbsp;<B>GRUPO DE CORREO:</B> " & sBusca & "  &nbsp;&nbsp;--  &nbsp;&nbsp;<B>MAIL : </B>" & xxx
        ELSE
       XXX=GETINFOUSER(sBusca)


       sGrpUsr = sImagenU & "&nbsp;&nbsp;" & sBusca & " - " & session("SGU_nombre")
       END IF

        sinput = sGrpUsr

    end if



        Celda = Celda & ofv.GenCelda("","ut style='border-bottom-style:dotted;border-bottom-width:1;border-bottom-color:#aaaaaa;' ","","","","","",sInput,"si")

        sInput = "&nbsp;"
        Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")

        Fila = ofv.genrow("","","","","","","",Celda,"si")

        Formulario = Formulario & fila

     rs.movenext
    loop

     ' sInput = "Lista de Usuarios"
      'Celda = ofv.GenCelda("","tt colspan=2 ","","","","","",sInput ,"si")

'response.write posi & "       " & senenviado
  if posi < 5 and senenviado <> "xxxxS" then

     sForm = "formu02"

      sAccion = "CoObjRecoADDUSR.asp?Opcion=ADDUSR&Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo & "&mensaje=" & msgcod
      if session("APLMDCLI") = "BGAL" THEN
      sInput = ofv.GenerarInput("","Agregar Recomendado","submit","","10","btsb","")
      else
      sInput = ofv.GenerarInput("","Agregar","submit","","10","btsb","")
      end if
      sInput = ofv.genform(sForm,"",sAccion,"","post","",sInput,"si")

  else

      sInput = "&nbsp;"

  end if

      Celda =  ofv.GenCelda("","ttt","LEFT","5%","","","",sInput,"si")
        sInput = "Lista de Usuarios"
      Celda = celdA & ofv.GenCelda("","tt colspan=2 ","","","","","",sInput ,"si")
      Fila = ofv.genrow("","","","","25","","",Celda,"si")

      Formulario = fila & Formulario


        sInput = "&nbsp;"
        Celda = ofv.GenCelda("","ut  ","","5%","","","",sInput,"si")
        Celda = Celda & ofv.GenCelda("","ut  ","","90%","","","",sInput,"si")
        Celda = Celda & ofv.GenCelda("","ut","","5%","","","",sInput,"si")
        Fila = ofv.genrow("","","","","","","",Celda,"si")
      '   if session("APLMDCLI") = "BGAL" THEN
      '    sInput = ofv.GenerarInput("","Enviar Recomendación","BUTTON","","10","btsb"," disp_confirm('Confirma el envio de esta recomendacion???','Formux2'); ")
      '    celda=ofv.GenCelda("","ut  colspan=3","","90%","","","",sInput,"si")
	'	  Fila = ofv.genrow("","","","","","","",celda,"si")
	'	 END IF

        Formulario = Formulario & fila
  sHTML = sHTML & ofv.GenTabla("","aaa brder style='border-style:solid;border-width:1;border-color:#990000;' ","","95% ","","0","","0","1","","",Formulario,"si")


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

    sHTML = sHTML & FormuX2  & FormuX2C


''''''''''''-----------------------------------------------------------------------

  sHTML = sHTML & "</center><br>"

'''''''''''''''''''''''''''''
'''''''
  sHTML = sHTML & "</div>"



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