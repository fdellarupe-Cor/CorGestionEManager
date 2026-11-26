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
  SOBJID = request.querystring("OBJID")
  ID = nCodigo


  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID

  sHTML = ofv.FormHeader(session("FormN"))

''  sHTML = "<html><body>"

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


''  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"
 '' sHTML = sHTML & "</div>"

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

 sHTML = sHTML & "<div id=zglos class=''>"
  sHTML = sHTML & "<br>"

''''''''''''-----------------------------------------------------------------------

  set cn = ofv.conectar(ofv.strconn2)


  SOPCION =REQUEST.QUERYSTRING("OPCION") & ""


  IF SOPCION = "SI" THEN


     nFecini = 0

     if nFecini < clng(cdate("01-01-1900")) then
        nFecini = clng(cdate("01-01-1900"))
     else
        nFecini = 0
     end if

        xmsgusr = "FAVORITO"
        dFechai = clng(date()) - nFecini
        dFechae = 0



        strsql = "Insert into objmsg ("
        strsql = strsql & " fecvig,fecexp,dias,texto,"
        strsql = strsql & " importancia,usr,"

        strsql = strsql & " MSGTIPOMESS,MSGAPLIC,MSGOBJCOD,MSGOBJTIP"
        strsql = strsql & " ) "
        strsql = strsql & " values ("
        strsql = strsql & dFechai & "," & dFechae & ",1,'" & xmsgusr & "',"
        strsql = strsql & "1,'" & Session("Usuario") & "',"
        strsql = strsql & "'1','" & wkaplic & "','" & nCodigo & "','" & sTipo & "'"
        strsql = strsql & ")"

      '   response.write strsql

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


 ELSEIF SOPCION="ELIM" THEN
 strsql = "DELETE FROM  objmsg WHERE OBJID=" & SOBJID
 'response.write strsql

 call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  END IF







  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


strsql ="select * from objprty where codobjeto=" & nCodigo  &  " and tipobjeto='" & sTipo & "'"

   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   if rs ("estado")&"" <> "3" then

   sInput = "Solo puede agregar documentos en estado vigente a favoritos!!!"
         Celda = ofv.GenCelda("","uti colspan=4 ","center","100%","","","",sInput,"si")

         Fila = ofv.genrow("","","","","30","","",Celda,"si")

         Celda = ofv.GenCelda("","ut colspan=4 ","","","","","","","si")
         Fila  = fila & ofv.genrow("","","","","","","",Celda,"si")

        ' sInput = ofv.GenerarInput("","ELIMINAR","submit","","10","btsb","")
        ' Celda = ofv.GenCelda("","uti colspan=4","","10%","","","",sInput,"si")

      'Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")

   else

   call ofv.cerrarconsulta(rs)
  Strsql = "Select * From objmsg where MSGTIPOMESS = '1' AND usr = '" & SESSION("USUARIO") & "' AND MSGOBJCOD = '" & nCodigo & "' AND MSGOBJTIP = '" & sTipo & "' "
 '' RESPONSE.WRITE STRSQL & "<BR>"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  if not rs.eof  then
      sInput = "Desea Eliminar este documento de Favoritos? "
      Celda = ofv.GenCelda("","uti colspan=4 ","center","100%","","","",sInput,"si")

      Fila = ofv.genrow("","","","","30","","",Celda,"si")

      Celda = ofv.GenCelda("","ut colspan=4 ","","","","","","","si")
      Fila  = fila & ofv.genrow("","","","","","","",Celda,"si")

      sInput = ofv.GenerarInput("","ELIMINAR","submit","","10","bietx","")
      Celda = ofv.GenCelda("","uti colspan=4","center","10%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")
      sAccion = "CoObjFAVORITOS.asp?Opcion=ELIM&Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo & "&OBJID=" & RS(0)

    ''  Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


  else



      sInput = "Desea agregar este documento a favoritos? "
      Celda = ofv.GenCelda("","uti colspan=4 ","center","100%","","","",sInput,"si")

      Fila = ofv.genrow("","","","","30","","",Celda,"si")

      Celda = ofv.GenCelda("","ut colspan=4 ","","","","","","","si")
      Fila  = fila & ofv.genrow("","","","","","","",Celda,"si")

      sInput = ofv.GenerarInput("","Agregar","submit","","10","bietx","")
      Celda = ofv.GenCelda("","uti colspan=4","center","10%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")
      sAccion = "CoObjFAVORITOS.asp?Opcion=SI&Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo

    ''  Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


  end if
end if
  Formulario =  ofv.GenTabla("","aaa","","95% ","","0","","0","1","","",Fila,"si")
  sHTML = sHTML & ofv.genform(sForm,"",sAccion,"","post","",Formulario,"si") & "<br>"

 '''' sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;' ","","95% ","","0","","0","1","","",Formulario,"si") & "<br>"






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

  sHTML = sHTML '& "</center>"

'''''''''''''''''''''''''''''
'''''''
  sHTML = sHTML & "</div></div>"



  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


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