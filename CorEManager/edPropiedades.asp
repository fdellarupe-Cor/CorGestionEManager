<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  aplcheck()

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
'response.write Session("Form2")
  bModificaresp = ofv.AModificar
  if session("Nomodificar") = "S" then

   ''  if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False

        bModificar = False

  end if
'bModificar = False
'response.write bModificar & "<br>"
  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")

  sHTML = ofv.FormHeader(session("FormN"))
 ' shtml= shtml &  "<script src='incfile/mascara.js' type='text/javascript' language='JavaScript'></script>" & chr(13)
  'shtml= shtml &  "<script type='text/javascript' src='incfile/masked_input_1.4.js' ></script>" & chr(13)


 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


 set cn = ofv.conectar(ofv.strconn2)
    strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
   ' response.write strsql
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      bVersion = "NO"
      xestado=""
      if not rs.eof then
         xestado=cstr(rs(1)&"")
        if rs(0) = "V" then bVersion = "SI"
      end if
  call ofv.cerrarconsulta(rs)

'response.write "estado=" & xestado & "   Nomodificar=" & session("Nomodificar") & "      SegRst=" &Session("SegRst") & " -- Segmo=" & Session("Segmo") & " -- bModificar" & bModificar & " antes "
'response.write bModificar & "<br>"
  if xestado="3" and session("Nomodificar") = "S" and (Session("SegRst") = "N" or (Session("SegRst") = "S" and Session("Segmo") = "S")) then
    if session("APLMDCLI") <> "BC" then
     ofv.ObtenerAtributos "TEMAS",""
     bModificar = ofv.AModificar
    end if
  end if
'response.write bModificar & "<br>"
'response.write Session("SegRst") & " -- " & Session("Segmo") & " -- " & bModificar

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"


    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs

  if  session("Error") <> "" then
      sHTML = sHTML & session("Error")
      session("Error") = ""
  end if

  sImagenI3 = "<B><FONT FACE=Corflaticon COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
  sImagenI4 = "<B><i class='las la-calendar-alt' style=font-size:14pt;COLOR:" & Session("CMFPNBOT") & "></i></B>"

  Fila = ""

  'set cn = ofv.conectar(ofv.strconn2)
  set cn2 = ofv.conectar(ofv.strconn0)


  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

  Strsql = "SELECT OBJPRTY.objid,OBJPRTY.codobjeto,OBJPRTY.tipobjeto,OBJPRTY.origen,OBJPRTY.instrumen,OBJPRTY.instid,OBJPRTY.fecemis,"
  Strsql = Strsql & " OBJPRTY.fecvigen,OBJPRTY.fecbaja,OBJPRTY.estado,OBJPRTY.autor,OBJPRTY.aprobo,OBJPRTY.elimino,OBJPRTY.fecvto,OBJPRTY.version,"
  Strsql = Strsql & " OBJPRTY.nrover,OBJPRTY.fecacta,OBJPRTY.nroacta,OBJPRTY.prtylink,OBJPRTY.prtylink2,OBJPRTY.fecpublic, Estados.estado FROM OBJPRTY, Estados "
  Strsql = Strsql & " where OBJPRTY.estado = Estados.objid and codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof Then

     sForm = "edPropiedades"

    'Estado
    Celda = ofv.GenCelda("","tt","","","","","","Estado","si")

    sInput = rs(21)
    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Version
    Celda = ofv.GenCelda("","tt","","20%","","","","Version","si")

    sSelect = ""
    if UCase(rs(14)) = "V" then sSelect = "selected"
    if bModificar then
       scomK  = "O" & chr(9) & "V"
       scomD  = "ORIGINAL" & chr(9) & "VERSION"
       sInput = ofv.gencomboFX(sForm,rs(14).name,rs(14),scomK,scomD," ")
    else
      if UCase(rs(14)) = "V" then
        sVersion = "VERSION"
      else
        sVersion = "ORIGINAL"
      end if
      sInput = sVersion
    end if
    Celda = Celda & ofv.GenCelda("","uti","","10%","","","",sInput,"si")
    if sSelect <> "" then
      if bModificar then
         if rs(15) = "" then
            xvernro = "1"
         else
            xvernro = rs(15)
         end if
        sInput = ofv.GenerarInput(rs(15).name,xvernro,"text",sForm,"3","dtn","")
      else
        sInput = "&nbsp;" & rs(15)
      end if
    else
        sInput = "&nbsp;"
    end if

    Celda = Celda & ofv.GenCelda("","ut","","70%","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Origenf
    Celda = ofv.GenCelda("","tt","","","","","","Origen","si")
    if bModificar then
      Strsql = "Select * From Origen Order By descri"

      scomKK = ""
	  scomDD = ""
  ''  response.write strsql & "<br>" & chr(13)
	  set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	  DO UNTIL rsniv.eof
	     scomKK = scomKK & rsniv(0) & chr(9)
	     scomDD = scomDD & rsniv(1) & chr(9)
	     RSNIV.MOVENEXT
	  LOOP

	  call ofv.cerrarconsulta(rsniv)

      sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3).value,scomKK,scomDD,"sb0")
 '''' sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,Strsql,cn,0,1,"")
    else
      if cstr(rs(3)) <> "" then
        Strsql = "Select descri From Origen Where objid = " & rs(3)
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
      end if
      sInput = sOrigen
    end if
    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Instrumento
    Celda = ofv.GenCelda("","tt","","","","","","Instrumento","si")
    if bModificar then
      Strsql = "Select * From Instrumen where estado= 3 Order By descri"

      scomKK = ""
	  scomDD = ""
  ''  response.write strsql & "<br>" & chr(13)
	  set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	  DO UNTIL rsniv.eof
	     scomKK = scomKK & rsniv(0) & chr(9)
	     scomDD = scomDD & rsniv(1) & chr(9)
	     RSNIV.MOVENEXT
	  LOOP
	  call ofv.cerrarconsulta(rsniv)

      sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4).value,scomKK,scomDD,"sb0")
  ''''sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,Strsql,cn,0,1,"")
    else
      if cstr(rs(4)) <> "" then
        Strsql = "Select descri From Instrumen Where objid = " & rs(4)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
      sInput = sInstrumento
    end if
    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    SMASINS = ""


    if  cstr(rs(4)) <> "0" AND cstr(rs(3)) <> "0"  and bModificar then

        sinput = "<a class=cc" & chr(34)
        sinput = sinput  & "underline" & chr(34) & "" & chr(34)
        sinput = sinput  & "" & chr(34) & ";'" & chr(34)
        sinput = sinput  & "" & chr(34) & ";" & chr(34)
        sinput = sinput  & "" & chr(34) & ";' "
        sinput = sinput  & "  href=EDPropiedadesMasIns.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
        sinput = sinput  & "  TITLE='Agregar origenes de documento adicionales.'  ><b>&nbsp;" & sImagenI3 & "...&nbsp;</b></a>"
        SMASINS = sinput
    END IF

    'Identificacion del Instrumento
    Celda = ofv.GenCelda("","tt","","","","","","Ident.&nbsp;Instrumento","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(5).name,rs(5),"textsb",sForm,"100","dt","")
    else
      sInput = ofv.GenerarInput("",rs(5),"readonly","","100","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput & SMASINS,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'link
    Celda = ofv.GenCelda("","tt","","","","","","Link Instrumento","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(18).name,rs(18),"textsb",sForm,"80","dt","")
    else
      sInput = rs(18)
    end if
    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'link
    Celda = ofv.GenCelda("","tt","","","","","","Link Analisis","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(19).name,rs(19),"textsb",sForm,"80","dt","")
    else
      sInput = rs(19)
    end if
    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Nro de acta de aprobacion
    Celda = ofv.GenCelda("","tt","","","","","","Nro. de Acta","si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(17).name,rs(17),"textsb",sForm,"20","dt","")
    else
      sInput = rs(17)
    end if

    Celda = Celda & ofv.GenCelda("","uti colspan=2","","","","","",sInput,"si")
    Filax = Fila &  ofv.GenRow("","","","","25","","",Celda,"si")


    'Sub Fechas......
    Celda = ofv.GenCelda("","ttt colspan=3","","","","","","Fechas","si")
   ' Fila = Fila &  ofv.GenRow("","","","","25","","",Celda,"si")
    Fila =  ofv.GenRow("","","","","25","","",Celda,"si")


    'Fecha de emision
    Celda = ofv.GenCelda("","tt","","","","","","Emision","si")

    if clng(rs(6)) = 0 then
      dFecha = ""
      dFechan = "-------"
    else

      dFechan = GTFechador(rs(6))


    end if


    sInput = dFechan
    Celda = Celda & ofv.GenCelda("","uTi colspan=2","","","","","",sInput ,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Fecha de vencimiento
    Celda = ofv.GenCelda("","tt","","","","","","Revision","si")

    sInputD = ofv.GenerarInput(rs(13).name,rs(13).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(13)) = 0 then
      dFecha = ""
      xfecact = "N"
      dFechan = "-------"
    else
      dFecha = cdate(rs(13) + clng(cdate("01-01-1900")))
      xfecact = rs(13)
   '   dFechan = Fechador(xfecact)
      dFechan = GTFechador(rs(13))

    end if

    if bModificar then

       sinput = "<a class=cc" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";' " & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";' "
       sinput = sinput  & "  href=EDPropiedadesfct.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo  & "&xfec=13&xfecact=" & xfecact & "  "
       sinput = sinput  & "  TITLE='Modificar fecha de vencimiento.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"

       sInput = dFechan & "&nbsp;" & sinput

       sinput = "<a class=cc" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";'" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";' "
       sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(13).name,GTCSMT,GTCFCH,GTCFCHDAT,"",4,2,0,0,0)
       sinput = sinput  & "  TITLE='Modificar fecha de vencimiento.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"

       sInput = dFechan & "&nbsp;" & sinput  & sInputD

    else
      sInput = dFechan
    end if
    Celda = Celda & ofv.GenCelda("","uTi colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Fecha de vigencia
    Celda = ofv.GenCelda("","tt","","","","","","Vigencia","si")

    sInputD = ofv.GenerarInput(rs(7).name,rs(7).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(7)) = 0 then
      dFecha = ""
      xfecact = "N"
      dFechan = "-------"
    else
      dFechan = GTFechador(rs(7))
    end if

    if bModificar  then

       sinput = "<a class=cc" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";'" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";' "
       if session("base")="K"  then
        sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(7).name,GTCSMT,GTCFCH,GTCFCHDAT,"",21,2,0,0,0)
       else
        select case session("APLMDCLI")
         case "GP"
           xIANOS=20
         case else
           xIANOS=10
        end select

       sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(7).name,GTCSMT,GTCFCH,GTCFCHDAT,"",xIANOS,2,0,0,0)
       end if
       sinput = sinput  & "  TITLE='Modificar fecha de vigencia.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"
       sInput = dFechan & "&nbsp;" & sinput  & sInputD
    else
       sInput = dFechan
    end if
    Celda = Celda & ofv.GenCelda("","uTi colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Fecha de publicacion
    Celda = ofv.GenCelda("","tt","","","","","","Publicacion","si")

    if not isnull(rs(20))  then
     if clng(rs(20)) = 0 then
      dFecha = ""
      dFechan = "-------"
     else
      dFecha = cdate(rs(20) + clng(cdate("01-01-1900")))
    ''  dFechan = Fechador(rs(20))
      dFechan = GTFechador(rs(20))
     end if
    end if

    sInput = dFechan
    Celda = Celda & ofv.GenCelda("","uTi colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Fecha de baja
    Celda = ofv.GenCelda("","tt","","","","","","Baja","si")

    if clng(rs(8)) = 0 then
      dFecha = ""
      dFechan = "-------"
    else
      dFecha = cdate(rs(8) + clng(cdate("01-01-1900")))
    ''  dFechan = Fechador(rs(8))
      dFechan = GTFechador(rs(8))
    end if

    sInput = dFechan
    Celda = Celda & ofv.GenCelda("","uTi colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    'Fecha de acta
    Celda = ofv.GenCelda("","tt","","","","","","Acta","si")

    sInputD = ofv.GenerarInput(rs(16).name,rs(16).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(16)) = 0 then
      dFecha = ""
      dFechan = "-------"
      xfecact = "N"
    else

      dFecha = cdate(rs(16) + clng(cdate("01-01-1900")))
      xfecact = rs(16)
    '''  dFechan = Fechador(xfecact)
      dFechan = GTFechador(rs(16))
    end if

    if bModificar then

       sinput = "<a class=cc" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";'" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";' "
       sinput = sinput  & "  href=EDPropiedadesfct.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo  & "&xfec=16&xfecact=" & xfecact & "  "
       sinput = sinput  & "  TITLE='Modificar fecha de acta.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"
       sInput = dFechan & "&nbsp;" & sinput
       sinput = "<a class=cc" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";'" & chr(34)
       sinput = sinput  & "" & chr(34) & ";" & chr(34)
       sinput = sinput  & "" & chr(34) & ";' "
       sinput = sinput  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(16).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,0,0,0,0)
       sinput = sinput  & "  TITLE='Modificar fecha de vigencia.'  ><b>&nbsp;" & sImagenI4 & "&nbsp;</b></a>"
       sInput = dFechan & "&nbsp;" & sinput  & sInputD

    else
       sInput = dFechan
    end if


    Celda = Celda & ofv.GenCelda("","uTi colspan=2","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objprty&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    if session("base")="K"     then
    xtabla= ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila,"no")
     xTABLA="<TABLE  cellpadding=0  CELLSPACING=0 ><TR><TD width=50% ><br><table>" & FILAX &"</table></TD><TD width=50% ><table  >" & FILA & "</table></TD></TR></TABLE>"
    'xtabla= ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario & Formulario2,"no")
    Formulario =  ofv.GenForm(sForm,"",sAccion,"","post","",xtabla,"no")
    formulario = Formulario & "</table></form>"
     sHTML = sHTML & formulario  '& "<center>"
    else

    Formulario2 = ""
     xtabla= ofv.GenTabla("","aaa","","90%","","0","","0","0","","",filax & fila,"no")
    'xtabla= ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario & Formulario2,"no")
    Formulario =  ofv.GenForm(sForm,"",sAccion,"","post","",xtabla,"no")
    formulario = Formulario & "</table></form>"


    Formulario2 = ""

    sHTML = sHTML & formulario  '& "<center>"
    end if
   ' sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario & Formulario2,"si")
    sHTML = sHTML & "<br>"
  end if

  call ofv.cerrarconsulta(rs)


  'links asociados
  fila=""
  strsql = "Select * From ObjLNK Where Ltipo = 'L' and codobjetoL = " & nCodigo
''response.write strsql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,volver)

   if not rs1.eof then
    celda =   ofv.GenCelda("","ttt","","","","","","ESQUEMAS QUE REFERENCIAN EL DOCUMENTO","si")
    Fila  =   ofv.GenRow("","","","","","","",Celda,"si")
     do while not rs1.eof
      celda= ofv.GenCelda("","ut","","","","","",LinkObj(rs1(1),rs1(2)),"si")
      Fila =  fila & ofv.GenRow("","","","","20","","",Celda,"si")
      rs1.movenext
     loop
     tabla=ofv.GenTabla("panel","aaa","","90%","","0","","0","0","","",fila,"si")
     sHTML = sHTML & tabla & "</body></html>"
   else
    sHTML = sHTML & "</body></html>"
   end if
end if
if session("base")="K"   then

Strsql = "Select * From objpropiedades where codobjeto=" & nCodigo & " and tipobjeto='" &sTipo  & "'"
 set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


 if rs.eof then
  sql=" insert into objpropiedades (codobjeto,tipobjeto) values (" & nCodigo & " ,'" & sTipo &  "')"
  call ofv.crearconsultaEx(sql,cn,1,parametros)
  Strsql = "Select * From objpropiedades where codobjeto=" & nCodigo & " and tipobjeto='" &sTipo  & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 end if

'response.write strsql

 sForm = sForm  & "prop"



   Celda = ofv.GenCelda("","tt","","20%","","","","TIEMPO GUARDA","si")
  '  if bModificar then
      Strsql = "Select * From FORMADMTGUARDA where estado=3  Order By descriPCION"
      sInput = ofv.GenerarCombo(sForm ,rs(5).name,rs(5).value,Strsql,cn,0,1,"")
 '   else
 '   sInput="222222"
 '   end if

   Celda = Celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
    Celda = ofv.GenCelda("","tt","","","","","","CATEGORIA","si")
	'    if bModificar then
	      Strsql = "Select * From FORMADMCATG  where estado=3 Order By descriPCION"
	      sInput = ofv.GenerarCombo(sForm ,rs(7).name,rs(7).value,Strsql,cn,0,1,"")

    '    end if

    Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



    Celda = ofv.GenCelda("","tt","","","","","","TIPO DE FORMULARIO","si")
	 '   if bModificar then
	      Strsql = "Select * From FORMADMTFORMU where estado=3 Order By descriPCION"
	      sInput = ofv.GenerarCombo(sForm ,rs(8).name,rs(8).value,Strsql,cn,0,2,"")

   ' end if

    Celda = Celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


     Celda = ofv.GenCelda("","tt","","","","","","CODIGO FORMULARIO","si")
		'
     'sInput = ofv.GenerarInput(rs(9).name,rs(9),"text",sForm,"40","dt","")
 		sinput = rs(9)
    ' sInput=replace(sInput,"TEXT","TEXT onkeyup=mascara(this,'-',patron,true)")
    ' sInput=replace(sInput,"TEXT","TEXT xonkeyup=mascara2(this,8,event)")




	     '    end if

	Celda = Celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


       Celda = ofv.GenCelda("","tt","","","","","","FIRMAS FORMULARIO","si")
'    if bModificar then
	  Strsql = "Select * From FORMADMTFIRMA where estado=3 Order By descriPCION"
	  sInput = ofv.GenerarCombo(sForm ,rs(10).name,rs(10).value,Strsql,cn,0,1,"")

	  '  end if

	   Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
       Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

        Celda = ofv.GenCelda("","tt","","","","","","BANCA","si")
	   		'    if bModificar then
	   Strsql = "Select * From FORMADMBANCA where estado=3 Order By descriPCION"
	   sInput = ofv.GenerarCombo(sForm ,rs(3).name,rs(3).value,Strsql,cn,0,2,"")

	   	  '  end if

	   Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
       Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



        Celda = ofv.GenCelda("","tt","","","","","","PROCESO","si")
	   		'    if bModificar then
	    Strsql = "Select * From FORMADMPROCESO where estado=3 Order By descriPCION"
	    sInput = ofv.GenerarCombo(sForm ,rs(13).name,rs(13).value,Strsql,cn,0,2,"")

	   	  '  end if

	   	Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
        Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



        Celda = ofv.GenCelda("","tt","","","","","","MODO IMPRESION","si")
        scomK = "SF" & chr(9) & "DF"
	    scomD = "SIMPLE FAZ" & chr(9) & "DOBLE FAZ"



        sInput = ofv.gencomboFX(sForm,"MODOIMPRESION",rs(11),scomK,scomD," ")

        Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
        Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

        Celda = ofv.GenCelda("","tt","","","","","","IMPRESION","si")
		scomK = "SI" & chr(9) & "NO"
	    scomD = "SI" & chr(9) & "NO"



		sInput = ofv.gencomboFX(sForm,"IMPRESION",rs(12),scomK,scomD," ")

		Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
        Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


          Celda = ofv.GenCelda("","tt","","","","","","CANT.COPIAS ","si")
         Strsql = "Select * From FORMADMTCOPIAS where estado=3 Order By descriPCION"
		  sInput = ofv.GenerarCombo(sForm ,"copias",rs("copias").value,Strsql,cn,0,1,"")


		Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
		Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

		sInput = "<textarea class=dt name=DESCRIPCION id=DESCRIPCION rows=4 cols=140 "
		   if bModificar  then
			  sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
			  sInput = sInput & " extenso'); return false;} else {"
			  sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
		   else
			  sInput = sInput & " readonly "
		   end if
        sInput = sInput & " >" &  trim(rs("DESCRIPCION")&"")  &"</textarea></font>"
        Celda = ofv.GenCelda("","tt","","","","","","DESCRIPCION","si")
        Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
		Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

      sAccion = ""
	    if bModificar then
	      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=objpropiedades&volver=" & vuelta
	      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	    end if
    fila =  ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila ,"si")

    Formulario = ofv.GenForm(sForm ,"",sAccion,"","post","",Fila,"si")
  '  sHTML = sHTML & "<center>"
    ' sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario ,"si") '& "</center>"
     sHTML = sHTML & formulario
end if

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if




      sHTML = sHTML & GTCCALENDAR("D")


'''''''
  sHTML = sHTML & "</span></div>"



  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs1)
  call ofv.cerrarconn(cn)
  response.write sHTML

%>

<%



 function LinkObj(id,tipo)
 dim rsx


  Strsql = "Select * From csobjetosRDO where codobjeto = " & id

''response.write strsql & "<br>"
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
%>