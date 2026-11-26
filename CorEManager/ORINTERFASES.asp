<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sopcion  = request.querystring("opcion") & ""

  sHTML = ofv.FormHeader(session("FormN"))

    APLCHECK()

    sInput = "INVENTARIO DE LOTES DE INTERFASE CONTABLE"
    Celda = ofv.gencelda("","CC ","","90%","","","",sInput,"si")

    BTIT1 = "<A CLASS=CC HREF='ORINTERFASESMAP.asp' >MAPEO DE CUENTAS</A>"
    Celda = Celda & ofv.gencelda("","VT title='MAPEO DE CUENTAS.' STYLE='background-color:#000099;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","",BTIT1,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Fila
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"



  Formulario = ""





  'Encabezados
  Formulario = ofv.Encabezados("CODIGO","CREACION","PROCESO","REVERSION","ESTADO","&nbsp;","&nbsp;","CON&nbsp;MAPEO")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



 XXX = VERARCHIVOS()

'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO

  if sopcion = "SI" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORLOTEING "
        strsql = strsql & " (CODIGO,FECHACREATE,FECHAPROCESO,FECHAREVER,ESTADO,Ccontable"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & SESSION.SESSIONID & "','" &  sfechai & "','0','0','0'"
        strsql = strsql & ",'N')"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "UPDATE ORLOTEING SET CODIGO = OBJID WHERE CODIGO = '" & SESSION.SESSIONID & "' "
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if



  WKSETTABLA = "ORLOTEING"

  Strsql = "Select * From " & WKSETTABLA & " Order By FECHACREATE DESC,codigo desc"
'  response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    wkestado = rs(5) & ""

    'Eliminar
    if bSuprimir AND wkestado = "0" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'CODIGO
    sInput = RS(1)
    Celda = Celda & ofv.gencelda("","CC STYLE='border-style:solid;border-width:1;border-color:#888888;' " ,"","5%","","","",sInput,"si")

    'FECCREATE
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","15%","","","",sInput,"si")

    'FECPROCESO
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","15%","","","",sInput,"si")

    'FECREVER

    STRSQL= "SELECT TRCESTADO FROM ORTRACER WHERE TRCLOTEING=" & RS("CODIGO") & " AND TRCESTADO <> 'REG'"
    'RESPONSE.WRITE STRSQL

	set rsEST = ofv.crearconsultaEx(StrSql,cn,1,parametros)


	IF rsEST.EOF THEN
	  BELIM=TRUE
	ELSE
	  BELIM=FALSE
	END IF

	call ofv.cerrarconsulta(rsEST)



    if clng(rs(4)) = 0  then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(4))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","15%","","","",sInput,"si")



   'ESTADO


    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           case "4"

                 wkestadoD = "<font color='#336699' >SIN DATOS</font>"


           case "5"

                 wkestadoD = "<font color='#990000' >CANCELADO</font>"



           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",sInput,"si")



       sAccion = "ORINTERFASESDET.asp?LOTE=" & RS(0)
       BTIT1 = "<INPUT TYPE='BUTTON' VALUE='DETALLE' CLASS='cc' title='Acceder al detalle del lote.' STYLE='cursor:pointer;background-color:#009900;color:#ffffff;' onclick='document.FCAN" & RS(0) & ".submit();'   > "
       Formu1 = ofv.GenForm("FCAN" & RS(0),"",sAccion,"","post","","","si")

       Celda = Celda & ofv.gencelda("","VT title='DETALLE DEL LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

    IF     wkestado = "1" THEN

       sAccion = "ORINTERFASESPROC.asp?LOTE=" & RS(0)
       BTIT1 = "<INPUT TYPE='BUTTON' VALUE='PROCESAR' CLASS='cc' title='Procesar el lote.' STYLE='cursor:pointer;background-color:#000099;color:#ffffff;' onclick=""disp_confirm('Confirma el procesamiento de este lote???','FINI" & posi & "');""  > "
       Formu2 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")


       Celda = Celda & ofv.gencelda("","VT title='PROCESAR LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

    ELSEIF     wkestado = "2"  AND BELIM THEN

       sAccion = "ORINTERFASESPROC.asp?LOTE=" & RS(0) & "&OPCION=REV"
       BTIT1 = "<INPUT TYPE='BUTTON' VALUE='REVERTIR' CLASS='BTr' title='REVERTIR el lote.' STYLE='cursor:pointer;background-color:#990000;' onclick=""disp_confirm('Confirma la reversion de este lote???','FINI" & posi & "');""  > "
       Formu2 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")


       Celda = Celda & ofv.gencelda("","VT title='PROCESAR LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

    ELSE
       Formu2 = ""
       SINPUT = "&nbsp;"
       Celda = Celda & ofv.gencelda("","UTI","","5%","","","",sInput,"si")

    END IF



       SCOMSEP = chr(9)
        scomK   = "S" &  SCOMSEP & "N"
        scomD   = "SI" & SCOMSEP & "NO"


       sInput = ofv.gencomboFX(sForm,"CCONTABLE",RS("CCONTABLE"),scomK,scomD,"")
        Celda = Celda & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si") & Formu1 & Formu2

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORINTERFASES.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select CODIGO From " & WKSETTABLA & " Order By FECHACREATE DESC "    ' orden pagina
  str2 = "Select CODIGO From " & WKSETTABLA & " Order By CODIGO "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  Celda = ""

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '''  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

''    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

function VERARCHIVOS()

  XNVER = "N"
  XNSEC = 0


  WSPARCHIVOS = getdata( "","RFILES","INTERFASES","")

  IF (WSPARCHIVOS & "") <> "" THEN

     TABARCHIVOS = SPLIT(WSPARCHIVOS,"|")

     FOR II=0 TO UBOUND(TABARCHIVOS) - 1
         WTABFILE = SPLIT(TABARCHIVOS(II),"###")
      '   RESPONSE.WRITE WTABFILE(0) & "<BR>"
      '   RESPONSE.WRITE WTABFILE(1) & "<BR>"

       WSPARCHN = WTABFILE(0)

       xxxfl= getdata( WSPARCHN,"FSP","INTERFASES","")

       IF (xxxfl & "") <> "" THEN

          ''RESPONSE.WRITE xxxfl & "<BR>" & CHR(13)

         WTABREGISTRO = SPLIT(xxxfl,CHR(9))


          SLOTE = 0

'''''''''''''''''''''''''''''''''''''''
''        INCORPORO LOTE NUEVO

        ' sfechai = clng(date()) - clng(cdate("01-01-1900"))
         ' response.write Cdate(WTABFILE(1))
          sfechai=WTABFILE(1)
          sfechai=FormatDateTime(cdate(sfechai), 2)
          sfechai=dateserial(year(sfechai),month(sfechai),day(sfechai))
          sfechai = clng(sfechai) - clng(cdate("01-01-1900"))



          strsql = "Insert into ORLOTEING "
          strsql = strsql & " (CODIGO,FECHACREATE,FECHAPROCESO,FECHAREVER,ESTADO,Ccontable"
          strsql = strsql & " ) Values ("
          strsql = strsql &  "'" & SESSION.SESSIONID & "','" &  sfechai & "','0','0','0','N'"
          strsql = strsql & ")"
           call ofv.crearconsultaEx(StrSql,cn,1,volver)


          strsql = "SELECT OBJID FROM ORLOTEING WHERE CODIGO = '" & SESSION.SESSIONID & "' "
 '''    response.write strsql & "<br>"
          set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rs1.eof then SLOTE = RS1(0)

          strsql = "UPDATE ORLOTEING SET CODIGO = OBJID, ESTADO = '1' WHERE CODIGO = '" & SESSION.SESSIONID & "' "
 '''    response.write strsql & "<br>"
          call ofv.crearconsultaEx(StrSql,cn,1,volver)


          FOR YY=0 TO UBOUND(WTABREGISTRO) - 1

      '       response.write MID(WTABREGISTRO(YY),7,2) & "-" & MID(WTABREGISTRO(YY),5,2) & "-" & MID(WTABREGISTRO(YY),1,4)  & "<br>"

              WSPFECCONT = CLNG(CDATE(MID(WTABREGISTRO(YY),7,2) & "-" & MID(WTABREGISTRO(YY),5,2) & "-" & MID(WTABREGISTRO(YY),1,4)  )) - clng(cdate("01-01-1900"))
              WSPFECVAL  = CLNG(CDATE(MID(WTABREGISTRO(YY),15,2) & "-" & MID(WTABREGISTRO(YY),13,2) & "-" & MID(WTABREGISTRO(YY),9,4) )) - clng(cdate("01-01-1900"))
              WSPCTACONT  = cdbl(MID(WTABREGISTRO(YY),17,16))
              WSPSUCURSAL  = MID(WTABREGISTRO(YY),33,3)
              WSPDEBCRE  = MID(WTABREGISTRO(YY),36,1)
              WSPMONTO  = MID(WTABREGISTRO(YY),37,13) & "." & MID(WTABREGISTRO(YY),50,2)
              WSPTEXTO  = MID(WTABREGISTRO(YY),52,65)
              WSPCLIENTE  = MID(WTABREGISTRO(YY),117,9)
              WSPSUMOTRE  = MID(WTABREGISTRO(YY),126,13)
              WSPEVENTO  = MID(WTABREGISTRO(YY),139,3)
              WSPPRODUCTO  = MID(WTABREGISTRO(YY),142,3)
              WSPCARACTER  = MID(WTABREGISTRO(YY),145,3)
              'NUEVOS
              WSPPROCESO    = MID(WTABREGISTRO(YY),148,6)
              WSPUBPROCESO    = MID(WTABREGISTRO(YY),154,6)
              WSSPAREASEC    = TRIM(MID(WTABREGISTRO(YY),160,6))
              WSSPAREASPORTE    = MID(WTABREGISTRO(YY),166,6)
              WSSPLINEANEGOCIO    = MID(WTABREGISTRO(YY),172,6)


              if  session("APLMDCLI")="BC" THEN
			     WSSASTCONTABLE   = MID(WTABREGISTRO(YY),178,10)
              end if


             sfechaL = clng(date()) - clng(cdate("01-01-1900"))

              strsql = "Insert into ORLOTEDET "
              strsql = strsql & " ("
              strsql = strsql & "LOTE,FECCONT,FECVALOR,CUENTA,SUCURSAL,"
              strsql = strsql & "MONTO,TEXTO,CLIENTE,FECPROC,ESTADO,"
              strsql = strsql & "INFDATA,INFEVENTO,INFPRODUCTO,INFCARACTER,INFTIPO,PROCESO,SUBPROCESO,AREASECTOR,AREASOPORTE,LINEANEGOCIOS"

              IF session("APLMDCLI")="BC" THEN
			     strsql = strsql & ",trcAsientoCont"
              END IF


              strsql = strsql & " ) Values ("
              strsql = strsql &  "'" & SLOTE & "','" & WSPFECCONT & "','" & WSPFECVAL & "','" & WSPCTACONT & "','" & WSPSUCURSAL & "',"
              strsql = strsql &  "'" & WSPMONTO & "','" & WSPTEXTO & "','" & WSPCLIENTE & "','0','0',"
              strsql = strsql &  "'" & WSPSUMOTRE & "','" & WSPEVENTO & "','" & WSPPRODUCTO & "','" & WSPCARACTER & "','2','" & WSPPROCESO & "','"
              strsql = strsql  & WSPUBPROCESO &"','" & WSSPAREASEC &"','" &  WSSPAREASPORTE &"','"  & WSSPLINEANEGOCIO &"'"

              IF session("APLMDCLI")="BC" THEN
			     strsql = strsql & ",'" & WSSASTCONTABLE &"'"
              END IF

              strsql = strsql & ")"
        'response.write strsql & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


              IF INSTR(WSPTEXTO & "","SIN DATOS PROCESADOS") > 0 THEN

                 strsql = "UPDATE ORLOTEING SET  ESTADO = '4' WHERE CODIGO = '" & SLOTE & "' "
 '''    response.write strsql & "<br>"
                 call ofv.crearconsultaEx(StrSql,cn,1,volver)

                 EXIT FOR
              END IF



          NEXT





      END IF

           XXXX = getdata( WTABFILE(0),"FMOV","INTERFASES","INTERFASES\PROCESADOS")


     NEXT




''''  RESPONSE.WRITE WSPARCHIVOS & "<BR>"

  END IF


end function

%>