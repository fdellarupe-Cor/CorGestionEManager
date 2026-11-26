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

 IF request.FORM.COUNT > 1 THEN
  SESSION(WSPSESION)=""
  SORLOTE= request.FORM("ORLOTE")
  'SORCUENTADESDE = request.FORM("ORCUENTADESDE")
  'SORCUENTAHASTA = request.FORM("ORCUENTAHASTA")
  FINICIO=REQUEST.FORM("FINICIO")
  FFIN=REQUEST.FORM("FFIN")
  SORCUENTA1=REQUEST.FORM("ORCUENTA1")
  SORCUENTA2=REQUEST.FORM("ORCUENTA2")
  SORCUENTA3=REQUEST.FORM("ORCUENTA3")
  SORESTADO=REQUEST.FORM("ORESTADO")
ELSE
  SORLOTE= request.QUERYSTRING("ORLOTE")
 ' SORCUENTADESDE = request.QUERYSTRING("ORCUENTADESDE")
 ' SORCUENTAHASTA = request.QUERYSTRING("ORCUENTAHASTA")
  FINICIO=REQUEST.QUERYSTRING("FINICIO")
  FFIN=REQUEST.QUERYSTRING("FFIN")
  SORCUENTA1=REQUEST.QUERYSTRING("ORCUENTA1")
  SORCUENTA2=REQUEST.QUERYSTRING("ORCUENTA2")
  SORCUENTA3=REQUEST.QUERYSTRING("ORCUENTA3")
  SORESTADO=REQUEST.QUERYSTRING("ORESTADO")
END IF


 IF SORLOTE="" THEN SORLOTE="TODOS"
' IF SORCUENTADESDE="" THEN SORCUENTADESDE="TODOS"
' IF sORCUENTAHASTA="" THEN sORCUENTAHASTA="TODOS"
 IF FINICIO="" THEN FINICIO="TODOS"
 IF FFIN="" THEN FFIN="TODOS"
 IF SORCUENTA1="" THEN SORCUENTA1="TODOS"
 IF SORCUENTA2="" THEN SORCUENTA2="TODOS"
 IF SORCUENTA3="" THEN SORCUENTA3="TODOS"
 IF SORESTADO=""  THEN SORESTADO="TODOS"





SCONECTOR=""

WKSETTABLA = "ORLOTEING"

  Strsql = "Select DISTINCT A.*  From " & WKSETTABLA & " A"
  STRSQL=STRSQL & " LEFT  JOIN ORLOTEDET B ON (A.codigo=B.LOTE )"

 IF SORLOTE <> "TODOS" THEN

	SORCODIGOW = " (A.CODIGO = 0" & SORLOTE & ") "

	SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
	SCONECTOR = " AND "

 END IF

   IF FINICIO <> "TODOS" THEN

  	       if IsDate(FINICIO) then
  	          WWSORFECINIH = CLNG(CDATE(FINICIO)) - CLNG(CDATE("01-01-1900"))

  	       else
  	          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
  	       end if

  	       SORFECINIHW   = " (B.FECVALOR > = " & WWSORFECINIH & ") "
             SCONECTOR=" AND "
  	       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW


      END IF

      IF FFIN <> "TODOS" THEN

  		       if IsDate(FFIN) then
  		          WWSORFECINIH = CLNG(CDATE(FFIN)) -  CLNG(CDATE("01-01-1900"))

  		       else
  		          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
  		       end if

  		       SORFECINIHW   = " (B.FECVALOR  < = " & WWSORFECINIH & ") "

  		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
  		       SCONECTOR=" AND "

       END IF



     sCUENTAS=""
     DELIM=""
     IF SORCUENTA1 <> "TODOS" THEN
      sCUENTAS="'" & SORCUENTA1 &"'" & DELIM
      DELIM=","
     END IF

     IF SORCUENTA2 <> "TODOS" THEN
	       sCUENTAS= sCUENTAS &  DELIM & "'" & SORCUENTA2 &"'"
	       DELIM=","
     END IF
     IF SORCUENTA3 <> "TODOS" THEN
	       sCUENTAS= sCUENTAS &  DELIM & "'" & SORCUENTA3 &"'"
	       DELIM=","
     END IF



      IF sCUENTAS <> "" THEN

  		       SORFECINIHW   = " (B.CUENTA  IN (" & sCUENTAS & ")) "

  		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
  		       SCONECTOR=" AND "

      END IF


      IF SORESTADO <> "TODOS" THEN



  		       SORFECINIHW   = " (B.ESTADO   = " & SORESTADO & ") "

  		       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
  		       SCONECTOR=" AND "

       END IF




 IF session(WSPSESION ) <> "" THEN
   swhere = " where "
 ELSE
  swhere = ""
END IF

  STRSQL =STRSQL & swhere & SESSION(WSPSESION)
  STRSQL =STRSQL & " ORDER BY A.OBJID DESC"

  sHTML = ofv.FormHeader(session("FormN"))



    sInput = "INVENTARIO DE LOTES DE INTERFASE CONTABLE"
    Celda = ofv.gencelda("","CC ","","90%","","","",sInput,"si")

    BTIT1 = "<A CLASS=CC HREF='ORINTERFASESMAPEX.asp' >MAPEO DE CUENTAS</A>"
    Celda = Celda & ofv.gencelda("","VT title='MAPEO DE CUENTAS.' STYLE='background-color:#000099;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","",BTIT1,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Fila
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"



  Formulario = ""





  'Encabezados
  Formulario = ofv.Encabezados("CODIGO","CREACION","PROCESO","REVERSION","ESTADO","&nbsp;","&nbsp;")

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

        strsqlX = "Insert into ORLOTEING "
        strsqlX = strsqlX & " (CODIGO,FECHACREATE,FECHAPROCESO,FECHAREVER,ESTADO"
        strsql = strsqlX & " ) Values ("
        strsqlX = strsqlX &  "'" & SESSION.SESSIONID & "','" &  sfechai & "','0','0','0'"
        strsqlX = strsqlX & ")"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSqlX,cn,1,volver)

        strsqlX = "UPDATE ORLOTEING SET CODIGO = OBJID WHERE CODIGO = '" & SESSION.SESSIONID & "' "
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSqlX,cn,1,volver)


  end if



  '
 'RESPONSE.WRITE STRSQL

  WKSETTABLA = "ORLOTEING"



  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0

 ' if not rs.eof Ilote=rs("lote")
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
    if clng(rs(4)) = 0 then
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

           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","10%","","","",sInput,"si")



       sAccion = "ORINTERFASESDETEX.asp?LOTE=" & RS(0) &"&PORIGEN=INTERFAZ"
       BTIT1 = "<INPUT TYPE='BUTTON' VALUE='DETALLE' CLASS='cc' title='Acceder al detalle del lote.' STYLE='cursor:pointer;background-color:#009900;color:#ffffff;' onclick='document.FCAN" & RS(0) & ".submit();'   > "
       Formu1 = ofv.GenForm("FCAN" & RS(0),"",sAccion,"","post","","","si")

       Celda = Celda & ofv.gencelda("","VT title='DETALLE DEL LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

    IF     wkestado = "1"  THEN

       STRSQL="SELECT  * FROM ORLOTEDET WHERE ESTADO=5 AND LOTE=0" & RS(0)

       set rseLIM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       IF  rseLIM.EOF THEN
       sAccion = "ORINTERFASESPROCEX.asp?LOTE=" & RS(0) &"&PORIGEN=INTERFAZ"
       BTIT1 = "<INPUT TYPE='BUTTON' VALUE='PROCESAR' CLASS='cc' title='Procesar el lote.' STYLE='cursor:pointer;background-color:#000099;color:#ffffff;' onclick=""disp_confirm('Confirma el procesamiento de este lote???','FINI" & posi & "');""  > "
       Formu2 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")
       ELSE
        BTIT1 = "<INPUT TYPE='BUTTON' VALUE='PROCESAR' CLASS='cc' title='selecionados.' STYLE='cursor:pointer;background-color:#FF0000;color:#ffffff;' onclick=""disp_confirm2('QUEDAN REGISTROS SELECIONADOS SIN CONFIRMAR');""  > "
        Formu2 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")
       END IF

       CALL OFV.CERRARCONSULTA(RSELIM)
       Celda = Celda & ofv.gencelda("","VT title='PROCESAR LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

    ELSEIF     wkestado = "2" THEN

       'sAccion = "ORINTERFASESPROCEX.asp?LOTE=" & RS(0) & "&OPCION=REV" & "&ORLOTE="& SORLOTE & "&ORCUENTADESDE=" & SORCUENTADESDE &"&ORCUENTAHASTA=" & SORCUENTAHASTA & "&FINICIO=" & FINICIO & "&FFIN=" & FFIN
       sAccion = "ORINTERFASESPROCEX.asp?&OPCION=REV&LOTE=" & RS(0) &"&PORIGEN=INTERFAZ"
       BTIT1 = "<INPUT TYPE='BUTTON' VALUE='REVERTIR' CLASS='BTr' title='REVERTIR el lote.' STYLE='cursor:pointer;background-color:#990000;' onclick=""disp_confirm('Confirma la reversion de este lote???','FINI" & posi & "');""  > "
       Formu2 = ofv.GenForm("FINI" & posi,"",sAccion,"","post","","","si")


       Celda = Celda & ofv.gencelda("","VT title='PROCESAR LOTE.' STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")

    ELSE
       Formu2 = ""
       SINPUT = "&nbsp;"
       Celda = Celda & ofv.gencelda("","UTI","","5%","","","",sInput,"si")

    END IF



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
    sAccion = "ORINTERFAZEX.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select CODIGO From " & WKSETTABLA & "  Order By CODIGO "    ' orden pagina
  str2 = "Select CODIGO From " & WKSETTABLA & " Order By CODIGO "    ' orden alfabetico


  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2("") )


  Celda =""

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



	sHTML = sHTML & "function disp_confirm2(disp_msg)"
	sHTML = sHTML & "  {"
	sHTML = sHTML & "    alert(disp_msg);"

    sHTML = sHTML & "  }"



''    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

function VERARCHIVOS()

  XNVER = "N"
  XNSEC = 0
  DIM STRSQL

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

          sfechai = clng(Cdate(WTABFILE(1))) - clng(cdate("01-01-1900"))



          strsql = "Insert into ORLOTEING "
          strsql = strsql & " (CODIGO,FECHACREATE,FECHAPROCESO,FECHAREVER,ESTADO"
          strsql = strsql & " ) Values ("
          strsql = strsql &  "'" & SESSION.SESSIONID & "','" &  sfechai & "','0','0','0'"
          strsql = strsql & ")"
 '''    response.write strsql & "<br>"
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
              WSIDEMPRESA=MID(WTABREGISTRO(YY),1,2)
              WSPFECCONT = CLNG(CDATE(MID(WTABREGISTRO(YY),9,2) & "-" & MID(WTABREGISTRO(YY),7,2) & "-" & MID(WTABREGISTRO(YY),3,4)  )) - clng(cdate("01-01-1900"))
              WSPFECVAL  = CLNG(CDATE(MID(WTABREGISTRO(YY),17,2) & "-" & MID(WTABREGISTRO(YY),15,2) & "-" & MID(WTABREGISTRO(YY),11,4) )) - clng(cdate("01-01-1900"))
              WSPNROLOTE  = MID(WTABREGISTRO(YY),19,5)
              WSPCTACONT  = MID(WTABREGISTRO(YY),24,12)
              WSPSUCURSAL  = MID(WTABREGISTRO(YY),36,5)
              WSPDEBCRE  = MID(WTABREGISTRO(YY),41,1)
              WSPMONTO  = MID(WTABREGISTRO(YY),42,13) & "." & MID(WTABREGISTRO(YY),55,2)
              WSPTEXTO  = MID(WTABREGISTRO(YY),57,30)
              WSPCLIENTE  = MID(WTABREGISTRO(YY),87,10)
              WSPSUMOTRE  = MID(WTABREGISTRO(YY),97,8)
              WSPEVENTO  = MID(WTABREGISTRO(YY),105,3)
              WSPPRODUCTO  = MID(WTABREGISTRO(YY),108,3)
              WSPCARACTER  = MID(WTABREGISTRO(YY),111,2)
              'NUEVOS
              WSPPROCESO    =  MID(WTABREGISTRO(YY),113,5)
              WSPUBPROCESO    = MID(WTABREGISTRO(YY),118,6)
              WSSPAREASEC    = MID(WTABREGISTRO(YY),124,20)
              WSSPAREASPORTE    =MID(WTABREGISTRO(YY),144,6)
              WSSPLINEANEGOCIO    = MID(WTABREGISTRO(YY),150,6)
              WSSUSUARIO		= MID(WTABREGISTRO(YY),156,30)


              '-----------------------------------------
              if not IsNumeric( MID(WTABREGISTRO(YY),113,5)) then WSPPROCESO=0
			  if not IsNumeric(MID(WTABREGISTRO(YY),118,6)) then WSPUBPROCESO=0
			  if not IsNumeric(MID(WTABREGISTRO(YY),124,20)) then WSSPAREASEC=0
			  if not IsNumeric(MID(WTABREGISTRO(YY),144,6)) then WSSPAREASPORTE=0
              if not IsNumeric(MID(WTABREGISTRO(YY),150,6)) then WSSPLINEANEGOCIO=0
              '-----------------------------------------


             sfechaL = clng(date()) - clng(cdate("01-01-1900"))

              strsql = "Insert into ORLOTEDET "
              strsql = strsql & " ("
              strsql = strsql & "LOTE,FECCONT,FECVALOR,CUENTA,SUCURSAL,"
              strsql = strsql & "MONTO,TEXTO,CLIENTE,FECPROC,ESTADO,"
              strsql = strsql & "INFDATA,INFEVENTO,INFPRODUCTO,INFCARACTER,INFTIPO,PROCESO,SUBPROCESO,AREASECTOR,AREASOPORTE,LINEANEGOCIOS,NROLOTE,USUARIO,IDEMPRESA,DEBCRED"
              strsql = strsql & " ) Values ("
              strsql = strsql &  "'" & SLOTE & "','" & WSPFECCONT & "','" & WSPFECVAL & "','" & WSPCTACONT & "','" & WSPSUCURSAL & "',"
              strsql = strsql &  "'" & WSPMONTO & "','" & WSPTEXTO & "','" & WSPCLIENTE & "','0','0',"
              strsql = strsql &  "'" & WSPSUMOTRE & "','" & WSPEVENTO & "','" & WSPPRODUCTO & "','" & WSPCARACTER & "','2','" & WSPPROCESO & "','"
              strsql = strsql  & WSPUBPROCESO &"','" & WSSPAREASEC &"','" &  WSSPAREASPORTE &"','"  & WSSPLINEANEGOCIO &"','"
              strsql = strsql  & WSPNROLOTE &"','" & WSSUSUARIO &"','" &  WSIDEMPRESA &"','" & WSPDEBCRE &"'"
              strsql = strsql & ")"
              CALL ofv.crearconsultaEx(StrSql,cn,1,volver)
               'RESPONSE.WRITE STRSQL

            '  STRSQL="PRC_EXECSQL '" & REPLACE(STRSQL,"'","''") & "'"
             ' STRSQL="PRC_EXECSQL '" & STRSQL & "'"
            '   response.write strsql

			'   SET VV= ofv.crearconsultaEx(StrSql,cn,1,volver)


		'	  set errores = cn.errors
		'	  if cn.errors.count = 0 then
		'	    IF NOT ISNULL(VV(0)) THEN
		'	     IDINSERT=VV(0)
	     '     END IF
	     '     END IF
         '     call ofv.cerrarconsulta(vv)
         '     strsql="update ORLOTEDET set "
         '     strsql= strsql & "  SUBPROCESO='" & TRIM(WSPUBPROCESO&"") &"' ,"
         '     strsql= strsql & "  AREASECTOR='" & TRIM(WSSPAREASEC) &"' ,"
         '     strsql= strsql & "  AREASOPORTE='" & TRIM(WSSPAREASPORTE &"' ,"
         '     strsql= strsql & "  LINEANEGOCIOS='" & TRIM(WSSPLINEANEGOCIO &"'"
         '     strsql= strsql & "  WHERE OBJID=" & IDINSERT
         '     RESPONSE.WRITE STRSQL & "<BR>"
         '     strsql="update ORLOTEDET set "
         '     strsql= strsql & "  NROLOTE='" & WSPNROLOTE &"' ,"
         '     strsql= strsql & "  USUARIO='" & WSSUSUARIO &"' ,"
         '     strsql= strsql & "  IDEMPRESA='" & WSIDEMPRESA &"' ,"
         '     strsql= strsql & "  DEBCRED='" & WSPDEBCRE &"'"
         '     strsql= strsql & "  WHERE OBJID=" & IDINSERT
         '     RESPONSE.WRITE STRSQL



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