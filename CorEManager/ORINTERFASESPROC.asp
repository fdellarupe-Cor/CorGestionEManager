<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Server.ScriptTimeout=360
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
  sLOTE  = request.querystring("LOTE") & ""

  APLCHECK()
  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  DIM TABMAPA(200,5)


  Strsql = "Select * From ORLOTEMAP ORDER BY CUENTA DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABMAPA(POSI,0) = RS(0) 'CODIGO
          TABMAPA(POSI,1) = RS(1) 'CUENTA
          TABMAPA(POSI,2) = RS(2) 'CARACTER
          TABMAPA(POSI,3) = RS(3) 'EVENTO1
          TABMAPA(POSI,4) = RS(4) 'EVENTO2
          TABMAPA(POSI,5) = RS(5) 'PRODUCTO

          TTABMAPA = POSI

        RS.MOVENEXT
     LOOP



  END IF

    DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))
    dFechaHOY = GTFechadorSh(DHOY)





'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO

  if sopcion = "SI" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORLOTEDET "
        strsql = strsql & " ("
        strsql = strsql & "LOTE,FECCONT,FECVALOR,CUENTA,SUCURSAL,"
        strsql = strsql & "MONTO,TEXTO,CLIENTE,FECPROC,ESTADO"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & SLOTE & "','0','0','','',"
        strsql = strsql &  "'0','DESCRIPCION DEL EVENTO','','0','0'"
        strsql = strsql & ")"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "UPDATE ORLOTEING SET ESTADO = '1' WHERE CODIGO = '" & SLOTE & "' "
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if



  Formulario = ofv.Encabezados("CODIGO","CREACION","PROCESO","REVERSION","ESTADO")

  WKSETTABLA = "ORLOTEING"

  Strsql = "Select * From " & WKSETTABLA & " WHERE CODIGO = '" & SLOTE & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    'CODIGO
    sInput = "LOTE:&nbsp;" & RS(1)
    Celda = ofv.gencelda("","CC COLSPAN=2 ","","5%","","","",sInput,"si")

    'FECCREATE
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECPROCESO
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECREVER
    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(4))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")


    CcONTABLE=RS("CCONTABLE")
    'CcONTABLE="N" 'RS("CCONTABLE")
   'ESTADO

    wkestado = rs(5) & ""
    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI","","10%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    Formulario = Formulario & Fila

   END IF
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"







  Formulario = ""

  WKSETTABLA = "ORLOTEDET"


'''''''''''''''''''''''''''''''''''''''
'' REVIERTO LOTE

  if sopcion = "REV" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))


    '    strsql = "DELETE FROM  ORTRACER   WHERE TRCLOTEING = '" & SLOTE & "' "
         strsql = "UPDATE ORTRACER  set trcestado='CAN' ,fecfinenr="& sfechai &"  WHERE TRCLOTEING = '" & SLOTE & "' "
'''  response.write strsql & "<br>" & CHR(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



        strsql = "UPDATE ORLOTEING SET ESTADO = '3', FECHAREVER = '" & DHOY & "'  WHERE CODIGO = '" & SLOTE & "' "
'''  response.write strsql & "<br>" & CHR(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    Celda = ""


    SINPUT = "LOTE " & SLOTE & " REVERTIDO -  "
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Filax  = ofv.genrow("","","","","25","","",Celda,"si")


  ELSE



  Strsql1 = "Select * From " & WKSETTABLA & " WHERE LOTE = '" & SLOTE & "' Order By CUENTA DESC"
'  response.write strsql1
  set rs = ofv.crearconsultaEx(StrSql1,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)



     sCODTRACER = GTkey(SLOTE) & RS(0)
     sFECREGIS  = DHOY

     WSPCARACTER = "0"
     WSPEVENTO1  = "0"
     WSPEVENTO2  = "0"
     WSPPRODUCTO = "0"
     WSPRASOC = "4"
     WSPDISTRIBUCION = "N"
     WSTRCPORCDISTRIB ="0"

     WSPINFOING = "INTERFASE"

     wkTIPO   = rs(15) & ""

     IF wkTIPO = "2" THEN

  ''''  INFO DE LA INTERFASE

        WSPINFOING = rs(11) & ""

  ''''  PRODUCTO

       strsql = "Select descri From ORTIPOPRODUCTO where OBJID = 0" & rs(13)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then
          WSPPRODUCTO = rs(13)
       ELSE
          WSPPRODUCTO = "0"

       END IF
       call ofv.cerrarconsulta(rsniv)

  ''''  CARACTER

       strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & RS(14)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then
          WSPCARACTER = RS(14)
       ELSE
          WSPCARACTER = "0"

       END IF
       call ofv.cerrarconsulta(rsniv)


  ''''  EVENTO

       Strsql = "Select DESCRI,TIPOEVENTO From  ORTIPOCATEGO   "
       Strsql = Strsql & " where OBJID = 0" & RS(12)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then

          WSPEVENTO1  = rsniv(1)
          WSPEVENTO2  = RS(12)

       ELSE

          WSPEVENTO1  = "0"
          WSPEVENTO2  = "0"

       END IF
       call ofv.cerrarconsulta(rsniv)


   ''''  PROCESO

		''  Strsql = "Select codobjeto From  OBJETOS   "
		''  Strsql = Strsql & " where codobjeto = 0" & RS("PROCESO")


     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select objid From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & RS("PROCESO")

     else

       Strsql = "Select CODOBJETO From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & RS("PROCESO")

     end if


		  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rsniv.eof then
			 WSPPROCESO  = rsniv(0)
		  ELSE
			 WSPPROCESO  = "0"
		  END IF
		  call ofv.cerrarconsulta(rsniv)


   ''''  SUBPROCESO

		''  Strsql = "Select codobjeto From  OBJETOS   "
		''  Strsql = Strsql & " where codobjeto = 0" & RS("SUBPROCESO")

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select objid From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & RS("SUBPROCESO")

     else

       Strsql = "Select CODOBJETO From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & RS("SUBPROCESO")

     end if



		  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rsniv.eof then
			 WSPSUBPROCESO  = rsniv(0)
		  ELSE
			 WSPSUBPROCESO  = "0"
		  END IF
		  call ofv.cerrarconsulta(rsniv)

   ''''  AREA/SECTOR


		 IF session("APLMDCLI")="BC" THEN
		  Strsql = "Select ECO From  ECO   "
		  Strsql = Strsql & " where ECOMAP  LIKE '" & RS("AREASECTOR") & " -%'"

		 ELSE

           Strsql = "Select ECO From  ECO   "
    	   Strsql = Strsql & " where ECO = 0" & RS("AREASECTOR")

		 END IF

		  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rsniv.eof then
			 WSPAREASECTOR  = rsniv(0)
		  ELSE
			 WSPAREASECTOR  = "0"
		  END IF
		  call ofv.cerrarconsulta(rsniv)



   ''''  AREA SOPORTE

		  Strsql = "Select ECO From  ECO   "
		  Strsql = Strsql & " where ECO = 0" & RS("AREASOPORTE")
		  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rsniv.eof then
			 WSPAREASOPORTE  = rsniv(0)
		  ELSE
			 WSPAREASOPORTE  = "0"
		  END IF
		  call ofv.cerrarconsulta(rsniv)


   ''''  LINEA DE NEGOCIOS

		  Strsql = "Select objid From  ORLINNEGO2   "
		  Strsql = Strsql & " where OBJID = 0" & RS("LINEANEGOCIOS")
		  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rsniv.eof then
			 WSPLINEANEGOCIOS  = rsniv(0)
			  ELSE
			 WSPLINEANEGOCIOS  = "0"
		  END IF
		  call ofv.cerrarconsulta(rsniv)


         IF session("APLMDCLI")="BC" THEN
          WSSASTCONTABLE=RS("TRCASIENTOCONT")
         end if

      if CcONTABLE="S" then

        FOR II=1 TO TTABMAPA

	               IF TABMAPA(II,1) = (rs(4) & "") THEN

	                   WSPCARACTER = TABMAPA(II,2)
	                   WSPEVENTO1  = TABMAPA(II,3)
	                   WSPEVENTO2  = TABMAPA(II,4)
	                   WSPPRODUCTO = TABMAPA(II,5)
	                   EXIT FOR
	               END IF


        NEXT
      end if




     ELSE



        FOR II=1 TO TTABMAPA
        '    resposne.write rs(4) & "=" & TABMAPA(II,1) =  (rs(4) & "")  & "<BR>"
            IF TABMAPA(II,1) = (rs(4) & "") THEN
               '  resposne.write rs(4) & "=QQ" & TABMAPA(II,1) =  (rs(4) & "")  & "<BR>"
                WSPCARACTER = TABMAPA(II,2)
                WSPEVENTO1  = TABMAPA(II,3)
                WSPEVENTO2  = TABMAPA(II,4)
                WSPPRODUCTO = TABMAPA(II,5)
                EXIT FOR
            END IF


        NEXT

     END IF

     'XX=0
     'IF  ISNUMERIC(rs(4)) THEN XX=RS(04)
     WSPCUENTA   = CDBL("0" & RS(4))
     WSPCLIENTE  = rs(8) & ""
     WSPFCONTA   = "0" & rs(2)
     WSPFVALOR   = "0" & rs(3)
     WSPMONTO    = "0" & REPLACE(RS(6) & "",",",".")
     WSPOBSERVA  = trim(RS(7) & "")
     IF LEN(WSPOBSERVA)=0 THEN WSPOBSERVA="SIN DESCRIPCION"

        strsql = "Insert into ORTRACER "
        strsql = strsql & " (CODTRACER,FECREGIS,FECORIG,FECIDENT, FECCONTA,"
        strsql = strsql & " TRCPROCESO, TRCPRODSER, TRCNEGOLINE, TRCIDENTACT, TRCEVENTO, "
        strsql = strsql & " TRCOBSERV, TRCIMPBRUTO, TRCIMPACRECEN, TRCIMPOTROS, TRCIMPRECUPERA, "
        strsql = strsql & " TRCIMPNETO, TRCIMPSEGURO,TRCESTADO, TRCUSRREG,"
        strsql = strsql & " fecfinal, trccaracter, trcproducto, trcriesgoasoc, trcsecsoporte,"
        strsql = strsql & " TRCEVENTORIESGO1, TRCEVENTORIESGO2, TRCLOTEING, TRCTIPOING, TRCINFOING,"
        strsql = strsql & " TRCCLIENTE, TRCCTACBLE,FECFINENR,TRCECORESP,TRCDISTRIB,TRCPORCDISTRIB,TRCIMPDESAFECTA,trcimpnoctble"

        IF session("APLMDCLI")="BC" THEN
		     strsql = strsql &  ",TRCASIENTOCONT"
        END IF

        strsql = strsql &  ",TRCMARCAACTP,TRCMARCASEND"

        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & sCODTRACER & "','" & sFECREGIS & "','" & WSPFVALOR & "',"
        strsql = strsql &  "'" & WSPFCONTA & "','" & WSPFCONTA & "',"
        strsql = strsql &  "'" & WSPPROCESO & "','" & WSPSUBPROCESO & "','"& WSPLINEANEGOCIOS & "','0','0','" & WSPOBSERVA & "','" & WSPMONTO & "','0','0','0','" & WSPMONTO & "','0','REG',"
        strsql = strsql & "'" & SESSION("USUARIO") & "',"
        strsql = strsql & "'" & WSPFCONTA & "','" & WSPCARACTER & "','" & WSPPRODUCTO & "','" & WSPRASOC & "','" & WSPAREASOPORTE &"',"
        strsql = strsql & "'" & WSPEVENTO1 & "','" & WSPEVENTO2 & "','" & SLOTE & "','2','" & WSPINFOING & "',"
        strsql = strsql & "'" & WSPCLIENTE & "','" & WSPCUENTA & "','0'"
        strsql = strsql &  ",'" & WSPAREASECTOR & "','" & WSPDISTRIBUCION & "','" & WSTRCPORCDISTRIB & "','0','0'"

        IF session("APLMDCLI")="BC" THEN
         strsql = strsql &  ",'" & WSSASTCONTABLE & "'"
        END IF

        strsql = strsql & ",'N','N')"
' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        strsql = "UPDATE ORLOTEDET SET ESTADO = '2', FECPROC = '" & DHOY & "'  WHERE OBJID = 0" & RS(0)
 ''' response.write strsql & "<br>" & CHR(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



    posi = posi + 1
    rs.Movenext
  Loop


        strsql = "UPDATE ORLOTEING SET ESTADO = '2', FECHAPROCESO = '" & DHOY & "'  WHERE CODIGO = '" & SLOTE & "' "
'''  response.write strsql & "<br>" & CHR(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    Celda = ""


    SINPUT = "LOTE " & SLOTE & " PROCESADO - TOTAL EVENTOS: " & POSI
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Filax  = ofv.genrow("","","","","25","","",Celda,"si")


  END IF



  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Filax,"si")

  'Botonera
  Agregar = ""
  Buscar = ""

  sAccion = "ORINTERFASES.asp"
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>