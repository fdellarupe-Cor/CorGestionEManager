<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

''''  response.buffer = true

''' verifica si es del circuito metodo 2.8.2 o anterior

     sTipo   = request.querystring("Tipobjeto")
     nCodigo = request.querystring("Codobjeto")

  set cn2 = ofv.conectar(ofv.strconn2)


  Strsql = "SELECT * FROM OBJWFL  "
  Strsql = Strsql & " Where codobjeto=0" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs.eof then

   ''  response.write "v " & rs(6) & "<br>"
     if rs(6)  then
        sMenu   = request.querystring("Menu")
        sSector = request.querystring("Sector")
        response.redirect "edObjWfl.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo  & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=ESPWFL"
     end if


  end if
  call ofv.cerrarconsulta(rs)


  mstdsc = ""

  WKFAUTO = "WKF0001693621131313343163"
  WKFAPRO = "WKF0001693673726313343163"
  DESLEVL = 9
  SUPLEVL = 9


     IF session("TCliente") = "RSK" THEN

         Strsql = "Select * From ORRSKCONFIG where APLIC = '" & WKAPLIC & "' "
         set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
         if not rs.eof then

            WKFAUTO = RS(5)
            WKFAPRO = RS(6)

            DESLEVL = 9
            SUPLEVL = 8
         END IF


        mstdsc = ""


         Strsql = "SELECT MSGCHAT FROM OBJWFL  "
         Strsql = Strsql & " Where codobjeto=" & session("NCliente")
         Strsql = Strsql & " And tipobjeto = '" & session("TCliente") & "'"
         set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)

         if not rs.eof then
            mstdsc = RS(0)
         END IF

       strsql = "SELECT * FROM  ORRSKCIRCUITO "
       strsql = strsql & " Where OBJID = 0" & session("NCliente")
       set xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not xrs.eof then
          IF XRS("APLIC") = "OR" THEN
             TIPOOBJ = "CIRCUITO AUTOEVALUACION"
          ELSE
             TIPOOBJ = "PLAN DE CONTINGENCIAS"
          END IF

          Strsql = "Select * From objetos where codobjeto = 0" & XRS("CODOBJETO")
          set xrs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not xrs.eof then
             TIPOOBJ2 = XRS(1)
             TIPOOBJ3 = XRS(3)
          end if

          SESSION("NOMREQ") = TIPOOBJ & ":&nbsp;" & TIPOOBJ2  & "&nbsp;-&nbsp;" & TIPOOBJ3

       END IF
       call ofv.cerrarconsulta(xrs)

     ELSE


         Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
         set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
         if not rs.eof then


            WKFAUTO = RS(19)
            WKFAPRO = RS(20)

            DESLEVL = 9
            SUPLEVL = RS(18)
         END IF

         mstdsc = ""


         Strsql = "SELECT MSGCHAT FROM OBJWFL  "
         Strsql = Strsql & " Where codobjeto=" & session("NCliente")
         Strsql = Strsql & " And tipobjeto = '" & session("TCliente") & "'"
         set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)

         if not rs.eof then
            mstdsc = RS(0)
         END IF


         StrSql = "Select * From OBJETOS where CODOBJETO = " & session("NCliente")
         set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
         if not rs.eof then

            Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
            set rsX1 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
            if not rsX1.eof then
               TIPOOBJ = RSX1(1)
            END IF


           SESSION("NOMREQ") = TIPOOBJ & ":&nbsp;" & RS(1)  & "&nbsp;-&nbsp;" & RS(3)
         END IF




     END IF




 ''' response.write " msg:" & mstdsc & "-" & SESSION("NOMREQ") & "<br><br><br>"


  call ofv.cerrarconn(cn2)

'''''' fin verifica



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))


  set cn = ofv.conectar(ofv.strconn5)
  set cn3 = ofv.conectar(ofv.strconn2)
  set cn4 = ofv.conectar(ofv.strconn6)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
  response.write bModificar
  sopcion = request.querystring("opcion")
  act = request.querystring("act")


  sopx = sopcion

  sHTML  = ofv.FormHeader(replace(session("FormN"),"_"," "))

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------





  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' >"



  celdas = ofv.GenCelda("","TTT","","","","","","AUTORIZACION DE DOCUMENTOS","si")
  filas  = ofv.GenRow("","","","","25","","",celdas,"si")
  sHTML  = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si") & "<br>"

'''  RESPONSE.WRITE SHTML

  filas  = ""

 ''' SHTML = ""

 ''' response.flush




  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_OPCION=" & SOPCION & "_ACT=S"





 DIM ATABLA(200,5)

 DIM ATABSW(200,3)

    sImagenO = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=corflaticon COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0232;</FONT></B>"
    sImagenC = "<B><FONT FACE=corflaticon COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenP = "<B><FONT FACE=corflaticon COLOR=#336699 SIZE=3>2</FONT></B>"
    sImagenRU = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenRG = "<B><FONT FACE=corflaticon COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
    sImagenRS = "<B><FONT FACE=corflaticon COLOR=#0000aa SIZE=3>G</FONT></B>"
    sImagenFI = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>5</FONT></B>"
    sImagenFF = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenSE = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=3>G</FONT></B>"
    sImagenMA = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>F</FONT></B>"

    sImagenA = "<B><FONT FACE=corflaticon COLOR=#00aa00 SIZE=3>&#0252;</FONT></B>"
    sImagenR = "<B><FONT FACE=corflaticon COLOR=#cc0000 SIZE=3>&#0251;</FONT></B>"


    Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsx.eof then

       session("WKFPSC")    = RSX(9)
       SESSION("SEGCKT")    = RSX(5)
       SESSION("WKF")       = RSX(0)
       SESSION("ORGUSR")    = RSX(15)

''''''''''''''''''''
       set cn2 = ofv.conectar(ofv.strconn0)


''SESSION("WKFETPU") = "U"
''SESSION("WKFEUSR") = "COREMANAGER"
''SESSION("WKFEMAI") = "@"

       sInputx = "COR@MANAGER.COM"

     IF SESSION("WKFETPU") = "D" THEN               ' EMISOR DEFAULT (EL RESPONSABLE DEL PASO)

        xxxx = getInfoUserw(SESSION("ORGUSR"))
        if session("SGU_mail") <> "" then    SINPUTX = SESSION("SGU_mail")

        SINPUTX =  SINPUTX & ""

     ELSEIF SESSION("WKFETPU") = "U" THEN           ' EMISOR PREDETERMINADO (PARAMETRIZADO EL LA BASE DE USUARIOS)

        WKFEUSR = SESSION("WKFEUSR")

        if WKFEUSR = "" then   WKFEUSR = "COREMANAGER"


        strsql = "Select EMAIL From USUARIOS where ID = '" & WKFEUSR & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sInputx = rsniv(0)
        call ofv.cerrarconsulta(rsniv)

     ELSE                                           ' EMISOR FIJO (USUARIO Y MAIL FIJOS EN CORCONECT)

       sInputx = SESSION("WKFEMAI")


     END IF



       if SINPUTX = "" then SINPUTX = "COR@MANAGER.COM"

       SESSION("ORGMAIL") = SINPUTX








''''''''''''''''''''''''''''''''


      SESSION("CKTMODAL")  = ""
       Strsql = "SELECT CKTMODAL FROM SEGCIRCUITO WHERE SEGCKT = '" & SESSION("SEGCKT") & "' "
       set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rsx2.eof then
          SESSION("CKTMODAL")       = rsx2(0)
       END IF
       call ofv.cerrarconsulta(rsx2)

       IF SESSION("CKTMODAL")  = "" THEN  SESSION("CKTMODAL")  = "S"

    END IF
    call ofv.cerrarconsulta(rsx)

''''''''''''''''''''''''''''''''''''
'' busca originante y destinatario

    ATABLA(0,0) =  2
    ATABLA(0,1) =  0

   session("WKFCLASE") = "M"

   StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & SESSION("WKFPSC") & "' "
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rs.eof then
      session("WKFCLASE") = RS(3)
   END IF


   IF session("WKFCLASE") = "E" THEN


     StrSql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & SESSION("SEGCKT") & "' ORDER BY ORDEN, ID"
'''   response.write StrSql
     set rsxp1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsxp1.eof then

        XINDT = 0

        DO WHILE NOT RSXP1.EOF

           XINDT = XINDT + 1


           ATABLA(XINDT,0) =  rsxp1(3)
           ATABLA(XINDT,1) =  rsxp1(5)
           ATABLA(XINDT,2) =  rsxp1(4)
           ATABLA(XINDT,3) =  rsxp1(2)
           ATABLA(XINDT,4) =  rsxp1(7)
           ATABLA(XINDT,5) =  rsxp1(8)

        RSXP1.MOVENEXT
        LOOP

        ATABLA(0,0) =  XINDT
        ATABLA(0,1) =  XINDT - 1

     end if
     call ofv.cerrarconsulta(rsxP1)


   ELSE

   StrSql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND MODO = 'O'"
'''   response.write StrSql
   set rsxp1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rsxp1.eof then

      ATABLA(1,0) =  "O"
      ATABLA(1,1) =  rsxp1(5)
      ATABLA(1,2) =  rsxp1(4)
      ATABLA(1,3) =  "NO"
      ATABLA(1,4) =  rsxp1(7)
      ATABLA(1,5) =  rsxp1(8)

   end if
    call ofv.cerrarconsulta(rsxP1)

   StrSql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND MODO = 'D'"
'''   response.write StrSql
   set rsxp1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rsxp1.eof then

      ATABLA(2,0) =  "D"
      ATABLA(2,1) =  rsxp1(5)
      ATABLA(2,2) =  rsxp1(4)
      ATABLA(2,3) =  "NO"
      ATABLA(2,4) =  rsxp1(7)
      ATABLA(2,5) =  rsxp1(8)

   end if
    call ofv.cerrarconsulta(rsxP1)
''''''''''''''''''''''''''''''''''''

  END IF



 IF SOPCION = "OK" OR SOPCION = "RCH" THEN


	 IF ACT = "S" THEN
	    sobserv = request.form("OBS")

		  StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & SESSION("WKFPSC") & "' "
		  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		  if not rs.eof then

		     psctipo = RS(14)

		     IF SOPCION = "OK" THEN
			XTIPO = "APR"
			XPASCOD = RS(12)
		     ELSE
			XTIPO = "RCH"
			XPASCOD = RS(13)
		     END IF

		    SEGCKTPSC = GTCKTPSC()
		  '''  response.write segcktpsc





			IF PSCTIPO = "U" OR SOPCION = "RCH" THEN

				    WXX = GENACCION(RS(5),XTIPO)


				    StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "

				    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				    if not rs2.eof then

				       XMODEL  = RS2(5)
				       XPDSC   = RS2(6) & ":" & RS2(7)


				       IF RS(3) = "E" THEN

					  XRTIPO  = RS2(15)
					  XRCOD   = RS2(16)
					  XRASG   = "NO"

                                          XRLVL   = 9
                                          XRSCD   = "0"


                                          IF XRTIPO = "USR" OR XRTIPO = "RSP" THEN
                                             XRASG   = XRCOD
                                             IF XRTIPO = "RSP" THEN
                                                XRTIPO = "USR"
                                                FOR YYY=1 TO ATABLA(0,0)

                                                    IF ATABLA(YYY,0) =  XRCOD THEN
					               XRCOD   = ATABLA(YYY,1)
                                                       XRASG   = XRCOD
                                                       XRLVL   = ATABLA(YYY,4)
                                                       XRSCD   = ATABLA(YYY,5)
                                                       EXIT FOR
                                                    END IF

                                                NEXT

                                             END IF
                                          END IF

				       ELSE


					  IF SOPCION = "RCH" THEN
					     XRTIPO  = ATABLA(1,2)
					     XRCOD   = ATABLA(1,1)
					     XRASG   = ATABLA(1,1)
                                             XRLVL   = ATABLA(1,4)
                                             XRSCD   = ATABLA(1,5)
					  ELSE
				             XRTIPO  = ATABLA(2,2)
					     XRCOD   = ATABLA(2,1)
					     XRASG   = "NO"
					     IF XRTIPO  = "USR" THEN
					        XRASG   = XRCOD
					     END IF
                                             XRLVL   = ATABLA(2,4)
                                             XRSCD   = ATABLA(2,5)
					  END IF

				       END IF

				''' GENERAR EL PROXIMO PASO
				       WXX = GENCKTPSC(SEGCKTPSC,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD)

 '''''''''''''''''''
 ' agregar envio de mail
                                       XXXX = GENCKTMAIL(XRTIPO,XRCOD,mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),XRLVL)
 '''''''''''''''''''


				    END IF
				    call ofv.cerrarconsulta(rs2)


                                 if  SESSION("CKTMODAL") = "N" THEN

                                     Strsql = "SELECT SEGCKTPSC FROM SEGPASO WHERE SEGCKT = '" & SESSION("SEGCKT") & "'"
                                     Strsql = Strsql & " AND MODEL =  '" &  session("WKFPSC") & "' AND FIN = 'NO' ORDER BY ID "
                                     set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                     if NOT rsx2.eof then
                                        DO WHILE NOT RSX2.EOF
                                           IF RSX2(0) <> SESSION("SEGCKTPSC") THEN

			''''''''''''''''''''''''''''' ACTUALIZA PASO Y CIRCUITO DEL WORKFLOW CORRIENTE ''''''''''''''''''''''''''''''''''

			                      strsql = "UPDATE SEGPASO SET FIN = '" & GETDH() & "', OKRCH = '" & SEGCKTPSC
                                              Strsql = Strsql & "', observaciones = 'CANCELADO POR RECHAZO DEL CIRCUITO.'"
                                              Strsql = Strsql & " WHERE SEGCKTPSC = '" & RSX2(0) & "' "
			                      call ofv.crearconsultaEx(StrSql,cn,1,parametros)

                                           END IF
                                        RSX2.MOVENEXT
                                        LOOP
                                     END IF
                                     call ofv.cerrarconsulta(rsx2)

                                  end if



			ELSE





				    set cn2 = ofv.conectar(ofv.strconn0)

 				    sniv = 0
 				    'strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
			           ' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
				   ' if not rsniv.eof then sniv = rsniv(0)
				   ' call ofv.cerrarconsulta(rsniv)
				     xxxx = getInfoUserw(SESSION("USUARIO"))
				     if isnumeric(session("SGU_idcostcenter")) then sniv=session("SGU_idcostcenter")


				     ITAB  = 0
				     itabs = 0
				     StrSql = "Select * From SEGCKTRESP WHERE SEGCKTPSC = '" & SESSION("WKFPSC") & "' AND SEGCKT = '" & SESSION("SEGCKT") & "' order by ORDEN "
			  '''                 response.write StrSql
				     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				     if not rs3.eof then
					DO UNTIL RS3.EOF
					   ITAB = ITAB + 1
                                           ATABLA(0,0) =  ATABLA(0,0) + 1
                                           ATABLA(0,1) =  ATABLA(0,1) + 1
				 '''          response.write itab

                                           ATABLA(ITAB + 2,0) =  "L"
                                           ATABLA(ITAB + 2,1) =  RS3(5)
                                           ATABLA(ITAB + 2,2) =  RS3(4)
                                           XWTIPO             =  RS3(4)
                                           XWCOD              =  RS3(5)
                                           XWSEC              =  RS3(8)
                                           ATABLA(ITAB + 2,3) =  "NO"
                                           ATABLA(ITAB + 2,4) =  RS3(7)
                                           ATABLA(ITAB + 2,5) =  RS3(8)

					''   if (cstr(sniv) = cstr(rs3(5))) or (cstr(sniv) = cstr(rs3(8))) then
                                           SELECT CASE XWTIPO

                                                  CASE "USR"


                                                        if XWCOD = SESSION("USUARIO") THEN
					                   itabs = itab
					                end if


                                                  CASE "GRP"

                                                        if XWCOD = SESSION("SEGGRP") THEN
					                   itabs = itab
					                end if

                                                  CASE "SEC"

                                                        if CSTR(XWCOD) = CSTR(SNIV) THEN
					                   itabs = itab
					                end if

                                            END SELECT




					rs3.movenext
					LOOP
				     END IF

                                 xmodalsen = "S"
                                 if  SESSION("CKTMODAL") = "N" THEN
   ''    session("WKFPSC")    = RSX(9)
    ''   SESSION("SEGCKT")    = RSX(5)

                                     Strsql = "SELECT SEGCKTPSC FROM SEGPASO WHERE SEGCKT = '" & SESSION("SEGCKT") & "'"
                                     Strsql = Strsql & " AND MODEL =  '" &  session("WKFPSC") & "' AND FIN = 'NO' ORDER BY ID "
                                     set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                     if NOT rsx2.eof then
                                        DO WHILE NOT RSX2.EOF
                                           IF RSX2(0) <> SESSION("SEGCKTPSC") THEN
                                              xmodalsen = "N"
                                              EXIT DO
                                           END IF
                                        RSX2.MOVENEXT
                                        LOOP
                                     END IF
                                     call ofv.cerrarconsulta(rsx2)

                                 IF  xmodalsen = "S" THEN

				     itabs = itab

                                 ELSE

				     XMODEL  = RS(5)
				     XPDSC   = RS(6) & ":" & RS(7)

                                     XRTIPO  = ATABLA(itabs + 3,2)
			             XRCOD   = ATABLA(itabs + 3,1)
				     XRASG   = "NO"
                                     XRLVL   = ATABLA(itabs + 3,4)
                                     XRSCD   = ATABLA(itabs + 3,5)

                                 END IF





                                 END IF



                                 IF  xmodalsen = "S" THEN

				     if itabs = itab then

					WXX = GENACCION(RS(5),XTIPO)
					StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
					set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					if not rs2.eof then
					   XMODEL  = RS2(5)
					   XPDSC   = RS2(6) & ":" & RS2(7)


					   IF RS2(6) <> 999 THEN
					     XRTIPO  = ATABLA(1,2)
					     XRCOD   = ATABLA(1,1)
					     XRASG   = ATABLA(1,1)
                                             XRLVL   = ATABLA(1,4)
                                             XRSCD   = ATABLA(1,5)
					  ELSE
				             XRTIPO  = ATABLA(2,2)
					     XRCOD   = ATABLA(2,1)
					     XRASG   = "NO"
					     IF XRTIPO  = "USR" THEN
					        XRASG   = XRCOD
					     END IF
                                             XRLVL   = ATABLA(2,4)
                                             XRSCD   = ATABLA(2,5)
                                          end if



					END IF
					call ofv.cerrarconsulta(rs2)

				     elseif itabs < itab then
					 if itabs <> 0 then
					       XMODEL  = RS(5)
					       XPDSC   = RS(6) & ":" & RS(7)

				               XRTIPO  = ATABLA(itabs + 3,2)
					       XRCOD   = ATABLA(itabs + 3,1)
					       XRASG   = "NO"
                                               XRLVL   = ATABLA(itabs + 3,4)
                                               XRSCD   = ATABLA(itabs + 3,5)





					 else
					    WXX = GENACCION(RS(5),XTIPO)

					    StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
					    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					    if not rs2.eof then
					       XMODEL  = RS2(5)
					       XPDSC   = RS2(6) & ":" & RS2(7)


					       IF RS2(6) <> 999 THEN
					          XRTIPO  = ATABLA(1,2)
					          XRCOD   = ATABLA(1,1)
					          XRASG   = ATABLA(1,1)
                                                  XRLVL   = ATABLA(1,4)
                                                  XRSCD   = ATABLA(1,5)
					       ELSE
				                  XRTIPO  = ATABLA(2,2)
					          XRCOD   = ATABLA(2,1)
					          XRASG   = "NO"
					          IF XRTIPO  = "USR" THEN
					             XRASG   = XRCOD
					          END IF
                                                  XRLVL   = ATABLA(2,4)
                                                  XRSCD   = ATABLA(2,5)
                                               end if


					    END IF
					    call ofv.cerrarconsulta(rs2)
					 end if
				     end if






				''' GENERAR EL proximo PASO OK
				       WXX = GENCKTPSC(SEGCKTPSC,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD)

 '''''''''''''''''''
 ' agregar envio de mail
                                       XXXX = GENCKTMAIL(XRTIPO,XRCOD,mstdsc & "<br><br>" & XPDSC & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),XRLVL)
 '''''''''''''''''''

                                   END IF



			END IF

			''''''''''''''''''''''''''''' ACTUALIZA PASO Y CIRCUITO DEL WORKFLOW CORRIENTE ''''''''''''''''''''''''''''''''''

			    strsql = "UPDATE SEGPASO SET FIN = '" & GETDH() & "', OKRCH = '" & SEGCKTPSC & "', observaciones = '" & sobserv & "' WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
			    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			    IF RS(11) THEN

				     strsql = "UPDATE SEGCIRCUITO SET FIN = '" & GETDH() & "'  WHERE SEGCKT = '" & SESSION("SEGCKT") & "' "
				     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			    END IF


		  END IF




	 	  Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	 	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	  if NOT rsx.eof then


			  celdas = ofv.GenCelda("","TTr colspan=2","","","","","","WORKFLOW","si")
			  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","vt","","15%","","","","CIRCUITO:","si")
			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","vt","","15%","","","","PASO:","si")
			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","","<b>" & sImagenp & "</b> " & rsX(16),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","vt","","15%","","","","OPCION:","si")
			  IF SOPCION = "OK" THEN
			     SOPCDESC = " <b>" & sImagenA & "</b> " & "APRUEBA"
			  ELSE
			     SOPCDESC = " <b>" & sImagenR & "</b> " & "RECHAZA"
			  END IF

			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","",SOPCDESC,"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  IF RSX(11) <> "NO" THEN

			       celdas = ofv.GenCelda("","vt colspan=2 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ","CENTER","","","","","<B>PASO FINALIZADO:</B>","si")
			       celdas2 = ofv.GenCelda("","vt colspan=2","CENTER","","","","","<b>" & sImagenFF & "</b> " & RSX(11),"si")

			       filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			       Filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")


			  END IF



			  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"

	 	  END  IF






	 ELSE


	 	  Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	 	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	  if NOT rsx.eof then


			  celdas = ofv.GenCelda("","TTr colspan=2","","","","","","WORKFLOW","si")
			  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","vt","","15%","","","","CIRCUITO:","si")
			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","vt","","15%","","","","PASO:","si")
			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","","<b>" & sImagenP & "</b> " & rsX(16),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","vt","","15%","","","","OPCION:","si")
			  IF SOPCION = "OK" THEN
			     SOPCDESC = " <b>" & sImagenA & "</b> " & "APRUEBA"
			  ELSE
			     SOPCDESC = " <b>" & sImagenR & "</b> " & "RECHAZA"
			  END IF

			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","",SOPCDESC,"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


			  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"

	 	  END  IF


		    Celda = ofv.GenCelda("","vt colspan=3 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;' ","center","100%","","","","OBSERVACIONES/COMENTARIOS","si")
		    FilaS = Ofv.GenRow("","","","","","","",Celda,"si")

		    nCol = 80

		    wdescri = "SIN OBSERVACIONES"

		    t3b = "<font class=at><textarea class=dt "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=20 "
		    if bModificar then
		    '''  t3b = t3b & " onblur=Usuarios.submit() "
		      t3b = t3b & " onkeypress=" & chr(34)
		      t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
		      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
		    else
		      t3b = t3b & " readonly >"
		    end if
		    xdatos =  wdescri
		    t3b = t3b & xdatos & "</textarea></font>"
		    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","",t3b,"si")
		    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

                    sinput = "&nbsp;"
                    celda = ofv.GenCelda("","","center","","50%","","",sInput,"si")

                    sinput = "<input CLASS=BT type=submit value='CONTINUAR' size=10 >"

                    celda = celda & ofv.GenCelda("","","center","","1%","","",sInput,"si")

                    sinput = "&nbsp;"
                    celda = celda & ofv.GenCelda("","","center","","50%","","",sInput,"si")
		    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


		    if bModificar then
		      sAccion = "COWKFselect.ASP?CantRegMover=0&OPCION=" & SOPCION & "&ACT=S"
		      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
		    else
		      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
		    end if


		    sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",formulario0,"si") & "<BR><BR>"

	  END IF

 ELSE



	  Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
        ''  response.write strsql & "<br>" & chr(13)
	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsx.eof then


	  celdas = ofv.GenCelda("","TTr colspan=2","","","","","","WORKFLOW","si")
	  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","vt","","15%","","","","CIRCUITO:","si")
	  celdas = celdas & ofv.GenCelda("","vt","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","vt","","15%","","","","PASO:","si")
	  celdas = celdas & ofv.GenCelda("","vt","","80%","","","","<b>" & sImagenp & "</b> " & rsX(16),"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  sHTML = sHTML  & ofv.GenTabla("","AAA  ","","98%","","0","","0","0","","",filas,"si") & "<BR>"




	  Strsql = "SELECT * FROM CKTVIEW WHERE OKRCH = '" & session("SEGCKTPSC") & "' "
	  set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     ''     response.write strsql & "<br>" & chr(13)
	  if NOT rsxT.eof then

             PSCMODEL = RSXT(9)
             PSCMOT   = "A"

	     Strsql = "SELECT RCHPSC FROM WKFPASOS WHERE WKFPASCOD = '" & PSCMODEL & "' "
	     set rsxu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	     if NOT rsxU.eof then

                IF RSXU(0) = RSX(9) THEN PSCMOT   = "R"

             END IF

             SELECT CASE PSCMOT
                 CASE "A"
                      PSCMOT  = sImagenA & "&nbsp;POR APROBACION&nbsp;&nbsp;"
                 CASE "R"
                      PSCMOT  = sImagenR & "&nbsp;POR RECHAZO&nbsp;&nbsp;"
             END SELECT



	     celdas = ofv.GenCelda("","VT ","","50%","","","","VIENE DE <BIG><B>" & rsXT(16) & "</B></BIG>","si")
	     celdas = celdas & ofv.GenCelda("","VT ","RIGHT","50%","","","", PSCMOT & "<SMALL><SMALL>" & RSXT(11) & "</SMALL></SMALL>","si")
	     filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			set cn2 = ofv.conectar(ofv.strconn0)

			sniv  = "DESCONOCIDO"
                        SNSEC = "DESCONOCIDO"


			xxxx = getInfoUserw(rsXT(15))
			sniv = session("SGU_nombre")
			if isnumeric(session("SGU_idcostcenter")) then
			 SNSEC = GETRESPOCKT(session("SGU_idcostcenter"),"SEC")
			end if


	    celdas = ofv.GenCelda("","vt colspan=2","","","","","","RESPONSABLE:&nbsp;" & sImagenRU & "&nbsp;" & sniv,"si")
	    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
	    celdas = ofv.GenCelda("","vt colspan=2 ","","","","","","AREA/SECTOR:&nbsp;" & sImagenSE & "&nbsp;" & SNSEC,"si")
	    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	    celdas = ofv.GenCelda("","vt COLSPAN=2 ","CENTER","","","","","OBSERVACIONES","si")
	    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

		    nCol = 110
		    t3b = "<font class=at><textarea style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#f7aa22;background-color:#ffffff;color:#000080;' "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=4 "
                    t3b = t3b & " readonly >"
		    t3b = t3b & rsXT(17) & "</textarea></font>"


	    celdas = ofv.GenCelda("","vt COLSPAN=2  STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","center","","","","",t3b,"si")
	    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	    sHTML = sHTML  & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:ridge;BORDER-WIDTH:2;border-color:#f7aa22;' ","","98%","","0","","0","0","","",filas,"si") & "<BR><BR>"


          END IF


	  session("WKFPSC") = RSX(9)
	  SESSION("SEGCKT") = RSX(5)

	  IF RSX(11) <> "NO" THEN

	       celdas1 = ofv.GenCelda("","vt style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ","CENTER","","","","","<B>PASO FINALIZADO:</B>","si")
	       celdas2 = ofv.GenCelda("","vt","CENTER","","","","","<b>" & sImagenff & "</b> " & RSX(11),"si")


	       filas  = ofv.GenRow("","","","","10","","",celdas1,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")



	       sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"



	  ELSE

	    StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RSX(9) & "' "
	    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	    if not rs2.eof then

	       psctipo  = RS2(14)
	       pscclase = RS2(3)

        '''     response.write "paso n1<br>"


	'''response.write WSORIG & " - " & WUORIG & " - " & WSDEST & " - " & WUDEST & "<br>"


               IF (RS2(8) OR RS2(9) OR RS2(11)) and rs2(12) <> "" THEN

           '''  response.write "paso x1<br>"

		  IF PSCTIPO = "U" THEN

		     StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
		     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		     if not rs3.eof then
			SINPUT = RS3(6) & ":" & RS3(7)


				       IF pscclase = "E" THEN

					  XRTIPO  = RS3(15)
					  XRCOD   = RS3(16)
					  XRASG   = "NO"

                                          XRLVL   = 9
                                          XRSCD   = "0"

                                          IF XRTIPO = "USR" OR XRTIPO = "RSP" THEN
                                             XRASG   = XRCOD
                                             IF XRTIPO = "RSP" THEN
                                                XRTIPO = "USR"
                                                FOR YYY=1 TO ATABLA(0,0)

                                                    IF ATABLA(YYY,0) =  XRCOD THEN
					               XRCOD   = ATABLA(YYY,1)
                                                       XRASG   = XRCOD
                                                       XRLVL   = ATABLA(YYY,4)
                                                       XRSCD   = ATABLA(YYY,5)
                                                       EXIT FOR
                                                    END IF

                                                NEXT

                                             END IF
                                          END IF

				       ELSE

				          XRTIPO  = ATABLA(2,2)
					  XRCOD   = ATABLA(2,1)
					  XRASG   = "NO"
					  IF XRTIPO  = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = ATABLA(2,4)
                                          XRSCD   = ATABLA(2,5)


				       END IF


			   CKTTIPO  = XRTIPO
			   CKTCOD   = XRCOD
			   CKTUSR   = XRASG


			SINPUT2 = GETRESPOCKT(CKTCOD,CKTTIPO)
			SINPUT3 = GETRESPOCKT(CKTUSR,"USR")

		     ELSE
			SINPUT = "S/D"
			SINPUT2 = "S/D"
			SINPUT3 = "S/D"
		     END IF
		  ELSE
			set cn2 = ofv.conectar(ofv.strconn0)

			sniv = 0

			'strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
			'set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			'if not rsniv.eof then sniv = rsniv(0)
			'call ofv.cerrarconsulta(rsniv)

                         xxxx = getInfoUserw(SESSION("USUARIO"))

                         if isnumeric(session("SGU_idcostcenter")) then sniv=session("SGU_idcostcenter")

		     ITAB = 0
		     itabs = 0
		     StrSql = "Select * From SEGCKTRESP WHERE SEGCKTPSC = '" & RSX(9) & "' AND SEGCKT = '" & SESSION("SEGCKT") & "' order by ORDEN "
		     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		     if not rs3.eof then
			DO UNTIL RS3.EOF
			   ITAB = ITAB + 1

                                           ATABLA(0,0) =  ATABLA(0,0) + 1
                                           ATABLA(0,1) =  ATABLA(0,1) + 1
				 '''          response.write itab

                                           ATABLA(ITAB + 2,0) =  "L"
                                           ATABLA(ITAB + 2,1) =  RS3(5)
                                           ATABLA(ITAB + 2,2) =  RS3(4)
                                           ATABLA(ITAB + 2,3) =  "NO"
                                           ATABLA(ITAB + 2,4) =  RS3(7)
                                           ATABLA(ITAB + 2,5) =  RS3(8)

                                           XWTIPO             =  RS3(4)
                                           XWCOD              =  RS3(5)
                                           XWSEC              =  RS3(8)

					''   if (cstr(sniv) = cstr(rs3(5))) or (cstr(sniv) = cstr(rs3(8))) then
					''      itabs = itab
					''   end if

                                           SELECT CASE XWTIPO

                                                  CASE "USR"


                                                        if XWCOD = SESSION("USUARIO") THEN
					                   itabs = itab
					                end if


                                                  CASE "GRP"

                                                        if XWCOD = SESSION("SEGGRP") THEN
					                   itabs = itab
					                end if

                                                  CASE "SEC"

                                                        if CSTR(XWCOD) = CSTR(SNIV) THEN
					                   itabs = itab
					                end if

                                            END SELECT





			rs3.movenext
			LOOP
		     END IF
		     if itabs = itab then
			StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
			set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			if not rs3.eof then
			   SINPUT = RS3(6) & ":" & RS3(7)


				       IF pscclase = "E" THEN

					  XRTIPO  = RS3(15)
					  XRCOD   = RS3(16)
					  XRASG   = "NO"

                                          XRLVL   = 9
                                          XRSCD   = "0"

                                          IF XRTIPO = "USR" OR XRTIPO = "RSP" THEN
                                             XRASG   = XRCOD
                                             IF XRTIPO = "RSP" THEN
                                                XRTIPO = "USR"
                                                FOR YYY=1 TO ATABLA(0,0)

                                                    IF ATABLA(YYY,0) =  XRCOD THEN
					               XRCOD   = ATABLA(YYY,1)
                                                       XRASG   = XRCOD
                                                       XRLVL   = ATABLA(YYY,4)
                                                       XRSCD   = ATABLA(YYY,5)
                                                       EXIT FOR
                                                    END IF

                                                NEXT

                                             END IF
                                          END IF

				       ELSE

				          XRTIPO  = ATABLA(2,2)
					  XRCOD   = ATABLA(2,1)
					  XRASG   = "NO"
					  IF XRTIPO  = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = ATABLA(2,4)
                                          XRSCD   = ATABLA(2,5)


				       END IF


			   CKTTIPO  = XRTIPO
			   CKTCOD   = XRCOD
			   CKTUSR   = XRASG







			   SINPUT2 = GETRESPOCKT(CKTCOD,CKTTIPO)
			   SINPUT3 = GETRESPOCKT(CKTUSR,"USR")
			 ELSE
			   SINPUT  = "S/D"
			   SINPUT2 = "S/D"
			   SINPUT3 = "S/D"
			 END IF

		     elseif itabs < itab then
			 if itabs <> 0 then

			    SINPUT = RS2(6) & "." & itabs + 1 & ":" & RS2(7)

				               XRTIPO  = ATABLA(itabs + 3,2)
					       XRCOD   = ATABLA(itabs + 3,1)
					       XRASG   = "NO"
                                               XRLVL   = ATABLA(itabs + 3,4)
                                               XRSCD   = ATABLA(itabs + 3,5)


			   CKTTIPO  = XRTIPO
			   CKTCOD   = XRCOD
			   CKTUSR   = XRASG


			    SINPUT2 = GETRESPOCKT(CKTCOD,CKTTIPO)
			    SINPUT3 = GETRESPOCKT(CKTUSR,"USR")


			 else
			    StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
			    set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			if not rs3.eof then
			   SINPUT = RS3(6) & ":" & RS3(7)



				       IF pscclase = "E" THEN

					  XRTIPO  = RS3(15)
					  XRCOD   = RS3(16)
					  XRASG   = "NO"

                                          XRLVL   = 9
                                          XRSCD   = "0"

                                          IF XRTIPO = "USR" OR XRTIPO = "RSP" THEN
                                             XRASG   = XRCOD
                                             IF XRTIPO = "RSP" THEN
                                                XRTIPO = "USR"
                                                FOR YYY=1 TO ATABLA(0,0)

                                                    IF ATABLA(YYY,0) =  XRCOD THEN
					               XRCOD   = ATABLA(YYY,1)
                                                       XRASG   = XRCOD
                                                       XRLVL   = ATABLA(YYY,4)
                                                       XRSCD   = ATABLA(YYY,5)
                                                       EXIT FOR
                                                    END IF

                                                NEXT

                                             END IF
                                          END IF

				       ELSE

				          XRTIPO  = ATABLA(2,2)
					  XRCOD   = ATABLA(2,1)
					  XRASG   = "NO"
					  IF XRTIPO  = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = ATABLA(2,4)
                                          XRSCD   = ATABLA(2,5)


				       END IF


			   CKTTIPO  = XRTIPO
			   CKTCOD   = XRCOD
			   CKTUSR   = XRASG





			   SINPUT2 = GETRESPOCKT(CKTCOD,CKTTIPO)
			   SINPUT3 = GETRESPOCKT(CKTUSR,"USR")
			ELSE
			   SINPUT  = "S/D"
			   SINPUT2 = "S/D"
			   SINPUT3 = "S/D"
			END IF

			 end if
		     end if


		  END IF



		  celdas1 = ofv.GenCelda("","vt style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00cc00;' ","center","50%","","","","<B>SIGUE A PASO:<br> <big>" & SINPUT & "</big></B>","si")
		  celdas2 = ofv.GenCelda("","vt","","","","","","<b>" & sImagene & "</b> " & "<B>DESTINO: " & SINPUT2 & "</B>","si")
		  celdas3 = ofv.GenCelda("","vt","","","","","","<b>" & sImageni & "</b> " & "<B>USUARIO: " & SINPUT3 & "</B>","si")
		  celdas4 = ofv.GenCelda("","vt","","","","","","&nbsp;","si")

                 IF INSTR(session("SGU_nombre"),"NO DISPONIBLE") > 0 THEN

		    sinput = "<b>&nbsp;VALIDACION NO DISPONIBLE&nbsp;</b>"

                 ELSE

		  sinput = "<a  class=cc href=cowkfselect.asp?OPCION=OK"
		  sinput = sinput  & " ><b>&nbsp;APRUEBA&nbsp;</b></a>"


                 END IF





		  celdas5 = ofv.GenCelda("","vt","center","","50","","","<b>" & sImagenA & "</b> " & "<BIG>" & sinput & "</BIG>","si")

	       ELSE
          '''   response.write "paso 2<br>"

               IF RS2(11) THEN
'''''''''''


				       IF pscclase = "E" THEN

                                     '     response.write "- rs3(15) y (16):" & strsql & "<br>"
                                     '     response.write " rs2(11):" & rs2(11) & "- rs3(15) y (16):" & RS3(15) & "-" & RS3(16) & "<br>"

					  XRTIPO  = RS2(15)
					  XRCOD   = RS2(16)
					  XRASG   = "NO"

                                          XRLVL   = 9
                                          XRSCD   = "0"

                                          IF XRTIPO = "USR" OR XRTIPO = "RSP" THEN
                                             XRASG   = XRCOD
                                             IF XRTIPO = "RSP" THEN
                                                XRTIPO = "USR"
                                                FOR YYY=1 TO ATABLA(0,0)

                                                    IF ATABLA(YYY,0) =  XRCOD THEN
					               XRCOD   = ATABLA(YYY,1)
                                                       XRASG   = XRCOD
                                                       XRLVL   = ATABLA(YYY,4)
                                                       XRSCD   = ATABLA(YYY,5)
                                                       EXIT FOR
                                                    END IF

                                                NEXT

                                             END IF
                                          END IF

				       ELSE

				          XRTIPO  = ATABLA(2,2)
					  XRCOD   = ATABLA(2,1)
					  XRASG   = "NO"
					  IF XRTIPO  = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = ATABLA(2,4)
                                          XRSCD   = ATABLA(2,5)


				       END IF





			SINPUT2 = GETRESPOCKT(XRCOD,XRTIPO)
			SINPUT3 = GETRESPOCKT(XRASG,"USR")



'''''''''''
		  celdas1 = ofv.GenCelda("","vt style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00cc00;' ","center","50%","","","","<B>SIGUE A PASO:<br> <big>&nbsp;FIN&nbsp;DEL&nbsp;CIRCUITO</big></B>","si")
		  celdas2 = ofv.GenCelda("","vt","","","","","","<b>" & sImagene & "</b> " & "<B>DESTINO: " & SINPUT2 & "</B>","si")
		  celdas3 = ofv.GenCelda("","vt","","","","","","<b>" & sImageni & "</b> " & "<B>USUARIO: " & SINPUT3 & "</B>","si")
		  celdas4 = ofv.GenCelda("","vt","","","","","","&nbsp;","si")






                 IF INSTR(session("SGU_nombre"),"NO DISPONIBLE") > 0 THEN

		    sinput = "<b>&nbsp;VALIDACION NO DISPONIBLE&nbsp;</b>"

                 ELSE

		  sinput = "<a  class=cc href=cowkfselect.asp?OPCION=OK"
		  sinput = sinput  & " ><b>&nbsp;APRUEBA&nbsp;</b></a>"

                 END IF

		  celdas5 = ofv.GenCelda("","vt","center","","50","","","<b>" & sImagenA & "</b> " & "<BIG>" & sinput & "</BIG>","si")

               else

		 celdas1 = ofv.GenCelda("","","","50%","","","","&nbsp;","si")
		 celdas2 = ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas3 = ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas4 = ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas5 = ofv.GenCelda("","","","","","","","&nbsp;","si")

               end if

	       END IF

	       IF RS2(10)  and rs2(13) <> "" THEN

		  StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(13) & "' "
		  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		     if not rs3.eof then
			SINPUT = RS3(6) & ":" & RS3(7)





				       IF pscclase = "E" THEN

					  XRTIPO  = RS3(15)
					  XRCOD   = RS3(16)
					  XRASG   = "NO"

                                          XRLVL   = 9
                                          XRSCD   = "0"

                                          IF XRTIPO = "USR" OR XRTIPO = "RSP" THEN
                                             XRASG   = XRCOD
                                             IF XRTIPO = "RSP" THEN
                                                XRTIPO = "USR"
                                                FOR YYY=1 TO ATABLA(0,0)

                                                    IF ATABLA(YYY,0) =  XRCOD THEN
					               XRCOD   = ATABLA(YYY,1)
                                                       XRASG   = XRCOD
                                                       XRLVL   = ATABLA(YYY,4)
                                                       XRSCD   = ATABLA(YYY,5)
                                                       EXIT FOR
                                                    END IF

                                                NEXT

                                             END IF
                                          END IF

				       ELSE
					     XRTIPO  = ATABLA(1,2)
					     XRCOD   = ATABLA(1,1)
					     XRASG   = ATABLA(1,1)
                                             XRLVL   = ATABLA(1,4)
                                             XRSCD   = ATABLA(1,5)


				       END IF




			   CKTTIPO  = XRTIPO
			   CKTCOD   = XRCOD
			   CKTUSR   = XRASG




			SINPUT2 = GETRESPOCKT(CKTCOD,CKTTIPO)
			SINPUT3 = GETRESPOCKT(CKTUSR,"USR")

		     ELSE
			SINPUT = "S/D"
			SINPUT2 = "S/D"
			SINPUT3 = "S/D"
		     END IF

		  celdas1 = celdas1 & ofv.GenCelda("","vT style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#cc0000;' ","center","","","","","<B>SIGUE A PASO:<br> <big>" & SINPUT & "</big></B>","si")
		  celdas2 = celdas2 & ofv.GenCelda("","vt","","","","","","<b>" & sImagene & "</b> " & "<B>DESTINO: " & SINPUT2 & "</B>","si")
		  celdas3 = celdas3 & ofv.GenCelda("","vt","","","","","","<b>" & sImageni & "</b> " & "<B>USUARIO: " & SINPUT3 & "</B>","si")
		  celdas4 = celdas4 & ofv.GenCelda("","vt","","","","","","&nbsp;","si")

                 IF INSTR(session("SGU_nombre"),"NO DISPONIBLE") > 0 THEN

		    sinput = "<b>&nbsp;VALIDACION NO DISPONIBLE&nbsp;</b>"

                 ELSE

		    sinput = "<a  class=cc href=cowkfselect.asp?OPCION=RCH"
		    sinput = sinput  & " ><b>&nbsp;RECHAZA&nbsp;</b></a>"

                 END IF


		  celdas5 = CELDAS5 & ofv.GenCelda("","vt","center","","50","","","<b>" & sImagenR & "</b> " & "<BIG>" & sinput & "</BIG>","si")

	       ELSE

		 celdas1 = CELDAS1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas2 = CELDAS2 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas3 = CELDAS3 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas4 = CELDAS4 & ofv.GenCelda("","","","","","","","&nbsp;","si")
		 celdas5 = CELDAS5 & ofv.GenCelda("","","","","","","","&nbsp;","si")

	       END IF

	       filas  = ofv.GenRow("","","","","10","","",celdas1,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas3,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas4,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas5,"si")


	       sHTML = sHTML  & ofv.GenTabla("","aaa  ","","100%","","0","","-2","-1","","",filas,"si")

	   end if

	  END IF

	 END IF

end if













  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"


  Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
  Filann = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%   ","","0","","1","0","","",Filann,"si")



  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

 '' response.write sopcion & " " & sopx & "<br>"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

FUNCTION GENCKTPSC(XPSC,XSIG,XMODEL,XFIN,XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XNIVEL,XSUBCOD)

    SEGCKTPSC = XPSC



    XCOM = GETDH()
    IF XFIN = "SI" THEN
       XFIN = GETDH()
    ELSE
       XFIN = "NO"
    END IF

    strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC,OBSERVACIONES,NIVEL,SUBCODIGO) values "
    STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & SEGCKTPSC & "','" & XMODEL & "','" & GETDH() & "','" & XFIN & "','" & XSIG & "','" & XRTIPO & "','" & XRCOD & "','" & XRASG & "','" & XPDSC & "','" & XOBSERV & "','" & XNIVEL & "','" & XSUBCOD & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION




FUNCTION GENACCION(XPASO,XTPP)

 '''  response.write "entro accion: " & XPASO & " -- " & XTPP & " // "


    StrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & XPASO & "' AND (TIPO = '" & XTPP & "' OR TIPO = 'FIN') "
   ' response.write strsql
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS1.EOF THEN
       IF RS1(4) = "A" THEN
'''   response.write "estado: " & RS1(5) & "<br>"
          IF  session("TCliente") = "RSK" THEN
              GETWKFTAB = "ORRSKCIRCUITO"
          ELSE
              GETWKFTAB = "OBJPRTY"
          END IF

          XXXX = GETWKFMANAGER(session("NCliente"),session("TCliente"),RS1(5),GETWKFTAB,"WKF")
       END IF


    END IF




END FUNCTION





FUNCTION GENCKTMAIL(WETIPO,WECOD,WEMESSAGE,WEORIGEN,WELVL)


    XRESPO = WECOD & ""

    if XRESPO <> "" AND XRESPO <> "NO" then

       set cn2 = ofv.conectar(ofv.strconn0)

       XPARA = GETRESPOCKT(WECOD,WETIPO)
       STRSQL = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "

     XGUMENSAGEw = WEMESSAGE

     XMESAURL  = ""

     IF  SESSION("URLTYP")  = "V" THEN

         XMESAURL  =  left(request.servervariables("SCRIPT_NAME"),instr(RIGHT(request.servervariables("SCRIPT_NAME"),LEN(request.servervariables("SCRIPT_NAME"))-1),"/"))

     END IF

     IF SESSION("URLPORT")  = "N" THEN
        XMESAURL   = "http://" & request.servervariables("SERVER_NAME")  & XMESAURL
      ELSE
        XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
     END IF

    ' XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

     XGUMENSAGEw = replace(XGUMENSAGEw,"<APLURL>","<U><A HREF='" & XMESAURL  & "' TARGET='_blank' >")
     XGUMENSAGEw = replace(XGUMENSAGEw,"</APLURL>","</A></U>")

     WEMESSAGE = XGUMENSAGEw


       SELECT CASE WETIPO

        CASE "USR"
           'STRSQL = STRSQL & " ID = '" & XRESPO & "' "
           xxxx = getInfoUserw(XRESPO)
           if session("SGU_mail")<> "" then    SINPUTX= session("SGU_mail")


            stowho = SINPUTX & ""
	   	         if len(stowho) > 0 then
	   	            stowho = trim(stowho)

	   	            subject = "SOLICITUD DE AUTORIZACION"
	   	            IF WETIPO = "USR" THEN
	   	               XPARA2 = XPARA
	   	            ELSE
	   	               XPARA2 = XPARA & " / " & session("SGU_nombre")
	   	            END IF
	   	          ''  message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
                            Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
	   	            CALL SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
                          end if


        CASE "GRP"

           'STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "

           call  getUsrIngroup(XRESPO)
	                  A_SGU_UIGuserid =split(session("SGU_UIGuserid"),chr(9))
	                  A_SGU_UIGdescri =split(session("SGU_UIGdescri") ,chr(9))

	                '   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	                  ' DO UNTIL RSXPA.EOF
	                   for a_i=0 to ubound(SGU_UIGdescri)-1
	                      xxxx = getInfoUserw(A_SGU_UIGuserid(a_i))
	   		      stowho = session("SGU_mail") & ""
	   		      if len(stowho) > 0 then
	   			 stowho = trim(stowho)

	   			 subject = "SOLICITUD DE AUTORIZACION"
	   			 IF WETIPO = "USR" THEN
	   			    XPARA2 = XPARA
	   			 ELSE
	   			    XPARA2 = XPARA & " / " & A_SGU_UIGdescri(a_i)
	   			 END IF
	   			''' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
                                 Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
	   			 CALL SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
	   		      end if

	   		   'RSXPA.MOVENEXT
		           next



        CASE "SEC"
           XRESPO = 0 & XRESPO

           STRSQL = STRSQL & " SECTOR = " & XRESPO & " and ( NOT nivel >  " & WELVL & " ) "

       END SELECT






  ' set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  ' DO UNTIL RSXPA.EOF

  '    stowho = RSXPA(0)
  '    if len(stowho) > 0 then
  '       stowho = trim(stowho)

  '       subject = "SOLICITUD DE AUTORIZACION"
  '       IF WETIPO = "USR" THEN
  '          XPARA2 = XPARA
  '       ELSE
  '          XPARA2 = XPARA & " / " & RSXPA(1)
  '       END IF
  '       message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
  '       CALL SendMail(message, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
  '    end if

  ' RSXPA.MOVENEXT
  ' LOOP
  ' call ofv.cerrarconsulta(RSXPA)

    END IF

END FUNCTION


FUNCTION GETRESPOCKT(RESPCOD,RESPTIP)

  GETRESPOCKT = "SIN ASIGNAR"



    XRESPO = RESPCOD & ""

    if XRESPO <> "" AND XRESPO <> "NO" then

       set cn2 = ofv.conectar(ofv.strconn0)
       set cn3 = ofv.conectar(ofv.strconn2)

       SELECT CASE RESPTIP

        CASE "USR"

           xxxx = getInfoUserw(XRESPO)
           sInputx = session("SGU_nombre")

          ' strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
          ' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          ' if not rsniv.eof then sInputx = rsniv(1)
          ' call ofv.cerrarconsulta(rsniv)


        CASE "GRP"
           strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInputx = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

        CASE "SEC"
           XRESPO = 0 & XRESPO
           strsql = "Select * From ECO where ECO = " & XRESPO
      '''     response.write strsql
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInputx = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

       END SELECT

       GETRESPOCKT = sInputx

    END IF


END FUNCTION

FUNCTION GETDH()

    XKD = datepart("d",DATE())
    XKM = datepart("m",DATE())
    XKY = datepart("yyyy",DATE())





    GETDH = XKD & "/" & XKM & "/" & XKY & " " & TIME()


END FUNCTION

%>