<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0





  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))


  set cn = ofv.conectar(ofv.strconn5)
  set cn4 = ofv.conectar(ofv.strconn6)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")
  act = request.querystring("act")


  sopx = sopcion
  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_OPCION=" & SOPCION & "_ACT=S"

  WKFAUTO = "WKF0001693621131313343163"
  WKFAPRO = "WKF0001693673726313343163"
  DESLEVL = 9
  SUPLEVL = 9


  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
  if not rs.eof then
     SESSION("NOMREQ") = "REQUERIMIENTO " & RS(2) 
     Strsql = "Select * From TIPOREQ where TIPOREQ = '" & RS(15) & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        WKFAUTO = RS(37)
        WKFAPRO = RS(38)
        DESLEVL = RS(35)
        SUPLEVL = RS(36)
     END IF

  END IF


	StrSql = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'OR' "
	set rsX0 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)

	StrSql = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'AN' "
	set rsX1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)

	StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "'  "
	set rsX2 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)


	'' DIM ATABLA(50,2)

 DIM ATABLA(200,5) 

 DIM ATABSW(5,3) 

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0232;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenP = "<B><FONT FACE=WebDINGS COLOR=#336699 SIZE=3>2</FONT></B>"
    sImagenRU = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenRG = "<B><FONT FACE=WEBDINGS COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
    sImagenRS = "<B><FONT FACE=WEBDINGS COLOR=#0000aa SIZE=3>G</FONT></B>"
    sImagenFI = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>5</FONT></B>"
    sImagenFF = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenSE = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>G</FONT></B>"
    sImagenMA = "<B><FONT FACE=wingDINGS COLOR=#aa0000 SIZE=4>F</FONT></B>"

    sImagenA = "<B><FONT FACE=WINGDINGS COLOR=#00aa00 SIZE=3>&#0252;</FONT></B>"
    sImagenR = "<B><FONT FACE=WINGDINGS COLOR=#cc0000 SIZE=3>&#0251;</FONT></B>"


    Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsx.eof then
       session("WKFPSC")    = RSX(9)
       SESSION("SEGCKT")    = RSX(5)
       SESSION("WKF")       = RSX(0)
       SESSION("ORGUSR")    = RSX(15)
       set cn2 = ofv.conectar(ofv.strconn0)
       sInputx = "COR ON-TIME"
       strsql = "Select EMAIL From USUARIOS where ID = '" & RSX(15) & "' "
       set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)
       SESSION("ORGMAIL") = SINPUTX
    END IF
    call ofv.cerrarconsulta(rsx)

''''''''''''''''''''''''''''''''''''
'' busca originante y destinatario

    ATABLA(0,0) =  2
    ATABLA(0,1) =  0

    if SESSION("WKF") = WKFAUTO THEN

       ATABLA(1,0) =  "O"
       ATABLA(1,1) =  RSX2(13)
       ATABLA(1,2) =  "USR"
       ATABLA(1,3) =  "NO"
       ATABLA(1,4) =  9
       ATABLA(1,5) =  RSX0(3)

       ATABLA(2,0) =  "D"
       ATABLA(2,1) =  RSX1(3)
       ATABLA(2,2) =  "SEC"
       ATABLA(2,3) =  "NO"
       ATABLA(2,4) =  9
       ATABLA(2,5) =  RSX1(4)



    '   WSORIG = RSX0(3)
    '   WUORIG = RSX2(13)
    '   WSDEST = RSX1(3)
    '   WUDEST = "NO"

    END IF

    IF SESSION("WKF") = WKFAPRO THEN

       ATABLA(1,0) =  "O"
       ATABLA(1,1) =  RSX2(14)
       ATABLA(1,2) =  "USR"
       ATABLA(1,3) =  "NO"
       ATABLA(1,4) =  9
       ATABLA(1,5) =  RSX1(3)

       ATABLA(2,0) =  "D"
       ATABLA(2,1) =  RSX2(13)
       ATABLA(2,2) =  "USR"
       ATABLA(2,3) =  "NO"
       ATABLA(2,4) =  9
       ATABLA(2,5) =  RSX0(3)

    '   WSORIG = RSX1(3)
    '   WUORIG = RSX2(14)
    '   WSDEST = RSX0(3)
    '   WUDEST = RSX2(13)

   END IF


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


''''''''''''''''''''''''''''''''''''''''''''''''''''''

			if rs(0) = WKFAUTO THEN

			   WSORIG = RSX0(3)
			   WUORIG = RSX2(13)
			   WSDEST = RSX1(3)
			   WUDEST = "NO"

			END IF

			IF RS(0) = WKFAPRO THEN

			   WSORIG = RSX1(3)
			   WUORIG = RSX2(14)
			   WSDEST = RSX0(3)
			   WUDEST = RSX2(13)

			END IF



'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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
					  IF XRTIPO = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = 9
                                          XRSCD   = "0"
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
                                       XXXX = GENCKTMAIL(XRTIPO,XRCOD,XPDSC & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL)
 '''''''''''''''''''


				    END IF
				    call ofv.cerrarconsulta(rs2)

			ELSE






					set cn2 = ofv.conectar(ofv.strconn0)

					sniv = 0
					strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
					set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
					if not rsniv.eof then sniv = rsniv(0)
					call ofv.cerrarconsulta(rsniv)


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
                                           ATABLA(ITAB + 2,3) =  "NO"
                                           ATABLA(ITAB + 2,4) =  RS3(7)
                                           ATABLA(ITAB + 2,5) =  RS3(8)

					 ''  if (clng(sniv) = clng(rs3(5))) or (clng(sniv) = clng(rs3(8))) then
					    if (cstr(sniv) = cstr(rs3(5))) or (cstr(sniv) = cstr(rs3(8))) then

					      itabs = itab
					   end if


					''   ATABLA(ITAB,0) = RS3(4) 
					''   ATABLA(ITAB,1) = RS3(5) 

					''   ATABLA(ITAB,2) = RS3(6)
					rs3.movenext 
					LOOP    
				     END IF 
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


					'   IF RS2(6) = 999 THEN
					'      XRTIPO  = "SEC"
					'      XRCOD   = WSDEST
					'      XRASG   = WUDEST
					'   ELSE
					'      XRTIPO  = "SEC"
					'      XRCOD   = WSORIG
					'      XRASG   = WUORIG
					'   END IF  

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

					 '      XRTIPO  = "SEC"
					 '      XRCOD   = ATABLA(itabs + 1,1)
					 '      XRASG   = "NO"




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



				''' GENERAR EL PRIMER PASO OK
				       WXX = GENCKTPSC(SEGCKTPSC,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD)

 '''''''''''''''''''
 ' agregar envio de mail
                                       XXXX = GENCKTMAIL(XRTIPO,XRCOD,XPDSC & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL)
 '''''''''''''''''''


			END IF

			''''''''''''''''''''''''''''' ACTUALIZA PASO Y CIRCUITO DEL WORKFLOW CORRIENTE ''''''''''''''''''''''''''''''''''

			    strsql = "UPDATE SEGPASO SET FIN = '" & NOW() & "', OKRCH = '" & SEGCKTPSC & "', observaciones = '" & sobserv & "' WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
			    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			    IF RS(11) THEN

				     strsql = "UPDATE SEGCIRCUITO SET FIN = '" & NOW() & "'  WHERE SEGCKT = '" & SESSION("SEGCKT") & "' "
				     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			    END IF


		  END IF




	 	  Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	 	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	  if NOT rsx.eof then
	 
	 
			  celdas = ofv.GenCelda("","TTt colspan=2","","","","","","<BIG><B>WORKFLOW</B></BIG>","si")
			  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","ut","","15%","","","","CIRCUITO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","ut","","15%","","","","PASO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenp & "</b> " & rsX(16),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			  
			  celdas = ofv.GenCelda("","ut","","15%","","","","OPCION:","si")
			  IF SOPCION = "OK" THEN
			     SOPCDESC = " <b>" & sImagenA & "</b> " & "APRUEBA"
			  ELSE
			     SOPCDESC = " <b>" & sImagenR & "</b> " & "RECHAZA"
			  END IF
			     
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","",SOPCDESC,"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			  
			  IF RSX(11) <> "NO" THEN

			       celdas = ofv.GenCelda("","ut colspan=2 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ","CENTER","","","","","<B>PASO FINALIZADO:</B>","si")
			       celdas2 = ofv.GenCelda("","ut colspan=2","CENTER","","","","","<b>" & sImagenFF & "</b> " & RSX(11),"si")

			       filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			       Filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")


			  END IF
			  
			  

			  sHTML = sHTML  & ofv.GenTabla(""," ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"
	 	  
	 	  END  IF






	 ELSE
	 
	 
	 	  Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	 	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	  if NOT rsx.eof then
	 
	 
			  celdas = ofv.GenCelda("","ttt colspan=2","","","","","","<BIG><B>WORKFLOW</B></BIG>","si")
			  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","ut","","15%","","","","CIRCUITO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","ut","","15%","","","","PASO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenP & "</b> " & rsX(16),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			  
			  celdas = ofv.GenCelda("","ut","","15%","","","","OPCION:","si")
			  IF SOPCION = "OK" THEN
			     SOPCDESC = " <b>" & sImagenA & "</b> " & "APRUEBA"
			  ELSE
			     SOPCDESC = " <b>" & sImagenR & "</b> " & "RECHAZA"
			  END IF
			     
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","",SOPCDESC,"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			  

			  sHTML = sHTML  & ofv.GenTabla(""," ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"
	 	  
	 	  END  IF


		    Celda = ofv.GenCelda("","ut colspan=3 style='border-style:solid;border-width:1;border-color:#888888;background-color:#f9f9f9;' ","center","100%","","","","OBSERVACIONES/COMENTARIOS","si")
		    FilaS = Ofv.GenRow("","","","","","","",Celda,"si")

		    nCol = 80

		    wdescri = "SIN OBSERVACIONES" 

		    t3b = "<font class=at><textarea class=dt style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;background-color:#ffffff;color:#000080;' "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=15 "
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

                    sinput = "<input CLASS=BT type=submit value='CONTINUAR' size=10 >"
                    celda = ofv.GenCelda("","abc colspan=3","center","","1%","","",sInput,"si")
		    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


		    if bModificar then
		      sAccion = "COWKFREQUER.ASP?CantRegMover=0&OPCION=" & SOPCION & "&ACT=S"
		      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
		    else
		      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
		    end if


		    sHTML = sHTML  & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;' ","","100%","","0","","-2","-1","","",formulario0,"si") & "<BR><BR>"

	  END IF

 ELSE



	  Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsx.eof then


	  celdas = ofv.GenCelda("","TTt colspan=2","","","","","","<BIG><B>WORKFLOW</B></BIG>","si")
	  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","ut","","15%","","","","CIRCUITO:","si")
	  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","ut","","15%","","","","PASO:","si")
	  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenp & "</b> " & rsX(16),"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  sHTML = sHTML  & ofv.GenTabla(""," ","","98%","","0","","0","0","","",filas,"si") & "<BR>"




	  Strsql = "SELECT * FROM CKTVIEW WHERE OKRCH = '" & session("SEGCKTPSC") & "' "
	  set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
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

   

	  celdas = ofv.GenCelda("","uT ","","50%","","","","VIENE DE <B>" & rsXT(16) & "</B>","si")
	  celdas = celdas & ofv.GenCelda("","uT ","RIGHT","50%","","","", PSCMOT & "" & RSXT(11) & "","si")
	  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			set cn2 = ofv.conectar(ofv.strconn0)

			sniv  = "DESCONOCIDO"
                        SNSEC = "DESCONOCIDO"

			strsql = "Select DESCRIPCION,SECTOR From USUARIOS where ID = '" & rsXT(15) & "' "
			set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			if not rsniv.eof then
                               sniv = rsniv(0)
                               SNSEC = GETRESPOCKT(rsniv(1),"SEC") 
                        END IF
			call ofv.cerrarconsulta(rsniv)

	  celdas = ofv.GenCelda("","ut colspan=2","","","","","","RESPONSABLE:&nbsp;" & sImagenRU & "&nbsp;" & sniv,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
	  celdas = ofv.GenCelda("","ut colspan=2 ","","","","","","AREA/SECTOR:&nbsp;" & sImagenSE & "&nbsp;" & SNSEC,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","ut COLSPAN=2 STYLE='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;background-color:#f9f9f9;' ","CENTER","","","","","OBSERVACIONES","si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

		    nCol = 110
		    t3b = "<font class=at><textarea style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;background-color:#ffffff;color:#000080;' "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=4 "
                    t3b = t3b & " readonly >"
		    t3b = t3b & rsXT(17) & "</textarea></font>"


	  celdas = ofv.GenCelda("","ut COLSPAN=2   ","center","","","","",t3b,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  sHTML = sHTML  & ofv.GenTabla("","aaa style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#888888;' ","","98%","","0","","0","0","","",filas,"si") & "<BR><BR>"


          END IF


	  session("WKFPSC") = RSX(9)
	  SESSION("SEGCKT") = RSX(5)

	  IF RSX(11) <> "NO" THEN

	       celdas1 = ofv.GenCelda("","ut style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ","CENTER","","","","","<B>PASO FINALIZADO:</B>","si")
	       celdas2 = ofv.GenCelda("","ut","CENTER","","","","","<b>" & sImagenff & "</b> " & RSX(11),"si")


	       filas  = ofv.GenRow("","","","","10","","",celdas1,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")



	       sHTML = sHTML  & ofv.GenTabla(""," ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"



	  ELSE

	    StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RSX(9) & "' "
	    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	    if not rs2.eof then

	       psctipo  = RS2(14)
	       pscclase = RS2(3)

	       if rs2(0) = WKFAUTO THEN

		  WSORIG = RSX0(3)
		  WUORIG = RSX2(13)
		  WSDEST = RSX1(3)
		  WUDEST = "NO"

	       END IF
	       IF RS2(0) = WKFAPRO THEN

		  WSORIG = RSX1(3)
		  WUORIG = RSX2(14)
		  WSDEST = RSX0(3)
		  WUDEST = RSX2(13)

	       END IF

	'''response.write WSORIG & " - " & WUORIG & " - " & WSDEST & " - " & WUDEST & "<br>"

               IF (RS2(8) OR RS2(9) OR RS2(11)) and rs2(12) <> "" THEN

		  IF PSCTIPO = "U" THEN

		     StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & RS2(12) & "' "
		     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		     if not rs3.eof then
			SINPUT = RS3(6) & ":" & RS3(7)


				       IF pscclase = "E" THEN

					  XRTIPO  = RS3(15)
					  XRCOD   = RS3(16)
					  XRASG   = "NO"
					  IF XRTIPO = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = 9
                                          XRSCD   = "0"
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

			strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
			set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			if not rsniv.eof then sniv = rsniv(0)
			call ofv.cerrarconsulta(rsniv)


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
                                           

					   if (cstr(sniv) = cstr(rs3(5))) or (cstr(sniv) = cstr(rs3(8))) then

					      itabs = itab
					   end if


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
					  IF XRTIPO = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = 9
                                          XRSCD   = "0"
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
					  IF XRTIPO = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = 9
                                          XRSCD   = "0"
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



		  celdas1 = ofv.GenCelda("","ut style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00cc00;background-color:#f9f9f9;' ","center","50%","","","","<B>SIGUE A PASO:<br>" & SINPUT & "</B>","si")
		  celdas2 = ofv.GenCelda("","ut","","","","","","<b>" & sImagene & "</b> " & "DESTINO: " & SINPUT2 & "","si")
		  celdas3 = ofv.GenCelda("","ut","","","","","","<b>" & sImageni & "</b> " & "USUARIO: " & SINPUT3 & "","si")
		  celdas4 = ofv.GenCelda("","ut","","","","","","&nbsp;","si")


		  sinput = "<a  class=cc href=cowkfrequer.asp?OPCION=OK"
		  sinput = sinput  & " ><b>&nbsp;APRUEBA&nbsp;</b></a>" 

		  sinput = ""

		  sinput = sinput  & "<form name='frmwkfok' action='cowkfrequer.asp?OPCION=OK'"
		  sinput = sinput  & " method=post>"
		  sinput = sinput  & "<input type=submit size=10 name=botonok value=' APRUEBA ' style='background-color:#00cc00;color:#ffffff;' >"
		  sinput = sinput  & "</input>" & chr(13)
		  sinput = sinput  & "</form>"


		  celdas5 = ofv.GenCelda("","ut style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#dddddd;background-color:#f9f9f9;' ","center","","30","","",sinput,"si")

	       ELSE

               IF RS2(11) THEN
'''''''''''

				          XRTIPO  = ATABLA(2,2)
					  XRCOD   = ATABLA(2,1)
					  XRASG   = "NO"
					  IF XRTIPO  = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = ATABLA(2,4)
                                          XRSCD   = ATABLA(2,5)


			SINPUT2 = GETRESPOCKT(XRCOD,XRTIPO)
			SINPUT3 = GETRESPOCKT(XRASG,"USR")

	

'''''''''''
		  celdas1 = ofv.GenCelda("","ut style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00cc00;background-color:#f9f9f9;' ","center","50%","","","","<B>SIGUE A PASO:<br>&nbsp;FIN&nbsp;DEL&nbsp;CIRCUITO</B>","si")
		  celdas2 = ofv.GenCelda("","ut","","","","","","<b>" & sImagene & "</b> " & "DESTINO: " & SINPUT2 & "","si")
		  celdas3 = ofv.GenCelda("","ut","","","","","","<b>" & sImageni & "</b> " & "USUARIO: " & SINPUT3 & "","si")
		  celdas4 = ofv.GenCelda("","ut","","","","","","&nbsp;","si")


		  sinput = "<a  class=cc href=cowkfrequer.asp?OPCION=OK"
		  sinput = sinput  & " ><b>&nbsp;APRUEBA&nbsp;</b></a>" 

		  sinput = ""

		  sinput = sinput  & "<form name='frmwkfok' action='cowkfrequer.asp?OPCION=OK'"
		  sinput = sinput  & " method=post>"
		  sinput = sinput  & "<input type=submit size=10 name=botonok value=' APRUEBA ' style='background-color:#00cc00;color:#ffffff;' >"
		  sinput = sinput  & "</input>" & chr(13)
		  sinput = sinput  & "</form>"



		  celdas5 = ofv.GenCelda("","ut style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#dddddd;background-color:#f9f9f9;' ","center","","30","","",sinput,"si")

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
					  IF XRTIPO = "USR" THEN
					     XRASG   = XRCOD
					  END IF
                                          XRLVL   = 9
                                          XRSCD   = "0"
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

		  celdas1 = celdas1 & ofv.GenCelda("","uT style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#cc0000;background-color:#f9f9f9;' ","center","","","","","<B>SIGUE A PASO:<br>" & SINPUT & "</B>","si")
		  celdas2 = celdas2 & ofv.GenCelda("","ut","","","","","","<b>" & sImagene & "</b> " & "DESTINO: " & SINPUT2 & "","si")
		  celdas3 = celdas3 & ofv.GenCelda("","ut","","","","","","<b>" & sImageni & "</b> " & "USUARIO: " & SINPUT3 & "","si")
		  celdas4 = celdas4 & ofv.GenCelda("","ut","","","","","","&nbsp;","si")

		  sinput = "<a  class=cc href=cowkfrequer.asp?OPCION=RCH"
		  sinput = sinput  & " ><b>&nbsp;RECHAZA&nbsp;</b></a>" 

		  sinput = ""

		  sinput = sinput  & "<form name='frmwkfrch' action='cowkfrequer.asp?OPCION=RCH'"
		  sinput = sinput  & " method=post>"
		  sinput = sinput  & "<input type=submit size=10 name=botonrch value=' RECHAZA ' style='background-color:#cc0000;color:#ffffff;' >"
		  sinput = sinput  & "</input>" & chr(13)
		  sinput = sinput  & "</form>"

		  celdas5 = CELDAS5 & ofv.GenCelda("","ut style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#dddddd;background-color:#f9f9f9;' ","center","","30","","",sinput,"si")

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


	       sHTML = sHTML  & ofv.GenTabla("","aaa style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#dddddd;' ","","100%","","0","","2","-1","","",filas,"si")

	   end if

	  END IF

	 END IF

end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

 '' response.write sopcion & " " & sopx & "<br>"

FUNCTION GENCKTPSC(XPSC,XSIG,XMODEL,XFIN,XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XNIVEL,XSUBCOD)

    SEGCKTPSC = XPSC



    XCOM = NOW()
    IF XFIN = "SI" THEN
       XFIN = NOW()
    ELSE
       XFIN = "NO"
    END IF

    strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC,OBSERVACIONES,NIVEL,SUBCODIGO) values "
    STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & SEGCKTPSC & "','" & XMODEL & "','" & NOW() & "','" & XFIN & "','" & XSIG & "','" & XRTIPO & "','" & XRCOD & "','" & XRASG & "','" & XPDSC & "','" & XOBSERV & "','" & XNIVEL & "','" & XSUBCOD & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

FUNCTION GENACCION(XPASO,XTIPO)

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini


    StrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & XPASO & "' AND (TIPO = '" & XTIPO & "' OR TIPO = 'FIN') "
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS1.EOF THEN
       IF RS1(4) = "A" THEN     
           strsql = "UPDATE REQUERIMIENTO SET ESTADO = " & RS1(5) & " WHERE codigo = '" & session("NCliente") & "' "
           call ofv.crearconsultaEx(StrSql,cn4,1,parametros)
          
          SELECT CASE RS1(5)

                 CASE 3  ' PENDIENTE ASIGNACION ---- ACT FECHA SOLIC.

                      strsql = "UPDATE REQUERIMIENTO SET FECHASOL = " & vNuevoValor & " WHERE codigo = '" & session("NCliente") & "' "
                      call ofv.crearconsultaEx(StrSql,cn4,1,parametros)

                 CASE 6  ' ENTREGADO --------------- ACT FECHA FINALIZACION.

                      strsql = "UPDATE REQUERIMIENTO SET FECHAFIN = " & vNuevoValor & " WHERE codigo = '" & session("NCliente") & "' "
                      call ofv.crearconsultaEx(StrSql,cn4,1,parametros)



          END SELECT  

       END IF
    
    
    END IF
 



END FUNCTION


FUNCTION GENCKTMAIL(WETIPO,WECOD,WEMESSAGE,WEORIGEN,WELVL)


    XRESPO = WECOD & ""

    if XRESPO <> "" AND XRESPO <> "NO" then

       set cn2 = ofv.conectar(ofv.strconn0)

       XPARA = " DESTINATARIO " & GETRESPOCKT(WECOD,WETIPO) 
       STRSQL = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "


       SELECT CASE WETIPO
  
        CASE "USR"
           STRSQL = STRSQL & " ID = '" & XRESPO & "' "


        CASE "GRP"

           STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "


        CASE "SEC"
           XRESPO = 0 & XRESPO

           STRSQL = STRSQL & " SECTOR = " & XRESPO & " and ( NOT nivel >  " & WELVL & " ) "

       END SELECT






   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
   DO UNTIL RSXPA.EOF

      stowho = RSXPA(0)
      if len(stowho) > 0 then
         stowho = trim(stowho)
       
         subject = "SOLICITUD DE AUTORIZACION"
         IF WETIPO = "USR" THEN
            XPARA2 = XPARA
         ELSE
            XPARA2 = XPARA & " / " & RSXPA(1) 
         END IF 
         message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
         CALL SendMail(message, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
      end if

   RSXPA.MOVENEXT
   LOOP
   call ofv.cerrarconsulta(RSXPA)

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
           strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInputx = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

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

%>