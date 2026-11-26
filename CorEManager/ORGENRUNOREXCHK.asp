<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
Response.buffer = TRUE
ok = ofv.CheckUsuario()
IF ok <> "Y" THEN
   uv = ofv.uv
   IF uv = "NO" THEN ok = ofv.ValidarUsuario
   sHTML = ok
ELSE
   response.expires = 0
  '' server.scripttimeout=60*20
   SET cn = ofv.conectar(ofv.strconn2)

   CantRegMostrarx = 8
   CantRegMoverx = CDBL(request.querystring("CantRegMover"))
   vuelta = ofv.vueltaasp("../", CantRegMostrarx, Cantregmoverx)
   DIM IDs(2, 1000)
   DIM IDsEx(2, 1)
   DIM IDSpoS
   DIM POSI
   DIM CContable
   DIM primerReg
   DIM totalmontoRepet

   Dim TeventosRepetitivos
   Dim TeventosEncadenados
   DIM mes
   DIM fecha1
   DIM Ahijos(500)
   DIM aRecordSave(100, 10)
   DIM aRecordSave2(100, 0)

   DIM TOPEMONTO
   DIM ENCAMENORES
   DIM REPESUMA
   DIM SORFECINIHW1

   dim REGVA
   primerReg = TRUE
   sopcion = request.querystring("opcion") & ""
   idspOS = 0
   DIM EVENTORAIZ
   EVENTORAIZ = 1
   ULTEVENTORAIZ = 0
   PRIEVENTORAIZ = 0
   TeventosRepetitivos=0
   TeventosEncadenados=0
   totalmontoRepet =0
   SORFECINIHW1=""
   TOPEMONTO=0

   strimestre = 0
   stANO      = 0

   'strimestre= request.FORM("TRIMESTRE")

   ORFECHA = REQUEST.FORM("ORFECHA")

   IF NOT ISDATE(ORFECHA) THEN
      SX = GetFecTrim(strimestre, YEAR(DATE()), feciniT, fecfinT)
      feciniT = CLNG(CDATE(feciniT)) - CLNG(CDATE("01-01-1900"))
      fecfinT = CLNG(CDATE(fecfinT)) - CLNG(CDATE("01-01-1900"))
   ELSE
      iTRIM = GetTrimestre(CDATE(ORFECHA))
      SX = GetFecTrim(iTRIM, YEAR(CDATE(ORFECHA)), feciniT, fecfinT)
      feciniT = CLNG(CDATE(ORFECHA)) - CLNG(CDATE("01-01-1900"))
      fecfinT = CLNG(CDATE(fecfinT)) - CLNG(CDATE("01-01-1900"))
   END IF


   LOTE = REQUEST.QUERYSTRING("LOTE")

   STRSQL = "SELECT * FROM ORCONFIGBCRA"
   SET RS = OFV.crearconsultaEx(StrSql, cn, 1, volver)
   IF NOT RS.EOF THEN
      FCARGA          = RS("FECREGIS")
      FDESCUBRIMIENTO = RS("FECIDENT")
      FORIGEN         = RS("FECORIG")
      FCONTABLE       = RS("FECCONTA")
      FFINAL          = RS("FECFINAL")
      TOPEMONTO       = CDBL(RS("TOPEMONTO"))
      ENCAMENORES     = RS("ENCAMENORES")
      REPESUMA        = RS("REPESUMA")
   END IF

 '  response.write TOPEMONTO
   CALL OFV.CERRARCONSULTA(RS)


      nrolote     = 0 


   STRSQL = "SELECT * FROM ORREGINFOBCRA WHERE CODLOTE='" & LOTE & "'"
   SET RS = OFV.crearconsultaEx(StrSql, cn, 1, volver)
   IF NOT RS.EOF THEN
      strimestre = RS("TRIMESTRE")
      stANO      = RS("ANO")
      SX = GetFecTrim(strimestre, RS("ANO"), feciniT, fecfinT)
      feciniT     = CLNG(CDATE(feciniT)) - CLNG(CDATE("01-01-1900"))
      fecfinT     = CLNG(CDATE(fecfinT)) - CLNG(CDATE("01-01-1900"))
      ScodDiseno  = rs("coddiseno")
      Scodentidad = rs("codentidad")
      nrolote     = rs(0) 
   END IF
   CALL OFV.CERRARCONSULTA(RS)






   strsql = ""
   strsql = strsql & " SELECT ULTEVENTO FROM ORREGINFOBCRA "
   strsql = strsql & " where (ano = '0" & stANO & "' and trimestre < '0" & strimestre & "' ) or (ano < '0" & stANO & "' ) "
   strsql = strsql & " order by ano desc, trimestre desc, estado desc, objid desc "
   SET RS = OFV.crearconsultaEx(StrSql, cn, 1, volver)
   IF NOT RS.EOF THEN
      EVENTORAIZ    = CLNG("0" & rs(0))
      PRIEVENTORAIZ = EVENTORAIZ
      EVENTORAIZ    = EVENTORAIZ + 1 
   end if




   WSPSESION = "SWHERERUNOR"
   SESSION(WSPSESION) = ""

   SCONECTOR = ""

   WKSETTABLA = "ORTRACER"

   IF (session("APLMDLOR") & "") = "BE" THEN
      Strsql = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRIpcion AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
   ELSE
      Strsql = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
   END IF

   'strsql= strsql & " inner join orestados on (ORTRACER.trcestado=orestados.codigo and orestados.envia='S')"
   strsql = strsql & " inner join orestados on (ORTRACER.trcestado=orestados.codigo )"
   strsql = strsql & " left join eco E on (ORTRACER.trcecoresp=E.eco)  "


   IF (session("APLMDLOR") & "") = "BE" THEN
      STRSQL = STRSQL & " LEFT JOIN TABPROCESOS01 O ON (ORTRACER.trcPROCESO=O.OBJID)"
   ELSE
      STRSQL = STRSQL & " LEFT JOIN OBJETOS O ON (ORTRACER.trcPROCESO=O.CODOBJETO)"
   END IF


   strsql = strsql & " left join ORTIPOPRODUCTO OT on (ORTRACER.trcPRODUCTO=OT.OBJID)  "
   strsql = strsql & " left join ORTIPOCATEGO OTCAT on (ORTRACER.TRCEVENTORIESGO2 = OTCAT.OBJID)  "
   strsql = strsql & " left join ORRIESGOASOC ORAS on (ORTRACER.TRCRIESGOASOC = ORAS.OBJID)  "

   strsql = strsql & " left join ORLINNEGO2 OL2 on (ORTRACER.trcNEGOLINE=OL2.OBJID)  "
   strsql = strsql & " left join ORTIPOCARACTER OC on (ORTRACER.trcCARACTER=OC.OBJID)  "

   IF FCARGA <> "N" THEN



      SORFECINIHW = " (ORTRACER.FECREGIS > = " & feciniT & " and ORTRACER.FECREGIS < = "& fecfinT & ") "

      SORFECINIHW1 = " (b.FECREGIS > = " & feciniT & " and b.FECREGIS < = "& fecfinT & ")"  & SCONECTOR

      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
      SCONECTOR = " OR "

   END IF


   IF FDESCUBRIMIENTO <> "N" THEN



      SORFECINIHW = " (ORTRACER.FECIDENT > = " & feciniT & " and ORTRACER.FECIDENT < = "& fecfinT & ") "
      SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECIDENT > = " & feciniT & " and b.FECIDENT < = "& fecfinT & ") "

      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
      SCONECTOR = " OR "

   END IF



   IF FORIGEN <> "N" THEN



      SORFECINIHW = " (ORTRACER.FECORIG > = " & feciniT & " and ORTRACER.FECORIG < = "& fecfinT & ") "
      SORFECINIHW1 = SORFECINIHW1 & SCONECTOR &  " (b.FECORIG > = " & feciniT & " and b.FECORIG < = "& fecfinT & ") "
      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
      SCONECTOR = " OR "

   END IF

   IF FCONTABLE <> "N" THEN



      SORFECINIHW = " (ORTRACER.FECCONTA > = " & feciniT & " and ORTRACER.FECCONTA < = "& fecfinT & ") "
      SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECCONTA > = " & feciniT & " and b.FECCONTA < = "& fecfinT & ") "
      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
      SCONECTOR = " OR "

   END IF


   IF FFINAL <> "N" THEN



      SORFECINIHW = " (ORTRACER.FECFINENR > = " & feciniT & " and ORTRACER.FECFINENR < = "& fecfinT & ") "
      SORFECINIHW1 = SORFECINIHW1 & SCONECTOR & " (b.FECFINENR > = " & feciniT & " and b.FECFINENR < = "& fecfinT & ") "
      SESSION(WSPSESION) = SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
      SCONECTOR = " OR "

   END IF


   swhere = "  where NOT (ortracer.codTRACER like 'DOBLIGA%') "
   swhere = swhere & "  AND (ortracer.trccodtracerp IS NULL OR  ortracer.trccodtracerp = '') "



   IF session(WSPSESION ) <> "" THEN
      swhere = swhere & "  and ( ortracer.TRCMARCASEND <> 'S' AND  (" & SESSION(WSPSESION) & ") ) "
   END IF


   STRSQL = STRSQL & swhere
   strsql = strsql & " 	or (ortracer.TRCMARCAACTP <> 'S' AND (ORTRACER.TRCFECACTP > = " & feciniT & " and ORTRACER.TRCFECACTP < = "& fecfinT & "))"
   strsql = strsql & " 	or (ortracer.TRCMARCAACTH <> 'S' AND (ORTRACER.TRCFECACTH > = " & feciniT & " and ORTRACER.TRCFECACTH < = "& fecfinT & "))"


   STRSQL = STRSQL & " ORDER BY ORTRACER.TRCCTACBLE,oc.CODBCRA,ORTRACER.FECIDENT,ORTRACER.FECREGIS "
   sHTML = ofv.FormHeader(session("FormN"))


  '''  RESPONSE.WRITE STRSQL & "<BR><BR><BR><BR>"

   Formulario = ""





   'Encabezados
   '  Formulario = ofv.Encabezados("ID EVENTO","COD","ENTIDAD","F.INF","EV RAIZ","SECU","CARACTER","F.INI","F.DESC","F.REG.CONT","F.FIN","DESC","T.EVEN","L.NEGO","SUC","PROCESO","PROD/SERV","MONTO","PART.CONTABLE","COB.POL","VINC.RIESGO","AERA/SOP")
   'response.write strsql
   SET rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)










   TOTALREGBRUTOMENOR = CDBL("0")
   TOTALBRUTOMENOR = CDBL("0")

   TOTALREGBRUTOMAYOR = CDBL("0")
   TOTALBRUTOMAYOR = CDBL("0")

   TOTALENVIA = 0
   TOTALNOENVIA = 0

      POSIX = 0


   x = MOSTRARINFO(RS)

    RESPONSE.WRITE "<BR>-----------------------------<BR>total&nbsp;" & nrolote & ":&nbsp;" & posi & "<BR><BR>"


   sHTML = sHTML & ofv.GenTabla("", "", "", "80% ", "", "0", "", "0", "0", "", "", Formulario & x, "si")

   sAccion = "ORGENRUNOR.asp?CantRegMover=" & CantRegMoverX & ""
   Celda = ofv.BotonVolver(sAccion)




   CALL ofv.cerrarconn(cn)

   '

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''' FUNCIONES ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   FUNCTION MOSTRARINFO(RS)

      POSI = 0

      DO WHILE (NOT rs.EOF)

       '  if rs(0)=1893 then

            posiX=posiX + 1
           RESPONSE.WRITE rs(0) & "--- (R)<BR>"

         CALL VALIDALOTE2(RS, LOTE)
       '  end if
         '  posi = posi + 1



         rs.Movenext

      LOOP
' CALL GrabarReg(aRecordSave, ids, idspOS) 'se graban como repetitivos iguales dato pero no el monto


    IF totalmontoRepet >= TOPEMONTO  AND REPESUMA="S" THEN

                           CALL GrabarReg(aRecordSave, ids, idspOS-1,"1") 'se graban como repetitivos iguales dato pero no el monto


                     ELSE ' si no suman 1000 en el mes no van

                        FOR iz = 0 TO idspOS -1
                           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                           STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"& ids(0, iz) &"','1002')"

                         '  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                        NEXT

                     END IF

      MOSTRARINFO = fila
   END FUNCTION



   FUNCTION VALIDALOTE2 (RS, LOTE)


      IF rs("envia") = "S" THEN


         Xtipo = VerificarExtipo(rs, feciniT, fecfinT)
  ' if rs(0)= 1949 then response.write "<br><br><br><br>"&Xtipo&"<br><br><br><br>"
         ' XTipo N=Normal , R=Rectificado ,RH=Rectificado Hijo -- SEGUN SQL 3 CASOS DEL WHERE
         'RESPONSE.WRITE XTipo & "--"
    '     response.write rs("CARACTER") & "---" & Xtipo & "<br>"
         SELECT CASE XTipo


            CASE "N" ' NORMAL


               REGVA = VALIDACAMPOS(RS)


               '--------------------- TIPOS DE caracter -------------------------------------


               IF (CSTR(rs("CARACTER") & "") = "03" OR CSTR(rs("CARACTER") & "") = "04" ) AND REPESUMA="S"  THEN ' ----- eventos repetitivos

               'if rs(0)= 1957 then response.write "<br><br><br><br>"&REGVA&"<br><br><br><br>"

                IF REGVA THEN

                  IF primerReg  and CDBL(RS("TRCIMPBRUTO")) < TOPEMONTO  THEN

                     CALL guardarRs(aRecordSave,rs, 0)
                     primerReg = FALSE
                     fecha1 = CDATE(rs("FECIDENT") + CLNG(CDATE("1-1-1900")))
                     mes = MONTH(fecha1)
                     CContable = TRIM(rs("trcctaCble"))

                  END IF

                if CDBL(RS("TRCIMPBRUTO"))  >=  TOPEMONTO  and regva then

                  SET rShIJOSEnc = getEncadenados(rs)
                  CALL Grabar(rs, secuencia, rShIJOSEnc, STRIMESTRE)
                  call ofv.cerrarconsulta(rShIJOSEnc)
				'  CALL guardarRs(aRecordSave2,rs, 0)
				'  idsex(0, 0) = CDBL(rs(0))
				'  idsex(1, 0) = CDBL(REPLACE(RS("TRCIMPBRUTO"), ",", "."))
				'  idspOSz=1
				'  SET rShIJOSEnc = getEncadenados(rs)

				'
				'	 REGVAENC = VALIDACAMPOS(rShIJOSEnc)


				'	   IF REGVAENC THEN
		          '      idsex(0, idspOSz) = CDBL(rShIJOSEnc(0))
				'	    idsex(1, idspOSz) = CDBL(rShIJOSEnc("TRCIMPBRUTO"))
				'		idspOSz = idspOSz + 1
'
				'		ELSE
				'		 STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
				'		 STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"&rShIJOSEnc(0) &"','1002')"
'
				'		set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				'		END IF
'
				'	  rShIJOSEnc.movenext

				  ' loop
				  ' call ofv.cerrarconsulta(rShIJOSEnc)
''
				'  CALL GrabarReg(aRecordSave2, idsex, idspOSz-1)

                else

                   fecha2 = CDATE( rs("FECIDENT") + CLNG(CDATE("1-1-1900")))
                   CContable2 = TRIM(rs("trcctaCble"))

                 ' end if
                  IF mes = MONTH(fecha2) AND CContable = CContable2 THEN    ' MIENTRAS ESTE EN EL MISMO MES LOS AGRUPO y ctacontable

                     IF REGVA THEN
                      totalmontoRepet = totalmontoRepet + CDBL(RS("TRCIMPBRUTO"))

                      ids(0, idspOS) = CDBL(rs(0))
                      ids(1, idspOS) = CDBL(RS("TRCIMPBRUTO"))
                      idspOS = idspOS + 1

                     SET rShIJOSEnc = getEncadenados(rs)

                      do while not rShIJOSEnc.eof

                        REGVAENC = VALIDACAMPOS(rShIJOSEnc)

                        fecha2X = CDATE( rShIJOSEnc("FECIDENT") + CLNG(CDATE("1-1-1900")))
                        CContable2X = TRIM(rShIJOSEnc("trcctaCble"))

                          IF REGVAENC THEN

                           IF mes = MONTH(fecha2X) AND CContable = CContable2X THEN

                           totalmontoRepet = totalmontoRepet + CDBL(rShIJOSEnc("TRCIMPBRUTO"))
						   ids(0, idspOS) = CDBL(rShIJOSEnc(0))
						   ids(1, idspOS) = CDBL(rShIJOSEnc("TRCIMPBRUTO"))
                           idspOS = idspOS + 1

                           ELSE
                           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
						   STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"&rShIJOSEnc(0) &"','1002')"

                         '  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                           END IF

                          END IF
                         rShIJOSEnc.movenext

                     loop
                      call ofv.cerrarconsulta(rShIJOSEnc)
                     end if

                  ELSE         ' ME FIJO QUE SUME > TOPEMONTO cambio de mes
                   '  primerReg = TRUE



                     IF totalmontoRepet >= TOPEMONTO THEN

                           CALL GrabarReg(aRecordSave, ids, idspOS -1,"2" ) 'se graban como repetitivos iguales dato pero no el monto


                     ELSE ' si no suman 1000 en el mes no van

                        FOR iz = 0 TO idspOS-1
                           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                           STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"& ids(0, iz) &"','1002')"

                        '   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                        NEXT

                     END IF

                     CALL guardarRs(aRecordSave,rs, 0)
                     idspOS = 0

                     ids(0, idspOS) = CDBL(rs(0))
                     ids(1, idspOS) = CDBL(RS("TRCIMPBRUTO"))
                     idspOS = idspOS + 1
                     totalmontoRepet = CDBL(RS("TRCIMPBRUTO"))
                     fecha1 = CDATE(rs("FECIDENT") + CLNG(CDATE("1-1-1900")))
					 mes = MONTH(fecha1)
					 CContable = TRIM(rs("trcctaCble"))

					SET rShIJOSEnc = getEncadenados(rs)

				    do while not rShIJOSEnc.eof

					 REGVAENC = VALIDACAMPOS(rShIJOSEnc)

					 fecha2X = CDATE( rShIJOSEnc("FECIDENT") + CLNG(CDATE("1-1-1900")))
					 CContable2X = TRIM(rShIJOSEnc("trcctaCble"))

					   IF REGVAENC THEN

						IF mes = MONTH(fecha2X) AND CContable = CContable2X THEN

						 totalmontoRepet = totalmontoRepet + CDBL(rShIJOSEnc("TRCIMPBRUTO"))
					     ids(0, idspOS) = CDBL(rShIJOSEnc(0))
					     ids(1, idspOS) = CDBL(rShIJOSEnc("TRCIMPBRUTO"))
						 idspOS = idspOS + 1

						ELSE
						 STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
					     STRSQL = STRSQL & " VALUES ('"&LOTE&"','REPETITIVOS','"&rShIJOSEnc(0) &"','1002')"

						''set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
						END IF

					   END IF
					  rShIJOSEnc.movenext

				   loop
                   call ofv.cerrarconsulta(rShIJOSEnc)





                 END IF

                end if
                end if

               ELSE ' ------eventos raiz

                SET rShIJOS = getEncadenados(rs)



                IF CDBL(RS("TRCIMPBRUTO"))  < TOPEMONTO AND rShIJOS.EOF THEN

                  STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
                  STRSQL = STRSQL & " VALUES ('"& LOTE &"','MONTO','"& rs(0) &"','1000')"
                ''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

                ELSEIF CDBL(RS("TRCIMPBRUTO"))  >= TOPEMONTO THEN

                  IF REGVA THEN

				    CALL Grabar(rs, secuencia, rShIJOS, STRIMESTRE) 'como normal rShIJOS esta eof

                  END IF


                ELSE   ' CDBL(RS("TRCIMPBRUTO"))  < TOPEMONTO AND NOT rShIJOS.EOF

              '   response.write ENCAMENORES
                 If  ENCAMENORES ="S" then
				    Dtotal= CDBL(RS("TRCIMPBRUTO") )

				  IF REGVA THEN
			         CALL grabar(rs, secuencia, rShIJOS, STRIMESTRE) 'como normal
				  End if

				 ELSE

				   Dtotal= CDBL(RS("TRCIMPBRUTO") )  + CDBL(RS("trcimpotros") )  + CDBL(RS("TRCIMPNOCTBLE") ) + CDBL(RS("TRCIMPACRECEN") )


				   if DTOTAL >= TOPEMONTO  then

                    SET rShIJOSc = getEncadenados(rs)
                   do while not rShIJOSc.eof
                     if rShIJOSc("envia") = "N" then
                         Dtotal=Dtotal - CDBL(rShIJOSc("TRCIMPBRUTO") )
                     end if
                     rShIJOSc.movenext
                   loop
                     call ofv.cerrarconsulta(rShIJOSc)
                   end if


				'   if rs(0) = 1956 then response.write "<br>" & Dtotal  & CDBL(RS("TRCIMPBRUTO") )  + CDBL(RS("trcimpotros") )  + CDBL(RS("TRCIMPNOCTBLE") ) &" " & CDBL(RS("TRCIMPBRUTO") )  &"---"& CDBL(RS("trcimpotros") ) &"---"& CDBL(RS("TRCIMPNOCTBLE") )

				   IF DTOTAL >= TOPEMONTO THEN
					  IF REGVA THEN
					    CALL grabar(rs, secuencia, rShIJOS, STRIMESTRE) 'como normal
				      End if

				   ELSE ' NO LLEGA AL TOPEMONTO Y ENCAMENORES=N
            '           if rs(0) = 1956 then response.write "<br>" & Dtotal  & CDBL(RS("TRCIMPBRUTO") )  + CDBL(RS("trcimpotros") )  + CDBL(RS("TRCIMPNOCTBLE") ) &" " & CDBL(RS("TRCIMPBRUTO") )  &"---"& CDBL(RS("trcimpotros") ) &"---"& CDBL(RS("TRCIMPNOCTBLE") )
				      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
				      STRSQL = STRSQL & " VALUES ('"& LOTE &"','MONTO','"& rs(0) &"','1000')"
				'''      set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
				    '  SET rShIJOS2 = GetEncadenados(rs)
				       CALL GrabarERRh(rShIJOS, 1000,"ENCADENADO")
				     ' CALL OFV.CERRARCONSULTA(rShIJOS2)

				   END IF

                  END IF

                END IF

               END IF

         END SELECT

      ELSE ' else de NO Envia


         IF RS("trccodtracerp")&"" ="" THEN 'SOLO SI ES RAIZ

          STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
          STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rs(0)&"','1001')"
        '''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
          SET rShIJOS2 = GetEncadenados(rs)
      '  IF RS(0)= 1 THEN RESPONSE.WRITE rShIJOS2.EOF
          CALL GrabarERRh(rShIJOS2, 1001,"ESTADO DEL EVENTO RAIZ")
          CALL OFV.CERRARCONSULTA(rShIJOS2)

         ELSE

          IF RS("TRCMARCAACTP")="N" THEN
          STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
		  STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rs(0)&"','1001')"
      '''    set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
          END IF
         END IF

      END IF




   END FUNCTION








   CALL ofv.cerrarconsulta(rs)
   CALL ofv.cerrarconn(cn)

END IF
'-------------------------------------------------------------

FUNCTION GrabarReg(aRecordSave, ids, idspOS,WWWV) 'se graban como repetitivos iguales dato pero no el monto

   FOR iz = 0 TO idspOS


      CARACTER = aRecordSave(53, 0)
      FECORIG = aRecordSave(3, 0)
      FECIDENT = aRecordSave(4, 0)
      FECCONTA = aRecordSave(5, 0)
      FECFINAL = aRecordSave(23, 0)
      TRCOBSERV = mid(aRecordSave(11, 0)&"",1,300)
      TIPOEVENTO = aRecordSave(55, 0)
      LINEGO2 = aRecordSave(52, 0)
      SUCURSAL = mid(aRecordSave(49, 0)&"",1,50)
      PROCESO = mid(aRecordSave(50, 0)&"",1,50)
      PRODSERV = aRecordSave(51, 0)
      TRCSECSOPORTE = mid(aRecordSave(27, 0)&"",1,50)
      TRCCTACBLE = mid(aRecordSave(29, 0)&"",1,6)
      TRCIMPSEGURO = aRecordSave(21, 0)
      RIESGOASOC = aRecordSave(56, 0)
   '    IF ids(0, iz)=1958 THEN 
''response.write ids(0, iz) &"    " & TRCCTACBLE & " -" & WWWV & " <br> "
      strsql = "insert into ORREGINFOBCRADET "
      strsql = strsql & "(CODLOTE,EVENTORAIZ,SECUENCIA,CARACTER,FECINICIO,FECDESCUBRIMIENTO,FECREGCONTABLE,FECFINAL,DESCRIPCION,TIPOEVENTO,LINEANEGOCIO,SUCURSAL,PROCESO,PRODUCTO,MONTO,PARTIDACONTABLE,COBERPOLIZA,VINCOTROCRIESGO,AREASOPORTE,OBJIDORTRACER,TRIMESTRE,ESTADO)"
      strsql = strsql & " values ("
      strsql = strsql & "'" & lote & "',"& EVENTORAIZ &",'-1','"& CARACTER &"','"& FECORIG &"','"& FECIDENT &"',"
      strsql = strsql & "'" & FECCONTA & "','"& FECFINAL &"','"& TRCOBSERV &"','"& TIPOEVENTO &"',"
      strsql = strsql & "'" & LINEGO2 & "','" & SUCURSAL & "','"& PROCESO &"','"& PRODSERV & "','"& REPLACE(ids(1, iz),",",".") &"',"
      IF CDBL("0" & TRCSECSOPORTE) > 0 THEN
         AREASOP = 1
      ELSE
         AREASOP = 0
      END IF
      strsql = strsql & "'" & TRCCTACBLE & "','"& TRCIMPSEGURO &"','"& RIESGOASOC&"','"&AREASOP &"',"
      Strsql = strsql & "'" & ids(0, iz) & "','"& strimestre &"','N')"

   '''  RESPONSE.WRITE STRSQL & "<BR>"

   ''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

XXXXXXXXXXXXXXXXXXXXX


      posi = posi + 1
      TeventosRepetitivos=TeventosRepetitivos + 1
   NEXT
   ULTEVENTORAIZ = EVENTORAIZ
   EVENTORAIZ = EVENTORAIZ + 1
END FUNCTION



'-------------------------------------------------------------

FUNCTION VerificarExtipo(rs, feciniT, fecfinT)

   XVA=FALSE

  IF (RS("trccodtracerp") & "" = "") THEN



      IF FCARGA = "S" THEN

       XVA = XVA OR ( (RS("FECREGIS") >= feciniT AND RS("FECREGIS") =< fecfinT))

      END IF

      IF FDESCUBRIMIENTO = "S" THEN
         XVA=XVA OR ((RS("FECIDENT") >= feciniT AND RS("FECIDENT") =< fecfinT) )
      END IF

      IF FORIGEN ="S" THEN
       XVA=XVA OR ((RS("FECORIG") >= feciniT AND RS("FECORIG") =< fecfinT) )
      END IF

      IF FCONTABLE ="S" THEN

       XVA=XVA OR ((RS("FECCONTA") >= feciniT AND RS("FECCONTA") =< fecfinT) )

      END IF


       IF FFINAL ="S" THEN

         XVA=XVA OR ((RS("FECFINENR") >= feciniT AND RS("FECFINENR") =< fecfinT) )

       END IF






      IF XVA THEN
         VerificarExtipo = "N" '& RS(0)

      ELSE
      ''   VerificarExtipo = "R" '& RS(0)
         VerificarExtipo = "N" '& RS(0)
      END IF
   ELSE
      VerificarExtipo = "R" '& RS(0)
   END IF



END FUNCTION

'------------------------------------------------------------------

FUNCTION guardarRs(Ar,Rse, Ix)

   FOR iz = 0 TO rse.fields.COUNT -1
      ' response.write iz & "<br>"
    '  aRecordSave(iz, Ix) = CSTR( rse(iz) & "")
       aR(iz, Ix) = CSTR( rse(iz) & "")
   NEXT

END FUNCTION



'------------------------------------------------------------------------------------------------

FUNCTION Grabar(rs, XEVENTORAIZ, rShIJOS, strimestre) 'como normal rShIJOS esta eof
   'Strsql = "Select *,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia From " & WKSETTABLA
  dim xsecuencia

  IF REGVA THEN
   strsql = "insert into ORREGINFOBCRADET "
   strsql = strsql & "(CODLOTE,EVENTORAIZ,SECUENCIA,CARACTER,FECINICIO,FECDESCUBRIMIENTO,FECREGCONTABLE,FECFINAL,DESCRIPCION,TIPOEVENTO,LINEANEGOCIO,SUCURSAL,PROCESO,PRODUCTO,MONTO,PARTIDACONTABLE,COBERPOLIZA,VINCOTROCRIESGO,AREASOPORTE,OBJIDORTRACER,TRIMESTRE,ESTADO)"
   strsql = strsql & " values ("
   strsql = strsql & "'" & lote & "',"& EVENTORAIZ &",0,'"& RS("CARACTER") &"','"& RS("FECORIG")&"','"&RS("FECIDENT")&"',"
   strsql = strsql & "'" & RS("FECCONTA") & "','"& RS("FECFINAL") &"','"& mid(RS("TRCOBSERV")&"",1,300)&"','"&RS("TIPOEVENTO")&"',"
   strsql = strsql & "'" & RS("LINEGO2") & "','" & mid(RS("SUCURSAL")&"",1,50) & "','"& mid(RS("PROCESO")&"",1,50) &"','"& mid(RS("PRODSERV")&"",1,50) &"','"&replace(RS("TRCIMPBRUTO"),",",".")&"',"

   IF CDBL("0" & RS("TRCSECSOPORTE")) > 0 THEN
      AREASOP = 1
   ELSE
      AREASOP = 0
   END IF
   strsql = strsql & "'" & mid(RS("TRCCTACBLE"),1,6) & "','"& RS("TRCIMPSEGURO") &"','"& RS("RIESGOASOC")&"','"&AREASOP &"',"
   Strsql = strsql & "'" & RS(0) & "','"& strimestre &"','N')"
'if rs(0)=323 then response.write "<br>" & strsql & "<br>"
 '''''''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

XXXXXXXXXXXXXXXXXXXXXXXX

  end if
   IF NOT RSHIJOS.EOF THEN
      XSECUENCIA = 1
      DO WHILE NOT RSHIJOS.EOF

         REGVA2 = VALIDACAMPOS(RSHIJOS)

         IF regva2 AND rshijos("envia") = "S" THEN

            strsql = "insert into ORREGINFOBCRADET "
            strsql = strsql & "(CODLOTE,EVENTORAIZ,SECUENCIA,CARACTER,FECINICIO,FECDESCUBRIMIENTO,FECREGCONTABLE,FECFINAL,DESCRIPCION,TIPOEVENTO,LINEANEGOCIO,SUCURSAL,PROCESO,PRODUCTO,MONTO,PARTIDACONTABLE,COBERPOLIZA,VINCOTROCRIESGO,AREASOPORTE,OBJIDORTRACER,TRIMESTRE,ESTADO)"
            strsql = strsql & " values ("
            strsql = strsql & "'" & lote & "',"& EVENTORAIZ &","&XSECUENCIA&",'"& RSHIJOS("CARACTER") &"','"& RSHIJOS("FECORIG")&"','"&RSHIJOS("FECIDENT")&"',"

            strsql = strsql & "'" & RSHIJOS("FECCONTA") & "','"& RSHIJOS("FECFINAL") &"','"& mid(RSHIJOS("TRCOBSERV")&"",1,300) &"','"&RSHIJOS("TIPOEVENTO")&"',"
            strsql = strsql & "'" & RSHIJOS("LINEGO2") & "','" & mid(RSHIJOS("SUCURSAL")&"",1,50) & "','"& mid(RSHIJOS("PROCESO")&"",1,50) &"','"& mid(RSHIJOS("PRODSERV")&"",1,50)&"','"& replace(RSHIJOS("TRCIMPBRUTO"),",",".")&"',"

            IF CDBL("0" & RS("TRCSECSOPORTE")) > 0 THEN
               AREASOP = 1
            ELSE
               AREASOP = 0
            END IF
            strsql = strsql & "'" & mid(RSHIJOS("TRCCTACBLE")&"",1,6) & "','"& RSHIJOS("TRCIMPSEGURO") &"','"& RSHIJOS("RIESGOASOC")&"','"&AREASOP &"',"
            Strsql = strsql & "'" & RSHIJOS(0) & "','"& strimestre &"','N')"
        ''''''''''    set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
            XSECUENCIA = XSECUENCIA + 1
            posi = posi + 1
            TeventosEncadenados=TeventosEncadenados + 1
         ELSE
            if rshijos("envia") <> "S" then
             STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
             STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rshijos(0)&"','1001')"
            end if
            REGVA = FALSE
          '''''''''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

         END IF
         RSHIJOS.MOVENEXT


      LOOP



   END IF
   ULTEVENTORAIZ = EVENTORAIZ
   EVENTORAIZ = EVENTORAIZ + 1
END FUNCTION

'-----------------------------------------------------------------

FUNCTION GrabarERRh(rShIJOSZ, ERROR,estado) 'como normal rShIJOS esta eof
   'Strsql = "Select *,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia From " & WKSETTABLA

  ' IF RS(0)=1956 THEN RESPONSE.WRITE RSHIJOSZ.EOF & "<BR><BR><BR><BR><BR><BR><BR><BR><BR><BR>"
   DO WHILE NOT RSHIJOSZ.EOF

 '     IF rshijosZ(0)=1898 THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','"&estado&"','"&rshijosZ(0)&"','"& ERROR &"')"
      REGVA = FALSE
 ''''''''''     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

      IF rShIJOSZ("ENVIA") <> "S" THEN
       STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
	   STRSQL = STRSQL & " VALUES ('"&LOTE&"','ESTADO','"&rshijosZ(0)&"','1001')"
	   REGVA = FALSE
   ''''''''     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
      END IF



   '   RESPONSE.WRITE STRSQL & "<BR>"
   '   END IF
      xxx=ValidaCamposH(RSHIJOSZ)
      RSHIJOSZ.MOVENEXT


   LOOP



END FUNCTION


'----------------------------------------------------------------

FUNCTION ValidaCampos(rs)
   DIM BREGVA

   BREGVA = TRUE



 select case rs("CARACTER")


 case "05"

      if RS("FECORIG") > 0 then

       IF RS("FECIDENT") <  RS("FECORIG") THEN
           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
           STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
           '     RESPONSE.WRITE STRSQL
            BREGVA = FALSE
      ''''''''     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        END IF
       end if


       if RS("FECFINAL") > 0 then
        IF RS("FECFINAL") < RS("FECIDENT") THEN
           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
           STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
           '    RESPONSE.WRITE STRSQL
            BREGVA = FALSE
     '''''''      set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        END IF
        end if




      IF NOT ISNUMERIC (rs("CARACTER")) THEN
         STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
         STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
         '     RESPONSE.WRITE STRSQL
         BREGVA = FALSE
    '''''''     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
      END IF



     IF TRIM(rs("trcctaCble")&"") ="" then

       STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
        STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','10')"
        '     RESPONSE.WRITE STRSQL
        BREGVA = FALSE
     '''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


    END IF



 case else
    if RS("FECORIG") > 0 then

    IF RS("FECIDENT") <  RS("FECORIG") THEN
        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
        STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
        '     RESPONSE.WRITE STRSQL
         BREGVA = FALSE
     '''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
     END IF
    end if

     if RS("FECFINAL") > 0 then
     IF RS("FECFINAL") < RS("FECIDENT") THEN
        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
        STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
        '    RESPONSE.WRITE STRSQL
         BREGVA = FALSE
     ''''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
     END IF
     end if



   IF NOT ISNUMERIC(RS("TRCIMPSEGURO")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','COBERTURA DE SEGURO','"&rs(0)&"','5')"
      '    RESPONSE.WRITE STRSQL
  '''''''    set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
      BREGVA = FALSE
   END IF

   IF FALSE THEN
  ' IF BUSCAEXCAR(RS("trcobserv"), "") THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','DESCRIPCION','"&rs(0)&"','10')"
      '    RESPONSE.WRITE STRSQL
      BREGVA = FALSE
    ''''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

     IF RS("SUCURSAL") & "" = "" THEN
   'IF BUSCAEXCAR(RS("SUCURSAL"), "") OR RS("SUCURSAL") & "" = ""  THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','SUCURSAL','"&rs(0)&"','10')"
      '    RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   ''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

  IF RS("PRODSERV") &""= "" THEN
  ' IF BUSCAEXCAR(RS("PRODSERV"), "") OR RS("PRODSERV") &""= "" THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','PRODUCTO','"&rs(0)&"','10')"
      '    RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   '''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

   IF RS("PROCESO")& "" =""   THEN
 '  IF BUSCAEXCAR(RS("PROCESO"), "") OR RS("PROCESO")& "" =""  THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','PROCESO','"&rs(0)&"','10')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   '''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF



   IF NOT ISNUMERIC (rs("LINEGO2")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','LINEA DE NEGOCIO','"&rs(0)&"','09')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   '''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF


   IF NOT ISNUMERIC (rs("CARACTER")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   '''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF



   IF NOT ISNUMERIC (rs("TIPOEVENTO")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','TIPO DE EVENTO','"&rs(0)&"','09')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   ''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

  if  rs("CARACTER")="01" or rs("CARACTER")="03" or rs("CARACTER")="05" or rs("CARACTER")="06" or rs("CARACTER")="07" then

  IF TRIM(rs("trcctaCble")&"")="" then

    STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
     STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
     '     RESPONSE.WRITE STRSQL
     BREGVA = FALSE
  ''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

  end if
  END IF


   if  rs("CARACTER")="02" or rs("CARACTER")="04" or rs("CARACTER")="08" or rs("CARACTER")="09"  then

   IF TRIM(rs("trcctaCble")&"") <> ""   then

     STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   ''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

   end if


   IF  rs("FECCONTA") > 0  then

     STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
     STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA CONTABLE','"&rs(0)&"','18')"
         '     RESPONSE.WRITE STRSQL
     BREGVA = FALSE
   ''''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

   end if

   END IF

 '  IF NOT ISNUMERIC (rs("RIESGOASOC")) THEN

 '        STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
 ''        STRSQL = STRSQL & " VALUES ('"&LOTE&"','RIESGO ASOCIADO','"&rs(0)&"','09')"
 '        '     RESPONSE.WRITE STRSQL
 '        BREGVA = FALSE
 '        set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
 '  END IF

end select

   ValidaCampos = BREGVA

   if BREGVA=false then
    SET rShIJOS3 = GetEncadenados(rs)
    CALL GrabarERRh(rShIJOS3, 1003,"ENCADENADO")
    CALL OFV.CERRARCONSULTA(rShIJOS3)
   end if

END FUNCTION

FUNCTION ValidaCamposH(rs)
   DIM BREGVA


    BREGVA=true


  select case rs("CARACTER")


 case "05"

      if RS("FECORIG") > 0 then

       IF RS("FECIDENT") <  RS("FECORIG") THEN
           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
           STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
           '     RESPONSE.WRITE STRSQL
            BREGVA = FALSE
         ''''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        END IF
       end if


       if RS("FECFINAL") > 0 then
        IF RS("FECFINAL") < RS("FECIDENT") THEN
           STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
           STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
           '    RESPONSE.WRITE STRSQL
            BREGVA = FALSE
       ''''''    set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
        END IF
        end if




      IF NOT ISNUMERIC (rs("CARACTER")) THEN
         STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
         STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
         '     RESPONSE.WRITE STRSQL
         BREGVA = FALSE
    '''''''     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
      END IF



     IF TRIM(rs("trcctaCble")&"") ="" then

       STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
        STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','10')"
        '     RESPONSE.WRITE STRSQL
        BREGVA = FALSE
     '''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


    END IF



 case else

  if RS("FECORIG") > 0 then

      IF RS("FECIDENT") <  RS("FECORIG") THEN
          STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
          STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DESCUBRIMIENTO','"&rs(0)&"','14')"
          '     RESPONSE.WRITE STRSQL
           BREGVA = FALSE
        '''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
       END IF
      end if

      if RS("FECFINAL") > 0 then
       IF RS("FECFINAL") < RS("FECIDENT") THEN
          STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
          STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA DE FINALIZACION','"&rs(0)&"','14')"
          '    RESPONSE.WRITE STRSQL
           BREGVA = FALSE
       ''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
       END IF
     end if


   IF NOT ISNUMERIC(RS("TRCIMPSEGURO")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','COBERTURA DE SEGURO','"&rs(0)&"','5')"
      '    RESPONSE.WRITE STRSQL
    ''''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
       BREGVA = FALSE
   END IF


   IF FALSE THEN
'   IF BUSCAEXCAR(RS("trcobserv"), "") THEN



      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','DESCRIPCION','"&rs(0)&"','10')"
      '    RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   '''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

   IF RS("SUCURSAL") & "" = "" THEN
  ' IF BUSCAEXCAR(RS("SUCURSAL"), "") OR RS("SUCURSAL") & "" = ""  THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','SUCURSAL','"&rs(0)&"','10')"
      '    RESPONSE.WRITE STRSQL
       BREGVA = FALSE
    '''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

    IF RS("PRODSERV") &""= "" THEN
   'IF BUSCAEXCAR(RS("PRODSERV"), "") OR RS("PRODSERV") &""= "" THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','PRODUCTO','"&rs(0)&"','10')"
      '    RESPONSE.WRITE STRSQL
       BREGVA = FALSE
   '''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF

   IF RS("PROCESO")& "" =""  THEN
   'IF BUSCAEXCAR(RS("PROCESO"), "") OR RS("PROCESO")& "" =""  THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','PROCESO','"&rs(0)&"','10')"
      '     RESPONSE.WRITE STRSQL
        BREGVA = FALSE
    ''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF



   IF NOT ISNUMERIC (rs("LINEGO2")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','LINEA DE NEGOCIO','"&rs(0)&"','09')"
      '     RESPONSE.WRITE STRSQL
       BREGVA = FALSE
    '''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF


   IF NOT ISNUMERIC (rs("CARACTER")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','CARACTER','"&rs(0)&"','09')"
      '     RESPONSE.WRITE STRSQL
       BREGVA = FALSE
    '''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF



   IF NOT ISNUMERIC (rs("TIPOEVENTO")) THEN
      STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','TIPO DE EVENTO','"&rs(0)&"','09')"
      '     RESPONSE.WRITE STRSQL
       BREGVA = FALSE
   ''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
   END IF


 if  rs("CARACTER")="01" or rs("CARACTER")="03" or rs("CARACTER")="05" or rs("CARACTER")="06" or rs("CARACTER")="07" then

   IF TRIM(rs("trcctaCble")&"") = ""  then

     STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
    ''''''  set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

  end if
END IF

if  rs("CARACTER")="02" or rs("CARACTER")="04" or rs("CARACTER")="08" or rs("CARACTER")="09"  then

   IF TRIM(rs("trcctaCble")&"") <> ""   then

     STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
      STRSQL = STRSQL & " VALUES ('"&LOTE&"','CUENTA CONTABLE','"&rs(0)&"','18')"
      '     RESPONSE.WRITE STRSQL
      BREGVA = FALSE
   '''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

   end if


   IF  rs("FECCONTA") > 0  then

     STRSQL = "INSERT INTO ORRIERRLOTE (CODLOTE,COLUMNA,nroevento,NROERROR)"
     STRSQL = STRSQL & " VALUES ('"&LOTE&"','FECHA CONTABLE','"&rs(0)&"','18')"
         '     RESPONSE.WRITE STRSQL
     BREGVA = FALSE
  ''''''''   set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

   end if

   END IF

END SELECT
ValidaCamposh=BREGVA
END FUNCTION


function GetEncadenados(rs)


IF (session("APLMDLOR") & "") = "BE" THEN
   Strsql = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRIpcion AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
 ELSE
   Strsql = "Select ORTRACER.*,E.descri as Sucursal,O.DESCRI AS PROCESO,OT.DESCRI AS PRODSERV,OL2.CODBCRA AS LINEGO2 ,OC.CODBCRA AS CARACTER,orestados.envia,OTCAT.CODBCRA AS TIPOEVENTO,ORAS.CODBCRA AS RIESGOASOC  From " & WKSETTABLA
END IF


 'strsql= strsql & " inner join orestados on (ORTRACER.trcestado=orestados.codigo and orestados.envia='S')"
 strsql = strsql & " inner join orestados on (ORTRACER.trcestado=orestados.codigo )"
 strsql = strsql & " left join eco E on (ORTRACER.trcecoresp=E.eco)  "


 IF (session("APLMDLOR") & "") = "BE" THEN
	STRSQL = STRSQL & " LEFT JOIN TABPROCESOS01 O ON (ORTRACER.trcPROCESO=O.OBJID)"
 ELSE
	STRSQL = STRSQL & " LEFT JOIN OBJETOS O ON (ORTRACER.trcPROCESO=O.CODOBJETO)"
 END IF


 strsql = strsql & " left join ORTIPOPRODUCTO OT on (ORTRACER.trcPRODUCTO=OT.OBJID)  "
 strsql = strsql & " left join ORTIPOCATEGO OTCAT on (ORTRACER.TRCEVENTORIESGO2 = OTCAT.OBJID)  "
 strsql = strsql & " left join ORRIESGOASOC ORAS on (ORTRACER.TRCRIESGOASOC = ORAS.OBJID)  "

 strsql = strsql & " left join ORLINNEGO2 OL2 on (ORTRACER.trcNEGOLINE=OL2.OBJID)  "
 strsql = strsql & " left join ORTIPOCARACTER OC on (ORTRACER.trcCARACTER=OC.OBJID)  "
 strsql = strsql & " WHERE ORTRACER.TRCCODTRACERP = '" & (RS("CODTRACER") & "") & "' AND ("



  STRSQL = STRSQL & SESSION(WSPSESION) '& SORFECINIHW1

  SORFECINIHW = " (TRCFECACTP > = " & feciniT & " and TRCFECACTP < = "& fecfinT & ") "
  STRSQL = STRSQL & "  OR " & SORFECINIHW
  STRSQL = STRSQL & ") ORDER BY ORTRACER.FECIDENT,ORTRACER.FECREGIS,oc.cODBCRA,ORTRACER.TRCCTACBLE"

  set GetEncadenados= ofv.crearconsultaEx(StrSql, cn, 1, parametros)


   do while not GetEncadenados.eof
      posiX=posiX + 1
      RESPONSE.WRITE GetEncadenados(0) & "--- (E)<BR>"
      GetEncadenados.movenext
   loop

  set GetEncadenados= ofv.crearconsultaEx(StrSql, cn, 1, parametros)

end function




%>
