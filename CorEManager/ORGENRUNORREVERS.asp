<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sopcion  = request.querystring("opcion") & ""
  lote     = request.querystring("lote")

  APLCHECK()

  FILA = ""
  set cn = ofv.conectar(ofv.strconn2)

  POSI = 0

  WWWFEC = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))

  if sOPCION="FIN" then


     STRSQL = "SELECT TRIMESTRE,ANO FROM ORREGINFOBCRA "
     STRSQL = STRSQL &" WHERE CODLOTE = '" & LOTE & "'"
  ' RESPONSE.WRITE "<BR>" & STRSQL
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF NOT RS.EOF THEN

        IF (RS(1) & "") = "2010" AND (RS(0) & "") = "1" THEN
           WWWFEC = "40296"
        END IF


     END IF


     STRSQL = " SELECT C.OBJID AS KEYTRACER,C.TRCMARCASEND,C.TRCMARCAACTP,C.TRCMARCAACTH,C.TRCFECACTP,C.TRCFECACTH,B.* "
     STRSQL = STRSQL & " FROM ORREGINFOBCRADET B  "
     STRSQL = STRSQL & " inner JOIN ORTRACER C  ON (B.OBJIDORTRACER=C.OBJID) "
     STRSQL = STRSQL & " WHERE B.CODLOTE='" &  LOTE & "' ORDER BY B.OBJIDORTRACER"
 '  RESPONSE.WRITE STRSQL
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     DO WHILE Not rs.EOF 


        strsqlx = "insert into ORTRACERSEND ("
        strsqlx = strsqlx & " OBJIDORTRACER,CODLOTE,TIPOENVIO,TRCMARCASEND,TRCMARCAACTP,"
        strsqlx = strsqlx & " TRCMARCAACTH,TRCFECACTP,TRCFECACTH"
        strsqlx = strsqlx & " ) values ("
        strsqlx = strsqlx & " '" & RS(0) & "','" & RS(7) & "','" & "N" & "','" & RS(1) & "','" & RS(2) & "', "
        strsqlx = strsqlx & " '" & RS(3) & "','" & RS(4) & "','" & RS(5) & "' "
        strsqlx = strsqlx & " ) "
      'response.write strsqlx

        SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)


        strsqlx = "UPDATE ORTRACER SET "
        strsqlx = strsqlx & " TRCMARCASEND = 'S',"
        strsqlx = strsqlx & " TRCMARCAACTP = 'S',"
        strsqlx = strsqlx & " TRCMARCAACTH = 'S',"
        strsqlx = strsqlx & " TRCFECACTP   = '" & WWWFEC & "',"
        strsqlx = strsqlx & " TRCFECACTH   = '" & WWWFEC & "'"
        strsqlx = strsqlx & " WHERE OBJID  = " &  RS(0) & " "

      'response.write strsqlx

        SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)



        RS.MOVENEXT
        posI=posI+1

     LOOP



     STRSQL = " SELECT C.OBJID AS KEYTRACER,C.TRCMARCASEND,C.TRCMARCAACTP,C.TRCMARCAACTH,C.TRCFECACTP,C.TRCFECACTH "
     STRSQL = STRSQL & " FROM ORTRACER C  "
     STRSQL = STRSQL & " WHERE C.TRCMARCAACTP in ('K','X','KR','XR') ORDER BY C.OBJID"
 '  RESPONSE.WRITE STRSQL
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     DO WHILE Not rs.EOF 

       WWWMARCARST = "N"

       IF (RS(2) & "") = "KR" OR (RS(2) & "") = "XR" THEN
          WWWMARCARST = "R"
       END IF 


        strsqlx = "insert into ORTRACERSEND ("
        strsqlx = strsqlx & " OBJIDORTRACER,CODLOTE,TIPOENVIO,TRCMARCASEND,TRCMARCAACTP,"
        strsqlx = strsqlx & " TRCMARCAACTH,TRCFECACTP,TRCFECACTH"
        strsqlx = strsqlx & " ) values ("
        strsqlx = strsqlx & " '" & RS(0) & "','" & LOTE & "','" & RS(2) & "','" & RS(1) & "','" & RS(2) & "', "
        strsqlx = strsqlx & " '" & RS(3) & "','" & RS(4) & "','" & RS(5) & "' "
        strsqlx = strsqlx & " ) "
      'response.write strsqlx

        SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)


        strsqlx = "UPDATE ORTRACER SET "
        strsqlx = strsqlx & " TRCMARCASEND = '" & RS(1) & "',"
        strsqlx = strsqlx & " TRCMARCAACTP = '" & WWWMARCARST & "',"
        strsqlx = strsqlx & " TRCMARCAACTH = 'N',"
        strsqlx = strsqlx & " TRCFECACTP   = '" & WWWFEC & "',"
        strsqlx = strsqlx & " TRCFECACTH   = '" & WWWFEC & "'"
        strsqlx = strsqlx & " WHERE OBJID  = " &  RS(0) & " "

      'response.write strsqlx

        SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)



        RS.MOVENEXT
        posI=posI+1

     LOOP



     STRSQL = "UPDATE ORREGINFOBCRA SET "
     STRSQL = STRSQL & " ESTADO = 4 "
     STRSQL = STRSQL &" WHERE CODLOTE = '" & LOTE & "'"
  ' RESPONSE.WRITE "<BR>" & STRSQL
     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)




  ELSE


     STRSQL = " SELECT * FROM ORTRACERSEND  "
     STRSQL = STRSQL & " WHERE CODLOTE = '" &  LOTE & "' ORDER BY OBJIDORTRACER"
 '  RESPONSE.WRITE STRSQL
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     DO WHILE Not rs.EOF 


        strsqlx = "UPDATE ORTRACER SET "
        strsqlx = strsqlx & " TRCMARCASEND = '" & RS(4) & "',"
        strsqlx = strsqlx & " TRCMARCAACTP = '" & RS(5) & "',"
        strsqlx = strsqlx & " TRCMARCAACTH = '" & RS(6) & "',"
        strsqlx = strsqlx & " TRCFECACTP   = '" & RS(7) & "',"
        strsqlx = strsqlx & " TRCFECACTH   = '" & RS(8) & "'"
        strsqlx = strsqlx & " WHERE OBJID  = " &  RS(1) & " "

      'response.write strsqlx

        SET xxxx = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)



        RS.MOVENEXT
        posI=posI+1

     LOOP

     STRSQL = " DELETE FROM ORTRACERSEND  "
     STRSQL = STRSQL & " WHERE CODLOTE = '" &  LOTE & "' "
  ' RESPONSE.WRITE "<BR>" & STRSQL
     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)


     STRSQL = "DELETE FROM ORRIERRLOTE      WHERE CODLOTE = '" & LOTE & "' "
     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

     STRSQL = "DELETE FROM ORREGINFOBCRADET WHERE codLOTE = '" & LOTE & "' "
'   response.write STRSQL & "<br>"
     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, volver)



     STRSQL = "UPDATE ORREGINFOBCRA SET "
     STRSQL = STRSQL & " TOTAL           = 0"
     STRSQL = STRSQL & " ,totalMontoTri   = 0"
     STRSQL = STRSQL & " ,TOTALINFORMAR   = 0"
     STRSQL = STRSQL & " ,TOTALNOINFORMAR = 0"
     STRSQL = STRSQL & " ,totrep          = 0"
     STRSQL = STRSQL & " ,totenca         = 0"
     STRSQL = STRSQL & " ,totdifrep       = 0"
     STRSQL = STRSQL & " ,MONTOTOTALNOINF = 0"
     STRSQL = STRSQL & " ,ESTADO          = 1"
     STRSQL = STRSQL & " WHERE CODLOTE = '" & LOTE & "'"
  ' RESPONSE.WRITE "<BR>" & STRSQL
     set XXXC = ofv.crearconsultaEx(StrSql, cn, 1, parametros)





  end if


  RESPONSE.REDIRECT "ORGENRUNOR.ASP"


  call ofv.cerrarconn(cn)





end if


%>