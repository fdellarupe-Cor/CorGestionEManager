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
   server.scripttimeout=60*20
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


   Strsql = "Select ORTRACER.objid, ortracer.codtracer, ortracer.trccodtracerp  From " & WKSETTABLA

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
   swhere = swhere & "  AND (ortracer.trccodtracerp IS NULL OR  ortracer.trccodtracerp = '') and ( "



   IF session(WSPSESION ) <> "" THEN
      swhere = swhere & "   (ortracer.TRCMARCASEND <> 'S' AND  (" & SESSION(WSPSESION) & ") ) 	or "
   END IF


   STRSQL = STRSQL & swhere
   strsql = strsql & "  (ortracer.TRCMARCAACTP <> 'S' AND (ORTRACER.TRCFECACTP > = " & feciniT & " and ORTRACER.TRCFECACTP < = "& fecfinT & "))"
   strsql = strsql & " 	or (ortracer.TRCMARCAACTH <> 'S' AND (ORTRACER.TRCFECACTH > = " & feciniT & " and ORTRACER.TRCFECACTH < = "& fecfinT & "))"
   strsql = strsql & " 	) "

   STRSQL = STRSQL & " ORDER BY ORTRACER.objid "
 


  '''  RESPONSE.WRITE STRSQL & "<BR><BR><BR><BR>"

   Formulario = ""

   posi=0

   RESPONSE.WRITE "<BR>-------------------------------------------------------<BR>&nbsp;LOTE&nbsp;" & nrolote
   RESPONSE.WRITE "<BR>-------------------------------------------------------<BR>"

   SET RS = OFV.crearconsultaEx(StrSql, cn, 1, volver)

   do while not rs.eof

      SORFECINIHW = " (ORTRACER.TRCFECACTP > = " & feciniT & " and ORTRACER.TRCFECACTP < = "& fecfinT & ") "


      posi=posi + 1

      wwwevento = rs(0)

      for yxx=1 to (6-len(wwwevento)) 

          wwwevento = "." & wwwevento
      next

      RESPONSE.WRITE wwwevento & "&nbsp;----- (R)<BR>"


      Strsql = "Select ORTRACER.objid, b.objid  From ortracer "
      strsql = strsql & " inner join ortracer b on ( ORTRACER.TRCCODTRACERP = b.codtracer ) "
      strsql = strsql & " WHERE b.CODTRACER = '" & (RS(1) & "") & "' AND ("
      STRSQL = STRSQL & SESSION(WSPSESION) 
      STRSQL = STRSQL & "  OR " & SORFECINIHW
      STRSQL = STRSQL & ") ORDER BY ORTRACER.objid"
 '''   RESPONSE.WRITE STRSQL & "<BR><BR><BR><BR>"
      set GetEncadenados= ofv.crearconsultaEx(StrSql, cn, 1, parametros)
      do while not GetEncadenados.eof
         posi=posi + 1

         wwwevento = GetEncadenados(0)

         for yxx=1 to (6-len(wwwevento)) 

             wwwevento = "." & wwwevento
         next

         RESPONSE.WRITE wwwevento & "&nbsp;--------- (E: " & GetEncadenados(1) & ")<BR>"
         GetEncadenados.movenext
      loop


    rs.movenext
   loop

    RESPONSE.WRITE "<BR>------------------------------------------------------------------<BR>TOTAL&nbsp;LOTE&nbsp;" & nrolote & ":&nbsp;" & posi
    RESPONSE.WRITE "<BR>------------------------------------------------------------------<BR>"

 end if

   






%>
