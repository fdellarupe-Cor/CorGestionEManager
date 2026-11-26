<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

  response.expires=0

  CantRegMoverx = request.querystring("Cantregmover")
  ID = request.querystring("ID")

  set cn = ofv.conectar(ofv.strconn2)

nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if



  Strsql = "Select * From EMENCUESTAS where objid = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

         xnovcod = "NEC" & cstr(rs(0))


        dFechai = rs(3)
        dFechae = rs(4)
        xhora = time()
        xmsgusr = rs(6)
        xtipo = 51
        if clng(rs(8)) = 2 then xtipo = 2


        strsql = "Insert into novedades ("
        strsql = strsql & "codobjeto,tipobjeto,codigo,tipo,"
        strsql = strsql & "fecinicio,fecexp,mensaje,importancia,"
        strsql = strsql & "relacion,workflowid,contacto,hora,"
        strsql = strsql & "MSGVERSION,MSGTITULO,MSGCOPETE,MSGIMAGEN,"
        strsql = strsql & " MSGURL,MSGURLDSC,"
        strsql = strsql & " MSGVAENC,MSGVAFORO,MSGCODENC,MSGCODFORO"
        strsql = strsql & ") "
        strsql = strsql & " values ("
        strsql = strsql & "0,' ','" & xnovcod & "'," & xtipo & ","
        strsql = strsql & "" & dFechai & "," & dFechae & ",'" & REPLACE(xmsgusr & "","'","''") & "'," & REPLACE(rs(7) & "","'","''") & ","
        strsql = strsql & "0,0,'" & rs(13) & "','" & xhora & "',"
        strsql = strsql & "'2','" & REPLACE(rs(17) & "","'","''") & "','" & REPLACE(rs(18) & "","'","''") & "','" & REPLACE(rs(19) & "","'","''") & "',"
        strsql = strsql & "'" & REPLACE(rs(20) & "","'","''") & "','" & REPLACE(rs(21) & "","'","''") & "',"
        strsql = strsql & "'S','N','" & rs(0) & "','0'"
        strsql = strsql & ")"

      '  RESPONSE.WRITE STRSQL & "<BR>"

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        strsql = "update EMENCUESTAS set  ENCHORA = 'S' where objid = " & ID
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        strsql = "update EMENCUESTASdist set fecexp = " & dFechae & " where codigo = " & ID
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


      IF (rs(19) & "") = "1" THEN


        Strsql = "Select * From EMENCUESTASIMG  where ENCcodigo = " & ID & " ORDER BY IMGsecuen, OBJID"
        set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        DO WHILE NOT rs1.eof 

           strsql = "Insert into NOVEDADESIMG ("
           strsql = strsql & "NOVEDAD,SECUEN,TIPO,DIRECTORIO,IMAGEN"
           strsql = strsql & ") "
           strsql = strsql & " values ("
           strsql = strsql & "'" & xnovcod & "','" & REPLACE(rs1(2) & "","'","''") & "','" & REPLACE(rs1(3) & "","'","''") & "','" & REPLACE(rs1(4) & "","'","''") & "','" & REPLACE(rs1(5) & "","'","''") & "'"
           strsql = strsql & ")"

         '  RESPONSE.WRITE STRSQL & "<BR>"

           call ofv.crearconsultaEx(StrSql,cn,1,parametros)


         RS1.MOVENEXT
        LOOP

       END IF





     if clng(rs(8)) = 2 then

         Strsql = "Select * From EMENCUESTASDIST  where codigo = " & ID
         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rsx.eof then

            do until rsx.eof
               xusr = rsx(2)
               xsect = rsx(3)


               strsql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
               strsql = strsql & " values ('" & xnovcod & "','" & xusr & "'," & xsect & "," & dFechae & ",'" & xhora & "')"

               call ofv.crearconsultaEx(StrSql,cn,1,parametros)

            rsx.movenext
            loop
         end if
         call ofv.cerrarconsulta(rsx)


     else

               strsql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
               strsql = strsql & " values ('" & xnovcod & "','TODOS',999," & dFechae & ",'" & xhora & "')"

               call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     end if


  end if





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

       response.redirect "AdENCUESTAS.asp?CantRegMover=" & CantRegMoverx




%>