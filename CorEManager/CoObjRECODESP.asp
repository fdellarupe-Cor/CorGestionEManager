<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%
Dim Scodobj,stipobj

ok=ofv.CheckUsuario()

  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  ID      = request.querystring("mensaje")

  set cn = ofv.conectar(ofv.strconn2)

nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if



  Strsql = "Select * From novegenpar  where objid = " & ID
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then



         xnovcod = "RGP" & cstr(rs(0))


        dFechai = rs(1)
        dFechae = clng(rs(1)) + clng(rs(3))
        xhora   = time()
        xmsgusr = rs(4)
        xtipo   = rs(19)
        xtexto =rs("texto")
        XUSRORG=RS("CONTACTO")
        Scodobj="0" & rs("msgobjcod")
        stipobj=rs("msgobjtip")&""


       if rs(8) <> "S" then

        strsql = "Insert into novedades ("
        strsql = strsql & "codobjeto,tipobjeto,codigo,tipo,"
        strsql = strsql & "fecinicio,fecexp,mensaje,importancia,"
        strsql = strsql & "relacion,workflowid,contacto,hora,"
        strsql = strsql & "MSGVERSION,MSGTITULO,MSGCOPETE,MSGIMAGEN,"
        strsql = strsql & " MSGURL,MSGURLDSC,"
        strsql = strsql & " MSGVAENC,MSGVAFORO,MSGCODENC,MSGCODFORO"
        strsql = strsql & ") "
        strsql = strsql & " values ("
        strsql = strsql & "'" & nCodigo & "','" & sTipo & "','" & xnovcod & "'," & xtipo & ","
        strsql = strsql & "" & dFechai & "," & dFechae & ",'" & REPLACE(xmsgusr & "","'","''") & "'," & REPLACE(rs(5) & "","'","''") & ","
        strsql = strsql & "0,0,'" & rs(7) & "','" & xhora & "',"
        strsql = strsql & "'" & rs(9) & "','" & REPLACE(rs(10) & "","'","''") & "','" & REPLACE(rs(11) & "","'","''") & "','" & REPLACE(rs(12) & "","'","''") & "',"
        strsql = strsql & "'" & REPLACE(rs(13) & "","'","''") & "','" & REPLACE(rs(14) & "","'","''") & "',"
        strsql = strsql & "'" & rs(15) & "','" & rs(16) & "','" & rs(17) & "','" & rs(18) & "'"
        strsql = strsql & ")"

      '  RESPONSE.WRITE STRSQL & "<BR>"

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        strsql = "update novegenpar set fecexp = " & dFechae & ", enviado = 'S' where objid = " & ID
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

        strsql = "update novegenpardist set fecexp = " & dFechae & " where codigo = " & ID
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)





      IF (rs(12) & "") = "1" THEN


        Strsql = "Select * From novegenparIMG  where NOVEGENPAR = '" & ID & "' ORDER BY SECUEN, OBJID"
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


       else


          Strsql = "Select * From novedades  where codigo  = '" & xnovcod & "' "
          set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsx.eof then

             xhora   = rsx(12) & ""

          end if


       end if


     if clng(rs(6)) = 2 then

         Strsql = "Select * From novegenpardist  where codigo = " & ID
        ' response.write strsql
         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rsx.eof then
            DELIM=""
            USRMAILS=""
            do until rsx.eof
               xusr = rsx(2)
               xsect = rsx(3)

              if rs(8) <> "S" or (rs(8) = "S" and (rsx(4) & "") = "0") then

                 USRMAILS= USRMAILS & DELIM & xusr

                 strsql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
                 strsql = strsql & " values ('" & xnovcod & "','" & xusr & "'," & xsect & "," & dFechae & ",'" & xhora & "')"

                 call ofv.crearconsultaEx(StrSql,cn,1,parametros)

               if (rs(8) = "S" and (rsx(4) & "") = "0") then

                  strsql = "update novegenpardist set fecexp = " & dFechae & " where objid = " & rsx(0)
                  call ofv.crearconsultaEx(StrSql,cn,1,parametros)

               end if

               DELIM=CHR(9)

              end if

              ' RESPONSE.WRITE xusr & "-----" & xtexto & "-------" & USRMAILS
            rsx.movenext

            loop

           if USRMAILS <> "" then
              xxxx = GENCKTMAILEX("USR", USRMAILS,  xtexto , XUSRORG, "")
           end if

         end if

         call ofv.cerrarconsulta(rsx)


     else

      if rs(8) <> "S" then

               strsql = "Insert into novedadesdist (codigo,usugrp,sector,fecexp,hora) "
               strsql = strsql & " values ('" & xnovcod & "','TODOS',999," & dFechae & ",'" & xhora & "')"

               call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          end if

     end if


  end if



Function GENCKTMAILEX(WETIPO, WECOD, WEMESSAGE, WEORIGEN, WELVL)


    XRESPO = WECOD & ""

    xxxx = getInfoUser(WEORIGEN)
    If session("SGU_mail") <> "" Then
     WEORIGENMAIL = session("SGU_mail")
    ELSE
     XXX = getInfoUser("COREMANAGER")
     WEORIGENMAIL = session("SGU_mail")
    END IF

    IF WEORIGENMAIL="" THEN WEORIGENMAIL="COR@MANAGER.COM"

    If XRESPO <> "" And XRESPO <> "NO" Then

        Set cn2 = ofv.conectar(ofv.strconn0)

        XPARA = GETRESPOCKT(WEORIGEN, WETIPO)


        XGUMENSAGEw = WEMESSAGE

        XMESAURL = ""


        IF Session("CORUNQ") = "S" THEN

           XMESAURL = getobjurl(nCodigo,sTipo)

        ELSE

           If session("URLTYP") = "V" Then

              XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))

           End If


           If session("URLPORT") = "S" Then
             ' XMESAURL = "http://" & request.servervariables("SERVER_NAME") & XMESAURL

               XMESAURL = TipoSite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
           Else
               XMESAURL=   TipoSite("") & request.servervariables("SERVER_NAME") & XMESAURL
           End If

          ' XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL


        END IF




       ' XGUMENSAGEw = Replace(XGUMENSAGEw, "<APLURL>", "<U><A HREF='" & XMESAURL & "' TARGET='_blank' >")
        'XGUMENSAGEw = Replace(XGUMENSAGEw, "</APLURL>", "</A></U>")
        'XGUMENSAGEw=XGUMENSAGEw & "<br>" & "<U><A HREF='" & XMESAURL & "' TARGET='_blank' >VER...</A></U>"
        if session("APLMDCLI")="BPN" THEN
         XGUMENSAGEw=XGUMENSAGEw & "<br>" & "<U><A HREF='" & XMESAURL & "' TARGET='_blank' >Ver : " & EsqnomP(sTipo,nCodigo,ofv) & "</A></U>"
        ELSE
        XGUMENSAGEw=XGUMENSAGEw & "<br>" & "<U><A HREF='" & XMESAURL & "' TARGET='_blank' >VER...</A></U>"
        END IF
        WEMESSAGE = XGUMENSAGEw

        Select Case WETIPO

            Case "USR"

                ARMAILS=SPLIT(WECOD,CHR(9))
                DELIM=""
                sInputx=""

                FOR iMAIL=0 TO UBOUND(ARMAILS)

                IF mid(ARMAILS(iMAIL)& "",1,4)="GRM-" THEN
                   sgrupoM=mid( ARMAILS(iMAIL)& ""  ,5,len(ARMAILS(iMAIL)& "")  )
                 '  RESPONSE.WRITE sgrupoM
                   xxx=ucase(getMailDelGrupo(sgrupoM))

                   sInputx = sInputx & DELIM & xxx
                ELSE
                xxxx = getInfoUser(ARMAILS(iMAIL)& "")
                If session("SGU_mail") <> "" Then
                  sInputx = sInputx & DELIM & session("SGU_mail")

                END IF
                END IF

                DELIM=";"
                NEXT

                stowho = sInputx & ""
                If Len(stowho) > 0 Then
                    stowho = Trim(stowho)

                    subject = "RECOMENDACION"
                    If WETIPO = "USR" Then
                        XPARA2 = XPARA
                    Else
                        XPARA2 = XPARA & " / " & session("SGU_nombre")
                    End If

                    Xmessage = GenMDLMAILex(XPARA2, WEORIGEN, WEMESSAGE)
                    Call SendMail(Xmessage, stowho, WEORIGENMAIL, subject, session("MailComponent"), "SI", "1", "NO")
                End If




        End Select






    End If

End Function
Function GETRESPOCKT(RESPCOD, RESPTIP)

    GETRESPOCKT = "SIN ASIGNAR"
    sInputxEco = "SIN ASIGNAR"



    XRESPO = RESPCOD & ""

    If XRESPO <> "" And XRESPO <> "NO" Then

        Set cn2 = ofv.conectar(ofv.strconn0)
        Set cn3 = ofv.conectar(ofv.strconn2)

        Select Case RESPTIP

            Case "USR"

                xxxx = getInfoUserw(XRESPO)
                sInputx = session("SGU_nombre")

                ' strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
                ' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                ' if not rsniv.eof then sInputx = rsniv(1)
                ' call ofv.cerrarconsulta(rsniv)
                GETRESPOCKT = sInputx

            Case "GRP"
                Strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
                Set rsniv = ofv.crearconsultaEx(Strsql, cn2, 1, parametros)
                If Not rsniv.EOF Then sInputxEco = rsniv(1)
                Call ofv.cerrarconsulta(rsniv)
                GETRESPOCKT = sInputxEco

            Case "SEC"
                XRESPO = 0 & XRESPO
                Strsql = "Select * From ECO where ECO = " & XRESPO
                '''     response.write strsql
                Set rsniv = ofv.crearconsultaEx(Strsql, cn3, 1, parametros)
                If Not rsniv.EOF Then sInputxEco = rsniv(1)
                Call ofv.cerrarconsulta(rsniv)
                GETRESPOCKT = sInputxEco
        End Select

        '  GETRESPOCKT = sInputx

    End If


End Function


FUNCTION GenMDLMAILex (XGUPARA,XGUORIGEN,XGUMENSAGE)

''message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"


     XGUMENSAGEw = XGUMENSAGE & " "

     XGUMENSAGEw = replace(XGUMENSAGEw,"@@@NORMREQ@@@","<u>" & SESSION("NOMREQ") & "</u>")
     XGUMENSAGEw = replace(XGUMENSAGEw,"@@@FECREQ@@@",now())



     Whtml = ""

     SELECT CASE  SESSION("MAILOPC")

            CASE "HTML"

                  Whtml = Whtml & "<DIV STYLE='background-color:" & Session("CMFMENU") & ";height:40;width:600px;"
                  Whtml = Whtml & "BORDER-color:#E0E0E0;BORDER-STYLE:SOLID;BORDER-WIDTH:1;'  >"
                 ' Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<P VALIGN=MIDDLE >&nbsp;&nbsp;<font size=4 face='" & Session("CMFMEFTFC") & "' color='" & Session("CMFMEFT") & "' ><B>"
                  Whtml = Whtml & Session("CorTitApli") & "</B></FONT></P>"
                  Whtml = Whtml & "</DIV>"
                  Whtml = Whtml & "<BR>"

                  Whtml = Whtml & "<DIV STYLE='BORDER-color:" & Session("CMFMENU") & ";BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                  Whtml = Whtml & "height:50;width:600px;'  >"
                  Whtml = Whtml & "<CENTER>"

                 ' Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                 ' Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
                 ' Whtml = Whtml & "<P ALIGN=LEFT >"
                 ' Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
                 ' Whtml = Whtml & "&nbsp;&nbsp;DESTINATARIO</B></FONT></P>"
                 ' Whtml = Whtml & "</DIV>"

              '    Whtml = Whtml & "<br>"
              '    Whtml = Whtml & "<DIV STYLE='width:80%;' >"
              '    Whtml = Whtml & "<P WIDTH=450 ALIGN='JUSTIFY' >"
              '    Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B><U>"
              '    Whtml = Whtml & XGUPARA & "</U></B></FONT></P>"
              '    Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV STYLE='BORDER-color:#888888;BORDER-STYLE:SOLID;BORDER-WIDTH:1;"
                  Whtml = Whtml & "background-color:#F0F0F0;height:25;width:100%;'  >"
                  Whtml = Whtml & "<P ALIGN=LEFT >"
                  Whtml = Whtml & "<font size=1 color=#888888 face='" & Session("CMFMEFTFC") & "' ><B>"
                  Whtml = Whtml & "&nbsp;&nbsp;DESCRIPCION</B></FONT></P>"
                  Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV STYLE='width:80%;' >"
                  Whtml = Whtml & "<P ALIGN='JUSTIFY' >"
                  Whtml = Whtml & "<font size=2 color=#888888 face='" & Session("CMFMEFTFC") & "' >"
                  Whtml = Whtml & XGUMENSAGEw
                  Whtml = Whtml & "</FONT></P>"
                  Whtml = Whtml & "</DIV>"
                  Whtml = Whtml & "<BR>"
                  Whtml = Whtml & "</DIV>"

                  Whtml = Whtml & "<br>"
                  Whtml = Whtml & "<DIV ALIGN='RIGHT' STYLE='width:200px;' >"
                  Whtml = Whtml & "<P  >"
                  Whtml = Whtml & "<font size=1 color=#6699cc face='" & Session("CMFMEFTFC") & "' ><B><i>"
                  Whtml = Whtml & "COR&nbsp;E-SOLUTION&nbsp;SUITE&nbsp;V4.8.5</i></B></FONT></P>"
                  Whtml = Whtml & "</DIV>"



            CASE ELSE


                  Whtml = Whtml & SESSION("NOMAPLI")  & "<BR>" & CHR(13) & CHR(13)

                '  Whtml = Whtml & XGUPARA & ": <BR>" & CHR(13)

                '  Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & CHR(13)

                  Whtml = Whtml & XGUMENSAGEw  & "<BR>" & CHR(13)
                  Whtml = Whtml & "------------------------------------------------------------------------"  & "<BR>" & CHR(13)




     END SELECT


     GenMDLMAILex = Whtml



 END FUNCTION

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

      response.redirect "CoObjRecomendar.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & NCodigo



%>