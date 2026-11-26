<%

  WKSIGUE = "N"


  if wkaplic = "PL" then


     wkopcionV = request.querystring("opcionV") 

     sPath  = Request.ServerVariables("SCRIPT_NAME")
     spathi = len(sPath) - InStrRev(sPath, "/", -1)
     sPath2 = right(sPath, spathi)
 

     if wkopcionV = "SI" then

        wkrellink = request.querystring("rellink")
        wkcodV    = request.form("codV") 


        select case wkcodV

               case "XC"

                 ''   sHTML = sHTML & "<CENTER><p>entro por XC rel=" & wkrellink & "</p></center>"

                     SESSION("CODVINOBJ") = "XC"

               case "NC"

                 ''   sHTML = sHTML & "<CENTER><p>entro por NC rel=" & wkrellink & "</p></center>"

                     SESSION("CODVINOBJ") = "NC"

               case "SC"

                 ''   sHTML = sHTML & "<CENTER><p>entro por SC rel=" & wkrellink & "</p></center>"

                     SESSION("CODVINOBJ") = "SC"

                     RESPONSE.REDIRECT "EDVINCULOOBJ.ASP?ASP=" & sPath2 & "&" & REQUEST.QUERYSTRING



        end select

        IF wkcodV = "SC" or wkcodV = "NC"  or wkcodV = "XC" THEN

           strsql = "update Objetos set codvinOBJ = " & wkcodV
           strsql = strsql & " where  codobjeto = " & nCodigo
           call ofv.crearconsultaEx(StrSql,cn,1,volver)

        end if 



     end if




     IF SESSION("CODVINOBJ") <> "SC" AND SESSION("CODVINOBJ") <> "NC"  AND SESSION("CODVINOBJ") <> "XC" THEN


        wkrellink    = "NO"

        Strsql = "Select * From Objetos Where codobjeto = " & nCodigo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs.eof then 
           wkrellink    = RS(15) & ""
           IF WKRELLINK = "" OR WKRELLINK = "0" THEN wkrellink    = "NO"
           Strsql = "Select * From Objlnk Where codobjeto = " & nCodigo & " and (ltipo = 'LPL' or ltipo = 'IPL')"
           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs.eof then 
              wkreltipo    = trim(RS(10) & "")
           end if

        end if

        call ofv.cerrarconsulta(rs)


        sPath2 = sPath2 & "?" & request.querystring & "&opcionV=SI&rellink=" & wkrellink


        wkb1  = "N"
        wkb2  = "N"
        wkb3  = "N"

        wkcol = "1"

              SINPUTT1 = "<FONT COLOR='#993333' ><B><U>CONFIGURACION DE LA VINCULACION:</U></B></FONT><BR><BR>"


        if SESSION("CODVINOBJ") = "" then

           IF wkrellink = "NO" THEN

              SINPUTT1 = SINPUTT1 & "No se ha vinculado ningun documento a este procedimiento, "
              SINPUTT1 = SINPUTT1 & "la aplicación lo considerará como un procedimiento <b>SIN VINCULACION</b>.<br>"
              SINPUTT1 = SINPUTT1 & "Confirma esta clasificación? caso contrario deberá indicar el documento vinculado."

              wkb1  = "S"
              wkb2  = "N"
              wkb3  = "N"

              wkcol = "1"


           ELSEIF  wkreltipo = "IPL" THEN

              SINPUTT1 = SINPUTT1 & "Este procedimiento tiene un documento vinculado externo a la aplicación, "
              SINPUTT1 = SINPUTT1 & "pero esta clasificado como <b>SIN VINCULACION</b>.<br>"
              SINPUTT1 = SINPUTT1 & "Confirma esta clasificación, O desea modificar a <b>VINCULACION REFERENCIAL</b>?."

              wkb1  = "S"
              wkb2  = "S"
              wkb3  = "N"

              wkcol = "2"

           ELSE

              SINPUTT1 = SINPUTT1 & "Este procedimiento tiene un documento vinculado perteneciente a la aplicación, "
              SINPUTT1 = SINPUTT1 & "pero esta clasificado como <b><u>SIN VINCULACION</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "Confirma esta clasificación, desea modificar a <b><u>VINCULACION REFERENCIAL</u></b>,"
              SINPUTT1 = SINPUTT1 & " o desea modificar a <b><u>VINCULACION ESTRICTA</u></b>?."

              wkb1  = "S"
              wkb2  = "S"
              wkb3  = "S"

              wkcol = "3"

           END IF





        elseif SESSION("CODVINOBJ") = "N" then 

           IF wkrellink = "NO" THEN

              SINPUTT1 = SINPUTT1 & "No se ha vinculado ningun documento a este procedimiento, "
              SINPUTT1 = SINPUTT1 & "pero esta clasificado como un procedimiento con <b><u>VINCULACION REFERENCIAL</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "desea modificar a <b><u>SIN VINCULACION</u></b>? caso contrario deberá indicar el documento vinculado."

              wkb1  = "S"
              wkb2  = "N"
              wkb3  = "N"

              wkcol = "1"

           ELSEIF  wkreltipo = "IPL" THEN

              SINPUTT1 = SINPUTT1 & "Este procedimiento tiene un documento vinculado externo a la aplicación, "
              SINPUTT1 = SINPUTT1 & "y esta clasificado como <b><u>VINCULACION REFERENCIAL</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "Confirma esta clasificación?."

              wkb1  = "N"
              wkb2  = "S"
              wkb3  = "N"

              wkcol = "1"

           ELSE

              SINPUTT1 = SINPUTT1 & "Este procedimiento tiene un documento vinculado perteneciente a la aplicación, "
              SINPUTT1 = SINPUTT1 & "pero esta clasificado como <b><u>VINCULACION REFERENCIAL</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "Confirma esta clasificación, "
              SINPUTT1 = SINPUTT1 & " o desea modificar a <b><u>VINCULACION ESTRICTA</u></b>?."

              wkb1  = "N"
              wkb2  = "S"
              wkb3  = "S"

              wkcol = "2"

           END IF



        else


           IF wkrellink = "NO" THEN

              SINPUTT1 = SINPUTT1 & "No se ha vinculado ningun documento a este procedimiento, "
              SINPUTT1 = SINPUTT1 & "pero esta clasificado como un procedimiento con <b><u>VINCULACION ESTRICTA</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "desea modificar a <b><u>SIN VINCULACION</u></b>? caso contrario deberá indicar el documento vinculado."

              wkb1  = "S"
              wkb2  = "N"
              wkb3  = "N"

              wkcol = "1"

           ELSEIF  wkreltipo = "IPL" THEN

              SINPUTT1 = SINPUTT1 & "Este procedimiento tiene un documento vinculado externo a la aplicación, "
              SINPUTT1 = SINPUTT1 & "pero esta clasificado como un procedimiento con <b><u>VINCULACION ESTRICTA</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "Desea modificar a <b><u>VINCULACION REFERENCIAL</u></b>?."

              wkb1  = "N"
              wkb2  = "S"
              wkb3  = "N"

              wkcol = "1"

           ELSE

              SINPUTT1 = SINPUTT1 & "Este procedimiento tiene un documento vinculado perteneciente a la aplicación, "
              SINPUTT1 = SINPUTT1 & "y esta clasificado como un procedimiento con <b><u>VINCULACION ESTRICTA</u></b>.<br>"
              SINPUTT1 = SINPUTT1 & "Desea modificar a <b><u>VINCULACION REFERENCIAL</u></b>, o confirma esta clasificación?. "

              wkb1  = "N"
              wkb2  = "S"
              wkb3  = "S"

              wkcol = "2"

           END IF


        end if 


              BOTONT1  = "<INPUT TYPE=BUTTON CLASS=BT VALUE='SIN VINCULACION' onclick='document.FVINC1.submit();' >"
              BOTONT2  = "<INPUT TYPE=BUTTON CLASS=BT VALUE='REFERENCIAL' onclick='document.FVINC2.submit();' >"
              BOTONT3  = "<INPUT TYPE=BUTTON CLASS=BT VALUE='ESTRICTA' onclick='document.FVINC3.submit();' >"

              Celda = ofv.GenCelda("","ut COLSPAN=" & wkcol & "  valign=top STYLE='text-transform:uppercase;text-align:justify;background-color:#ffffff;' ","","","150","","",SINPUTT1,"si")
              Fila = ofv.GenRow("","","","","","","",Celda,"si")

              celda = ""

              if wkb1 = "S" then Celda = Celda & ofv.GenCelda("","vt","center","","","","",BOTONT1,"si")
              if wkb2 = "S" then Celda = Celda & ofv.GenCelda("","vt","center","","","","",BOTONT2,"si")
              if wkb3 = "S" then Celda = Celda & ofv.GenCelda("","vt","center","","","","",BOTONT3,"si")
              Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")




        Formulario = ""

        WKSIGUE = "N"
        sHTML = sHTML & "<CENTER><br><br><br>" & ofv.GenTabla("","AAA border=1 BGCOLOR='#DDDDDD' ","","60%","","0","","5","4","","",Fila,"si") & "</CENTER>"

        sAccion    = spath2

        sinputh    = ofv.GenerarInput("CODV","XC","hidden","","1","dt","")
        Formulario = Formulario & ofv.GenForm("FVINC1","aaa style='display:none;' ",sAccion,"","post","",sinputh,"si")

        sinputh    = ofv.GenerarInput("CODV","NC","hidden","","1","dt","")
        Formulario = Formulario & ofv.GenForm("FVINC2","aaa style='display:none;' ",sAccion,"","post","",sinputh,"si")

        sinputh    = ofv.GenerarInput("CODV","SC","hidden","","1","dt","")
        Formulario = Formulario & ofv.GenForm("FVINC3","aaa style='display:none;' ",sAccion,"","post","",sinputh,"si")

        sHTML = sHTML & Formulario

     ELSE
        WKSIGUE = "S"
     END IF



  else

     WKSIGUE = "S"

  end if

%>