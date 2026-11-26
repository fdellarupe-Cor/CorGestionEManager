<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx   = cdbl(request.querystring("CantRegMover"))
  vuelta          = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL          = request.querystring("MODAL")
  sVISTA          = request.querystring("CVISTA")

  sMODALMAIN      = request.querystring("MODALMAIN")







  if sMODAL = "" then sMODAL = "SET"

  if sMODAL = "SET" AND SESSION("CVISTA") = "VISTATEMP" then sMODAL = "GETTEMP"

  if sMODAL = "SETSEL" THEN 
     SESSION("CVISTA") = sVISTA 
     sMODAL = "SET"
  END IF

  WKFVISTANOM = SESSION("NOMVISTA")

  WKVISION = ""


  DIM TABVISTA(20,10)
  TTABVISTA = 0



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn2)


  WKFPERVT = "N"
  WKFPERVN = "N"


     Strsql = "Select PERMITEMR,PERMITEMO,PERMITEVT,PERMITEVN From GPREQRESPONSABLES WHERE TIPOUSR = 'U' AND USUARIOSECTOR = '" & SESSION("USUARIO") & "' "
     set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rs2.eof then


        IF (RS2(2) & "") = "S"  THEN

            WKFPERVT = "S"

        END IF 

        IF (RS2(3) & "") = "S" THEN

            WKFPERVN = "S"

        END IF 


     END IF





    ''   WKMSTSEPITM = CHR(9)
    ''   WKMSTSEPDAT = CHR(8)

        WKMSTSEPITM = "|||"
        WKMSTSEPDAT = "###"

   wkproccod1 = "&nbsp;"
   wkproccod2 = "&nbsp;"
   wkproccod3 = "&nbsp;"
   wkproccod4 = "&nbsp;"

   wkproccod1S = ""
   wkproccod2S = ""
   wkproccod3S = ""
   wkproccod4S = ""

  sHTML = ""

  if sMODAL = "SET"  then




     IF sMODALVISTA  = "" THEN sMODALVISTA  = SESSION("CVISTA")

     Strsql = "Select DETVISTA1, DETVISTA2, CODVISTA, USUARIO, NOMVISTA From GPVISTAS WHERE CODVISTA = '" & sMODALVISTA & "' "


  ' response.write strsql & "<br>"

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     IF RSX1.EOF THEN

        sMODAL = "NEW" 

        SESSION("MODACT") = "N" 


     ELSE

        SESSION("MODACT") = "N"
      
        IF RSX1(3) = SESSION("USUARIO") THEN SESSION("MODACT") = "S" 


        WKFVISTA = RSX1(2)

     '''   SESSION("CVISTA") = RSX1(2)

        WKTXTVISTA = RSX1(0) & RSX1(1) & ""

        WKFVISTANOM = RSX1(4) & ""


     END IF



     call ofv.cerrarconsulta(RSX1)

  ELSEif sMODAL = "GETACT"  then

     WKFMODOACTUAL = request.form("MODOACTUAL") 

     IF WKFMODOACTUAL = "TMP" THEN
        SESSION("MODACT") = "N"
        sMODAL = "VISTATEMP" 
     END IF 

     IF WKFMODOACTUAL = "VACT" THEN

        sMODAL = "SIGACTV" 
     END IF 

     IF WKFMODOACTUAL = "VNVA" THEN

        sMODAL = "SIGACTN" 
     END IF 

  ELSEif sMODAL = "PROCACT"  then

     WKFMODOACTUAL = request.form("MODOACTUAL") 

     IF WKFMODOACTUAL = "KNC" THEN

        if SESSION("CVISTA") = "VISTATEMP" then
           sMODAL = "GETTEMP"
        ELSE
           sMODAL = "SET" 
        END IF
     END IF 

     IF WKFMODOACTUAL = "AYC" THEN

        sMODAL = "SAVEACT" 
     END IF 

     IF WKFMODOACTUAL = "AYL" THEN

        sMODAL = "SAVEACTL" 
     END IF 


  END IF


  if sMODAL = "SET"  then

     IF sMODALVISTA  = "" THEN sMODALVISTA  = SESSION("CVISTA")

     Strsql = "Select DETVISTA1, DETVISTA2, CODVISTA, USUARIO, NOMVISTA From GPVISTAS WHERE CODVISTA = '" & sMODALVISTA & "' "


  ' response.write strsql & "<br>"

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     IF RSX1.EOF THEN


     ELSE


        WKFVISTA = RSX1(2)

        WKTXTVISTA = RSX1(0) & RSX1(1) & ""

        WKFVISTANOM = RSX1(4) & ""


     END IF



     call ofv.cerrarconsulta(RSX1)


  ELSEif sMODAL = "NEW"  then

''''' PREFORMATEO DE NUEVA VISTA DESDE DEFAULT ''''''''''''''''''''''''''''''''''''''''''''''

        WKFVISTANOM = "VISTA NUEVA"

      SESSION("MODACT") = "N"

     WKTXTVISTA = "VISTA" & WKMSTSEPDAT & "NUEVA VISTA " & SESSION("SEGUSUARIO")
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "ORIGINANTE" & WKMSTSEPDAT & "-MIUSR-"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "ASIGNADO" & WKMSTSEPDAT & "TODOS"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "TIPOREQ" & WKMSTSEPDAT & "TODOS"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "ESTADO" & WKMSTSEPDAT & "TODOS"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "FECHAINI" & WKMSTSEPDAT & "-U7-"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "FECHAFIN" & WKMSTSEPDAT & "TODOS"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "CODSIST" & WKMSTSEPDAT & "TODOS"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "CODMOD" & WKMSTSEPDAT & "TODOS"
     WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & "GPCASOUSO" & WKMSTSEPDAT & "TODOS"



  ELSEif sMODAL = "GETACT" OR sMODAL = "VISTATEMP" then

''''' PARSEO DE LA VISTA DESDE EL FORM DE MODIFICACIONES ''''''''''''''''''''''''''''''''''''''''''''''

     IF sMODAL = "VISTATEMP" THEN WKFVISTANOM = "VISTA TEMPORAL"


     WKTXTVISTA = "VISTA###NUEVA VISTA " & SESSION("SEGUSUARIO")

     WKFACCMASMENOS = "N"
     WKFACCMASMENOSDAT = ""
 
     IF (WKFMODOACTUAL & "") <> "" THEN

         WKFACCTAB = SPLIT(WKFMODOACTUAL,"D")

         IF  WKFACCTAB(0) = "MAS" OR WKFACCTAB(0) = "MEN" THEN

             WKFACCMASMENOS    = WKFACCTAB(0)
             WKFACCMASMENOSDAT = WKFACCTAB(1)

         END IF


     END IF 

     WKFBASEFORM = request.form


     WKFTABFORM = SPLIT(WKFBASEFORM,"&")

     FOR MM=0 TO UBOUND(WKFTABFORM)

         WKFTABITMFORM = SPLIT(WKFTABFORM(MM),"=")
      ''   response.write WKFTABITMFORM(0) & ":" & WKFTABITMFORM(1) & "<br>" & chr(13)

        IF WKFTABITMFORM(0) = "MODOACTUAL" THEN


        ELSE

         WKFTABITMDATFORM = SPLIT(WKFTABITMFORM(0),"-")

      ''   response.write WKFTABITMDATFORM(1) & "*" & WKFTABITMDATFORM(2) & ":" & WKFTABITMFORM(1) & "<br>" & chr(13)


         TABVISTA(CLNG(WKFTABITMDATFORM(1)),CLNG(WKFTABITMDATFORM(2))) = WKFTABITMFORM(1)
         TABVISTA(CLNG(WKFTABITMDATFORM(1)),10) = CLNG(WKFTABITMDATFORM(2))       
         TTABVISTA = CLNG(WKFTABITMDATFORM(1))

       END IF       

     NEXT



     FOR II=1 TO TTABVISTA 

         WKTXTVISTA = WKTXTVISTA & WKMSTSEPITM & TABVISTA(II,0)

         IF  WKFACCMASMENOS  <> "N" THEN

             IF CLNG(WKFACCMASMENOSDAT) = II THEN

                IF  WKFACCMASMENOS  = "MAS" THEN

                    TABVISTA(II,10) = TABVISTA(II,10) + 1

                    TABVISTA(II,TABVISTA(II,10)) = 0



                ELSE

                    TABVISTA(II,10) = TABVISTA(II,10) - 1

                END IF

             END IF

         END IF


         FOR YY=1 TO TABVISTA(II,10) 

             WKTXTVISTA = WKTXTVISTA & WKMSTSEPDAT & TABVISTA(II,YY)

         NEXT 


     NEXT 


  ELSEif sMODAL = "GETTEMP"  then

''''' RETOMA VISTA TEMPORAL ''''''''''''''''''''''''''''''''''''''''''''''

       SESSION("MODACT") = "N"

       WKTXTVISTA = SESSION("VISTATEMP") & ""



  ELSEif sMODAL = "SIGACTN" OR  sMODAL = "SIGACTV" then

''''' PREPROCESO DE ACTUALIZACION/CREACION DE VISTA ''''''''''''''''''''''''''''''''''''''''''''''

       SESSION("MODACT") = "S"

       WKTXTVISTA = SESSION("VISTATEMP") & ""

       WKHOY      = CLNG(DATE()) - CLNG(CDATE("01-01-1900")) 

       if sMODAL = "SIGACTV" then

          Strsql = "Select * From GPVISTAS WHERE CODVISTA = '" & SESSION("CVISTA") & "' "
  ' response.write strsql & "<br>"

          set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

          IF NOT RSX1.EOF THEN

             WKFVISTANOM    = RSX1(9)

             WKFFRMCODVISTA = RSX1(1)
             WKFFRMFCARGA   = RSX1(2)
             WKFFRMFMODIF   = WKHOY
             WKFFRMESTADO   = RSX1(4)
             WKFFRMUSUARIO  = RSX1(5)
             WKFFRMTIPO     = RSX1(6)
             WKFFRMDETVISTA = WKTXTVISTA
             WKFFRMNOMVISTA = RSX1(9)
             WKFFRMESDEFAULT = RSX1(10)
             WKFFRMESPUBLIC = RSX1(11)


          END IF

       ELSE

             WKFVISTANOM    = "NUEVA VISTA " & SESSION("USUARIO")

             WKFFRMCODVISTA = GTrnd("RESP")
             WKFFRMFCARGA   = WKHOY
             WKFFRMFMODIF   = WKHOY
             WKFFRMESTADO   = "S"
             WKFFRMUSUARIO  = SESSION("USUARIO")
             WKFFRMTIPO     = "REQ"
             WKFFRMDETVISTA = WKTXTVISTA
             WKFFRMNOMVISTA = "NUEVA VISTA " & SESSION("USUARIO")
             WKFFRMESDEFAULT = "N"
             WKFFRMESPUBLIC = "N"


       END IF  

       SESSION("NOMVISTA") = WKFVISTANOM


  ELSEif sMODAL = "SAVEACT" or sMODAL = "SAVEACTL" then
''''' ACTUALIZACION/CREACION DE VISTA ''''''''''''''''''''''''''''''''''''''''''''''



      

             WKFFRMCODVISTA  = request.form("codvista") 
             WKFFRMFCARGA    = request.form("FCARGA") 
             WKFFRMFMODIF    = request.form("FMODIF") 
             WKFFRMESTADO    = request.form("ESTADO") 
             WKFFRMUSUARIO   = request.form("USUARIO") 
             WKFFRMTIPO      = request.form("TIPO") 
             WKFFRMDETVISTA  = request.form("DETVISTA") 
             WKFFRMNOMVISTA  = request.form("NOMVISTA") 
             WKFFRMESDEFAULT = request.form("ESDEFAULT") 
             WKFFRMESPUBLIC  = request.form("ESPUBLIC") 

             WKFFRMNOMVISTA  = UCASE(WKFFRMNOMVISTA & "")

             IF LEN(WKFFRMDETVISTA & "") > 2900 THEN
                WKFFRMDETVISTA1 = LEFT(WKFFRMDETVISTA,2900)
                WKFFRMDETVISTA2 = RIGHT(WKFFRMDETVISTA,LEN(WKFFRMDETVISTA & "") - 2900)

             ELSE
                WKFFRMDETVISTA1 = WKFFRMDETVISTA
                WKFFRMDETVISTA2 = ""
             END IF 

            SESSION("VISTATEMP") = WKFFRMDETVISTA

            SESSION("NOMVISTA") = WKFFRMNOMVISTA


       if   WKFFRMCODVISTA  = SESSION("CVISTA") then


            strsql = "UPDATE GPVISTAS SET "
            strsql = strsql & " FCARGA = '" & WKFFRMFCARGA & "',"
            strsql = strsql & " FMODIF = '" & WKFFRMFMODIF & "',"
            strsql = strsql & " ESTADO = '" & WKFFRMESTADO & "',"
            strsql = strsql & " USUARIO = '" & WKFFRMUSUARIO & "',"
            strsql = strsql & " NOMVISTA = '" & WKFFRMNOMVISTA & "',"
            strsql = strsql & " ESDEFAULT = '" & WKFFRMESDEFAULT & "',"
            strsql = strsql & " ESPUBLIC = '" & WKFFRMESPUBLIC & "',"
            strsql = strsql & " DETVISTA1 = '" & WKFFRMDETVISTA1 & "',"
            strsql = strsql & " DETVISTA2 = '" & WKFFRMDETVISTA2 & "',"
            strsql = strsql & " TIPO = '" & WKFFRMTIPO & "' "

            strsql = strsql & " WHERE CODVISTA = '" & WKFFRMCODVISTA & "' "
    
            call ofv.crearconsultaEx(StrSql,cn,1,parametros)



       else
            
            strsql = "insert into GPVISTAS ("
            strsql = strsql & "CODVISTA,FCARGA,FMODIF,ESTADO,"
            strsql = strsql & "USUARIO,NOMVISTA,ESDEFAULT,ESPUBLIC,"
            strsql = strsql & "DETVISTA1,DETVISTA2,TIPO"
            strsql = strsql & ") values ("
            strsql = strsql & "'"& WKFFRMCODVISTA & "','"& WKFFRMFCARGA & "','"& WKFFRMFMODIF & "','"& WKFFRMESTADO & "',"
            strsql = strsql & "'"& WKFFRMUSUARIO & "','"& WKFFRMNOMVISTA & "','"& WKFFRMESDEFAULT & "','"& WKFFRMESPUBLIC & "',"
            strsql = strsql & "'"& WKFFRMDETVISTA1 & "','"& WKFFRMDETVISTA2 & "','"& WKFFRMTIPO & "'"
            strsql = strsql & ")"
    
            call ofv.crearconsultaEx(StrSql,cn,1,parametros)

            SESSION("CVISTA") = WKFFRMCODVISTA


       end if 


       if sMODAL = "SAVEACTL" then sMODAL = "LISVISTA"



  END IF



    sForm = "formvistas"

    sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & " document." & sform & ".submit(); return false;" & CHR(34)

     wkproccod4 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Lista de vistas disponibles.' HREF='" & sAccion & "' >VISTAS</A>"

     wkproccod4S = " style='background-color:#00AA00;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "




   WKACCION = " - ACTUALIZACION DE VISTAS"

   Celda = ""
   Celda = Celda & ofv.GenCelda("","uti style='color:#ffffff;'  ","","40%","","","","LISTA DE REQUERIMIENTOS  " & WKACCION,"si")
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod1S,"CENTER","15%","","","",wkproccod1,"si")  
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod2S,"CENTER","15%","","","",wkproccod2,"si") 
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod3S,"CENTER","15%","","","",wkproccod3,"si") 
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod4S,"CENTER","15%","","","",wkproccod4,"si") 
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

   sAccion = "GPLISTAREQvista.asp?MODAL=LISVISTA"
   Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",fila,"si")


  sHTML = sHTML & ofv.GenTabla("","aaa bgcolor='#888888' style='border-style:solid;border-width:1;BORDER-COLOR:#000000;' valign='middle' align='center' ","","100% ","30","0","","2","1","","",Formulario,"si")


    sForm = "formlis"


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''' PANEL DE DEFINICION DE DATOS PARA LA VISTA ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



  if sMODAL <> "SAVEACT" AND sMODAL <> "SAVEACTL" AND sMODAL <> "SIGACTN" AND sMODAL <> "SIGACTV" AND sMODAL <> "LISVISTA" then


 
  

     SESSION("VISTATEMP") = WKTXTVISTA

     SESSION("NOMVISTA") = WKFVISTANOM

     WKTABLAITEM = SPLIT(WKTXTVISTA,WKMSTSEPITM) 

    filas = ""


   celdas = ""

    sInput = "&nbsp;"

    celdas = celdas & ofv.GenCelda("","VTI colspan=3 style='color:#FFFFFF;' ","","","","","","SELECCION DE PARAMETROS","si")

    CeldaS = CeldaS & ofv.GenCelda("","utI  style='BACKGROUND-COLOR:#DDDDDD;color:#888888;' ","RIGHT","","","","","VISTA:&nbsp;","si") 
    CeldaS = CeldaS & ofv.GenCelda("","utI COLSPAN=2 style='color:#FFFFFF;' ","","","","","",WKFVISTANOM,"si") 


    filas = filas & ofv.GenRow("","AAA BGCOLOR='#888888' ","","","25","","",celdas,"si")



     FOR II=1 TO UBOUND(WKTABLAITEM) 



         WKTABLADATOS = SPLIT(WKTABLAITEM(II),WKMSTSEPDAT) 

         FOR YY=0 TO UBOUND(WKTABLADATOS) 

             IF YY > 9 THEN
                EXIT FOR
             END IF 

             TABVISTA(II,YY) = WKTABLADATOS(YY)
             TABVISTA(II,10) = YY


         NEXT

         TTABVISTA = II

     NEXT


    WKTXTWHERE = ""
    WKSEP      = ""

    WKHOY      = CLNG(DATE()) - CLNG(CDATE("01-01-1900")) 


     FOR II=1 TO TTABVISTA 


         filasx1 = ""
         wkfctacampo = 0

         WKCODCAMPO = "D-" & II

         FOR YY=1 TO TABVISTA(II,10) 

             WKII = REPLACE("-" & CSTR(100 + YY),"-1","-")


             sInput = "&nbsp;&nbsp;"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Ut  ","RIGHT","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             IF YY > 1 THEN sInput1 = "&nbsp;O&nbsp;"

             if  TABVISTA(II,10) > 1 then

                 scomK = ""
                 scomD = ""

             else 

                 scomK = "TODOS" & chr(9)
                 scomD = "TODOS" & chr(9)

             end if


         SELECT CASE CSTR(II)

                CASE "1" 
                      WKFNOMDATO = "ORIGINANTE"
                      if  bModificar  then

                          scomK = scomK & "-MIUSR-" & chr(9) & "-USRAC-" & chr(9)
                          scomD = scomD & "MI USUARIO" & chr(9) & "USUARIOS A CARGO" & chr(9)
                          strsql = "Select * From USUARIOS    "
                          strsql = strsql & " WHERE ID IN ("
                             wkfsep3 = "" 
                             strsql2 = "Select * From GPREQRESPONSABLES WHERE TIPOUSR = 'U' "
                             set rsniv = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
                             DO UNTIL RSNIV.EOF
                                strsql = strsql & wkfsep3 & "'" & RSNIV(2) & "'"
                                wkfsep3 = "," 
                             RSNIV.MOVENEXT
                             LOOP

                          strsql = strsql & " )"
                          strsql = strsql & " Order By DESCRIPCION"
 ''  response.write strsql & "<br>" & chr(13)
                          set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                          DO UNTIL rsniv.eof
                             scomK = scomK & rsniv(0) & chr(9)
                             scomD = scomD & rsniv(1) & chr(9)
                          RSNIV.MOVENEXT
                          LOOP
                          call ofv.cerrarconsulta(rsniv)

                      	  sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")
     
                      else

                          SELECT CASE TABVISTA(II,YY)

                                 CASE "-MIUSR-"
                                      sInput = "MI USUARIO"

                                 CASE "-USRAC-"
                                      sInput = "USUARIOS A CARGO"

                                 CASE "TODOS"
                                      sInput = "TODOS"

                                 CASE ELSE

                                   strsql = "Select * From USUARIOS where ID = '" & TABVISTA(II,YY) & "' "
                                   set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                                   if not rsniv.eof then sInput = rsniv(1)
                                   call ofv.cerrarconsulta(rsniv)
                          END SELECT
                          sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sInput,"readonly","","70","dt","")

                      end if


                CASE "2" 
                      WKFNOMDATO = "RESPONSABLE"
                      if  bModificar  then

                          scomK = scomK & "-MIUSR-" & chr(9) & "-USRAC-" & chr(9)
                          scomD = scomD & "MI USUARIO" & chr(9) & "USUARIOS A CARGO" & chr(9)
                          strsql = "Select * From USUARIOS    "
                          strsql = strsql & " WHERE ID IN ("

                             wkfsep3 = "" 
                             strsql2 = "Select * From GPREQRESPONSABLES WHERE TIPOUSR = 'U' "
                             set rsniv = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
                             DO UNTIL RSNIV.EOF
                                strsql = strsql & wkfsep3 & "'" & RSNIV(2) & "'"
                                wkfsep3 = "," 
                             RSNIV.MOVENEXT
                             LOOP

                          strsql = strsql & " )"
                          strsql = strsql & " Order By DESCRIPCION"
                          set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                          DO UNTIL rsniv.eof
                             scomK = scomK & rsniv(0) & chr(9)
                             scomD = scomD & rsniv(1) & chr(9)
                          RSNIV.MOVENEXT
                          LOOP
                          call ofv.cerrarconsulta(rsniv)

                      	  sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")
     
                      else

                          SELECT CASE TABVISTA(II,YY)

                                 CASE "-MIUSR-"
                                      sInput = "MI USUARIO"

                                 CASE "-USRAC-"
                                      sInput = "USUARIOS A CARGO"

                                 CASE "TODOS"
                                      sInput = "TODOS"

                                 CASE ELSE

                                   strsql = "Select * From USUARIOS where ID = '" & TABVISTA(II,YY) & "' "
                                   set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                                   if not rsniv.eof then sInput = rsniv(1)
                                   call ofv.cerrarconsulta(rsniv)
                          END SELECT
                          sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sInput,"readonly","","70","dt","")

                      end if

                CASE "3" 
                      WKFNOMDATO = "T.REQUERIMIENTO"

                      if bModificar then

                         strsql = "Select * From TIPOREQNV Order By DESCRIPCION"
 ''  response.write strsql & "<br>" & chr(13)
                         set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         DO UNTIL rsniv.eof
                            scomK = scomK & rsniv(1) & chr(9)
                            scomD = scomD & rsniv(2) & chr(9)
                         RSNIV.MOVENEXT
                         LOOP
                         call ofv.cerrarconsulta(rsniv)

                         sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")

                      else


                          SELECT CASE TABVISTA(II,YY)

                                 CASE "TODOS"
                                      sniv = "TODOS"

                                 CASE ELSE

                                  strsql = "Select * From TIPOREQNV where TIPOREQCOD = '" & TABVISTA(II,YY) & "' "
                                  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                  if not rsniv.eof then sniv = rsniv(2)
                                  call ofv.cerrarconsulta(rsniv)
                          END SELECT

                          sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sniv,"readonly","","70","dt","")
                      end if
                CASE "4" 
                      WKFNOMDATO = "ESTADO"

                      if bModificar then

                         strsql = "Select * From GPestados Order By DESCRI"
                         set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         DO UNTIL rsniv.eof
                            scomK = scomK & rsniv(0) & chr(9)
                            scomD = scomD & rsniv(1) & chr(9)
                         RSNIV.MOVENEXT
                         LOOP
                         call ofv.cerrarconsulta(rsniv)

                         sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")

                      else
                          SELECT CASE TABVISTA(II,YY)

                                 CASE "TODOS"
                                      sniv = "TODOS"

                                 CASE ELSE

                                    strsql = "Select * From GPestados where OBJID = '" & TABVISTA(II,YY) & "' "
                                    set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                    if not rsniv.eof then sniv = rsniv(1)
                                    call ofv.cerrarconsulta(rsniv)
                          END SELECT


                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sniv,"readonly","","70","dt","")
                      end if

                CASE "5" 
                      WKFNOMDATO = "FEC.INICIO"
                      if  bModificar  then

                          scomK = scomK & "-U7-" & chr(9) & "-U15-" & chr(9) & "-UM-" & chr(9) & "-UB-" & chr(9) & "-US-" & chr(9) & "-UA-"
                          scomD = scomD & "ULT.7 DIAS" & chr(9) & "ULT.15 DIAS" & chr(9) & "ULT.MES" & chr(9) & "ULT.BIMESTRE" & chr(9) & "ULT.SEMESTRE" & chr(9) & "ULT.AÑO"

                      	  sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")
     
                      else

                          SELECT CASE TABVISTA(II,YY)

                                 CASE "-U7-"
                                      sInput = "ULT.7 DIAS"

                                 CASE "-U15-"
                                      sInput = "ULT.15 DIAS"

                                 CASE "-UM-"
                                      sInput = "ULT.MES"

                                 CASE "-UB-"
                                      sInput = "ULT.BIMESTRE"

                                 CASE "-US-"
                                      sInput = "ULT.SEMESTRE"

                                 CASE "-UA-"
                                      sInput = "ULT.AÑO"

                                 CASE ELSE
                                      sInput = "TODOS"
                          END SELECT
                          sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sInput,"readonly","","70","dt","")

                      end if


                CASE "6" 
                      WKFNOMDATO = "FEC.FIN"
                              
            
                      if  bModificar  then

                          scomK = scomK & "-U7-" & chr(9) & "-U15-" & chr(9) & "-UM-" & chr(9) & "-UB-" & chr(9) & "-US-" & chr(9) & "-UA-"
                          scomD = scomD & "ULT.7 DIAS" & chr(9) & "ULT.15 DIAS" & chr(9) & "ULT.MES" & chr(9) & "ULT.BIMESTRE" & chr(9) & "ULT.SEMESTRE" & chr(9) & "ULT.AÑO"

                      	  sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")
     
                      else

                          SELECT CASE TABVISTA(II,YY)

                                 CASE "-U7-"
                                      sInput = "ULT.7 DIAS"

                                 CASE "-U15-"
                                      sInput = "ULT.15 DIAS"

                                 CASE "-UM-"
                                      sInput = "ULT.MES"

                                 CASE "-UB-"
                                      sInput = "ULT.BIMESTRE"

                                 CASE "-US-"
                                      sInput = "ULT.SEMESTRE"

                                 CASE "-UA-"
                                      sInput = "ULT.AÑO"

                                 CASE ELSE
                                      sInput = "TODOS"
                          END SELECT
                          sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sInput,"readonly","","70","dt","")

                      end if

                CASE "7" 
                      WKFNOMDATO = "SISTEMA"
                      
                      if bModificar then

                         strsql = "Select * From GPSISTEMAS Order By DESCRI"
                         set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         DO UNTIL rsniv.eof
                            scomK = scomK & rsniv(1) & chr(9)
                            scomD = scomD & rsniv(2) & chr(9)
                         RSNIV.MOVENEXT
                         LOOP
                         call ofv.cerrarconsulta(rsniv)

                         sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")

                      else
                          SELECT CASE TABVISTA(II,YY)

                                 CASE "TODOS"
                                      sniv = "TODOS"

                                 CASE ELSE
                                    strsql = "Select * From GPSISTEMAS where CODSIST = '" & TABVISTA(II,YY) & "' "
                                    set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                    if not rsniv.eof then sniv = rsniv(2)
                                    call ofv.cerrarconsulta(rsniv)
                          END SELECT

                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sniv,"readonly","","70","dt","")
                      end if


                CASE "8" 
                      WKFNOMDATO = "MODULO"
                    
                      if bModificar then

                         strsql = "Select * From GPMODULOS Order By DESCRI"
                         set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                         DO UNTIL rsniv.eof
                            scomK = scomK & rsniv(2) & chr(9)
                            scomD = scomD & rsniv(3) & chr(9)
                         RSNIV.MOVENEXT
                         LOOP
                         call ofv.cerrarconsulta(rsniv)

                         sInput2 = ofv.gencomboFX(sForm,WKCODCAMPO & WKII,TABVISTA(II,YY),scomK,scomD," ")

                      else
                          SELECT CASE TABVISTA(II,YY)

                                 CASE "TODOS"
                                      sniv = "TODOS"

                                 CASE ELSE
                                  strsql = "Select * From GPMODULOS where CODMOD = '" & TABVISTA(II,YY) & "' "
                                  set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                  if not rsniv.eof then sniv = rsniv(2)
                                  call ofv.cerrarconsulta(rsniv)
                          END SELECT

                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,sniv,"readonly","","70","dt","")
                      end if



                CASE "9" 
                      WKFNOMDATO = "CASO DE USO"

                      if bModificar then
                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,TABVISTA(II,YY),"text",SFORM,"70","dt","")
                      else
                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,TABVISTA(II,YY),"readonly","","70","dt","")
                      end if


                CASE ELSE

                      if bModificar then
                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,TABVISTA(II,YY),"text",SFORM,"70","dt","")
                      else
                         sInput2 = ofv.GenerarInput(WKCODCAMPO & WKII,TABVISTA(II,YY),"readonly","","70","dt","")
                      end if



         END SELECT

            WKFSTYLER = "AAA "

         IF II = 9 THEN

            WKFSTYLER = "AAA STYLE = 'DISPLAY:NONE;' "

         END IF 



             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")
             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filasx1 = filasx1 & ofv.GenRow("",WKFSTYLER,"","","10","","",celdas,"si")



         NEXT


         WKFNOMDATO = TABVISTA(II,0)

         SELECT CASE CSTR(II)

                CASE "1" 
                      WKFNOMDATO = "ORIGINANTE"

                CASE "2" 
                      WKFNOMDATO = "RESPONSABLE"

                CASE "3" 
                      WKFNOMDATO = "T.REQUERIMIENTO"

                CASE "4" 
                      WKFNOMDATO = "ESTADO"

                CASE "5" 
                      WKFNOMDATO = "FEC.INICIO"

                CASE "6" 
                      WKFNOMDATO = "FEC.FIN"

                CASE "7" 
                      WKFNOMDATO = "SISTEMA"

                CASE "8" 
                      WKFNOMDATO = "MODULO"

                CASE "9" 
                      WKFNOMDATO = "CASO DE USO"


                CASE ELSE
                      WKFNOMDATO = TABVISTA(II,0)


         END SELECT

         wkprocbotmas = ""
         wkprocbotmen = ""

         if (ii > 6 or ii < 5) AND (TABVISTA(II,1) <> "TODOS" ) then

            if TABVISTA(II,10) < 5 then

               sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='MASD" & II & "'; document." & sform & ".submit(); return false;" & CHR(34)

               wkprocbotmas = "<INPUT CLASS=NT " & sFUNCION & " SIZE=1 style='COLOR:#ffffff;background-color:#009900;border-style:outset;border-width:2;CURSOR:HAND;' TITLE='Agregar dato.' VALUE=' + ' >"


            end if 


            if TABVISTA(II,10) > 1 then

               sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='MEND" & II & "'; document." & sform & ".submit(); return false;" & CHR(34)

               wkprocbotmen = "<INPUT CLASS=NT " & sFUNCION & " SIZE=1 style='COLOR:#ffffff;background-color:#993333;border-style:outset;border-width:2;CURSOR:HAND;' TITLE='Eliminar dato.' VALUE=' - ' >"

            end if 


         end if 



         sInput = WKFNOMDATO & "&nbsp;:&nbsp;"

         celdas = ""

         celdas = celdas & ofv.GenCelda("","UtI style='background-color:#336699;color:#ffffff;' ","","","","","",sInput,"si")



         sInput = ofv.GenerarInput(WKCODCAMPO & "-00",TABVISTA(II,0),"HIDDEN","","70","dt","") & "&nbsp;" & wkprocbotmas & "&nbsp;" & wkprocbotmen

         celdas = celdas & ofv.GenCelda("","Ut style='background-color:#F0F0F0;border-BOTTOM-style:solid;border-BOTTOM-width:1;BORDER-BOTTOM-COLOR:#AAAAAA;'  ","","","","","",sInput,"si")



     ''    sInput = "&nbsp;" & TABVISTA(II,10)
         sInput = "&nbsp;"

         celdas = celdas & ofv.GenCelda("","Ut colspan=4 style='background-color:#F0F0F0;border-BOTTOM-style:solid;border-BOTTOM-width:1;BORDER-BOTTOM-COLOR:#AAAAAA;' ","","","","","",sInput,"si")
         filas = filas & ofv.GenRow("",WKFSTYLER,"","","22","","",celdas,"si")

'''border-TOP-style:solid;border-TOP-width:1;BORDER-TOP-COLOR:#AAAAAA;

         sInput = "&nbsp;"
         celdas = ofv.GenCelda("","Ut  style='font-size:5pt;' ","","","","","",sInput,"si")
         celdas = celdas & ofv.GenCelda("","Ut colspan=5 style='font-size:5pt;' ","","","","","",sInput,"si")
         filasX1 = filasX1 & ofv.GenRow("",WKFSTYLER,"","","","","",celdas,"si")

         filas = filas & filasx1

         filasx1 = ""



     NEXT



    celdas = ""

    wkfmodoact = "NORMAL"

    sInput = ofv.GenerarInput("MODOACTUAL",wkfmodoact,"hidden",sform,"10","dt","") & "&nbsp;"

    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")

    sInput = "&nbsp;"
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","15%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")


    filas = filas & ofv.GenRow("","","","","5","","",celdas,"si")


   wkproccod1 = "&nbsp;"
   wkproccod2 = "&nbsp;"
   wkproccod3 = "&nbsp;"
   wkproccod4 = "&nbsp;"

   wkproccod1S = ""
   wkproccod2S = ""
   wkproccod3S = ""
   wkproccod4S = ""

''''''alert('hola');



     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='TMP'; document." & sform & ".submit(); return false;" & CHR(34)

     wkproccod1 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Usar como vista temporal.' HREF='" & sAccion & "' >VISTA TEMPORAL</A>"

     wkproccod1S = " style='background-color:#336699;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "

   IF SESSION("MODACT") = "S" AND WKFPERVN = "S" THEN

     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='VACT'; document." & sform & ".submit(); return false;" & CHR(34)
  ''   sFUNCION = ""


     wkproccod2 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Actualizar la vista actual.' HREF='" & sAccion & "' >ACTUALIZAR VISTA ACTUAL</A>"

     wkproccod2S = " style='background-color:#F7AA22;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "

   END IF

   IF WKFPERVN = "S" THEN

     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='VNVA'; document." & sform & ".submit(); return false;" & CHR(34)
  ''   sFUNCION = ""

     wkproccod3 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Guardar como nueva vista.' HREF='" & sAccion & "' >GUARDAR COMO NUEVA</A>"

     wkproccod3S = " style='background-color:#AA0000;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "

   END IF




    celdas = ""

    sInput = "&nbsp;"

    celdas = celdas & ofv.GenCelda("","Ut colspan=3 style='background-color:#dddddd;color:#888888;' ","","","","","",sInput,"si")

    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod2S,"CENTER","","","","",wkproccod2,"si") 
    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod3S,"CENTER","","","","",wkproccod3,"si") 
    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod1S,"CENTER","","","","",wkproccod1,"si") 

    filas = filas & ofv.GenRow("","AAA BGCOLOR='#dddddd' ","","","25","","",celdas,"si")





    sAccion = "GPLISTAREQvista.asp?MODAL=GETACT"
    Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",filas,"si")

    sHTML = sHTML & "<CENTER><BR>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","70% ","","0","","0","0","","",Formulario,"si")
    sHTML = sHTML & "</CENTER>" 
    WKFsHTML = sHTML



  END IF 

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''' ACTUALIZACION DE PARAMETROS DE LAS VISTA Y OPCIONES DE GUARDADO '''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if sMODAL = "SIGACTN" OR  sMODAL = "SIGACTV" then

    filas = ""


   celdas = ""

    sInput = "&nbsp;"

    celdas = celdas & ofv.GenCelda("","VTI colspan=3 style='color:#FFFFFF;' ","","","","","","ACTUALIZACION DE VISTA","si")

    CeldaS = CeldaS & ofv.GenCelda("","utI  style='BACKGROUND-COLOR:#DDDDDD;color:#888888;' ","RIGHT","","","","","VISTA:&nbsp;","si") 
    CeldaS = CeldaS & ofv.GenCelda("","utI COLSPAN=2 style='color:#FFFFFF;' ","","","","","",WKFVISTANOM,"si") 


    filas = filas & ofv.GenRow("","AAA BGCOLOR='#888888' ","","","25","","",celdas,"si")

''''''''''''''''''''''''' DATOS '''''''''''''''''''''''''''''

           '  WKFFRMCODVISTA = RSX1(1)
           '  WKFFRMFCARGA   = RSX1(2)
           '  WKFFRMFMODIF   = WKHOY
           '  WKFFRMESTADO   = RSX1(4)
           '  WKFFRMUSUARIO  = RSX1(5)
           '  WKFFRMTIPO     = RSX1(6)
           '  WKFFRMDETVISTA = WKTXTVISTA
           '  WKFFRMNOMVISTA = RSX1(9)
           '  WKFFRMESDEFAULT = RSX1(10)
           '  WKFFRMESPUBLIC = RSX1(11)



             sInput = "ESTADO"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Uti  ","","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")

             scomK = "S" & chr(9) & "N"
             scomD = "SI" & chr(9) & "NO"
            
             sInput2 = ofv.gencomboFX(sForm,"ESTADO",WKFFRMESTADO,scomK,scomD,"no")

             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


            sInput = "USUARIO"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Uti  ","","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")


             sInput2 = WKFFRMUSUARIO


             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

            sInput = "NOMBRE"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Uti  ","","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")

             sInput2 = ofv.GenerarInput("NOMVISTA",WKFFRMNOMVISTA,"agregar",SFORM,"70","dt","")


             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

            sInput = "ES DEFAULT"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Uti  ","","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")

             scomK = "S" & chr(9) & "N"
             scomD = "SI" & chr(9) & "NO"
            
             sInput2 = ofv.gencomboFX(sForm,"ESDEFAULT",WKFFRMESDEFAULT,scomK,scomD,"no")

             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

            sInput = "ES PUBLICA"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Uti  ","","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")

             scomK = "S" & chr(9) & "N"
             scomD = "SI" & chr(9) & "NO"
            
             sInput2 = ofv.gencomboFX(sForm,"ESPUBLIC",WKFFRMESPUBLIC,scomK,scomD,"no")

             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

            sInput = "&nbsp;"

             celdas = ""

             celdas = celdas & ofv.GenCelda("","Uti  ","","","","","",sInput,"si")

             sInput1 = "&nbsp;"

             celdas = celdas & ofv.GenCelda("","vt style='color:#aa0000;'  ","right","","","","",sInput1,"si")

             sInput2 = ""
             sInput2 = sInput2 & ofv.GenerarInput("CODVISTA",WKFFRMCODVISTA,"text",SFORM,"70","dt","")
             sInput2 = sInput2 & ofv.GenerarInput("FCARGA",WKFFRMFCARGA,"text",SFORM,"70","dt","")
             sInput2 = sInput2 & ofv.GenerarInput("FMODIF",WKFFRMFMODIF,"text",SFORM,"70","dt","")
             sInput2 = sInput2 & ofv.GenerarInput("USUARIO",WKFFRMUSUARIO,"text",SFORM,"70","dt","")
             sInput2 = sInput2 & ofv.GenerarInput("TIPO",WKFFRMTIPO,"text",SFORM,"70","dt","")
             sInput2 = sInput2 & ofv.GenerarInput("DETVISTA",WKFFRMDETVISTA,"text",SFORM,"70","dt","")


             celdas = celdas & ofv.GenCelda("","Ut colspan=4 ","","","","","",sInput2,"si")
             filas = filas & ofv.GenRow("","AAA STYLE='DISPLAY:NONE;' ","","","10","","",celdas,"si")


           '  WKFFRMCODVISTA = RSX1(1)
           '  WKFFRMFCARGA   = RSX1(2)
           '  WKFFRMFMODIF   = WKHOY
           '  WKFFRMESTADO   = RSX1(4)
           '  WKFFRMUSUARIO  = RSX1(5)
           '  WKFFRMTIPO     = RSX1(6)
           '  WKFFRMDETVISTA = WKTXTVISTA
           '  WKFFRMNOMVISTA = RSX1(9)
           '  WKFFRMESDEFAULT = RSX1(10)
           '  WKFFRMESPUBLIC = RSX1(11)



''''''''''''''''''''''''' FIN DATOS '''''''''''''''''''''''''''

    celdas = ""

    wkfmodoact = "NORMAL"

    sInput = ofv.GenerarInput("MODOACTUAL",wkfmodoact,"hidden",sform,"10","dt","") & "&nbsp;"

    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")

    sInput = "&nbsp;"
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","15%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")


    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


   wkproccod1 = "&nbsp;"
   wkproccod2 = "&nbsp;"
   wkproccod3 = "&nbsp;"
   wkproccod4 = "&nbsp;"

   wkproccod1S = ""
   wkproccod2S = ""
   wkproccod3S = ""
   wkproccod4S = ""

''''''alert('hola');



     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='AYC'; document." & sform & ".submit(); return false;" & CHR(34)

     wkproccod1 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Actualizar la vista y consultar.' HREF='" & sAccion & "' >ACTUALIZAR Y CONSULTAR</A>"

     wkproccod1S = " style='background-color:#336699;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "



     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='AYL'; document." & sform & ".submit(); return false;" & CHR(34)
  ''   sFUNCION = ""


     wkproccod2 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Actualizar la vista e ir a la lista.' HREF='" & sAccion & "' >ACTUALIZAR Y LISTA</A>"

     wkproccod2S = " style='background-color:#F7AA22;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "





     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='KNC'; document." & sform & ".submit(); return false;" & CHR(34)
  ''   sFUNCION = ""

     wkproccod3 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Cancelar actualizacion.' HREF='" & sAccion & "' >CANCELAR</A>"

     wkproccod3S = " style='background-color:#AA0000;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "




    celdas = ""

    sInput = "&nbsp;"

    celdas = celdas & ofv.GenCelda("","Ut colspan=3 style='background-color:#dddddd;color:#888888;' ","","","","","",sInput,"si")

    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod2S,"CENTER","","","","",wkproccod2,"si") 
    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod3S,"CENTER","","","","",wkproccod3,"si") 
    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod1S,"CENTER","","","","",wkproccod1,"si") 

    filas = filas & ofv.GenRow("","AAA BGCOLOR='#dddddd' ","","","25","","",celdas,"si")





    sAccion = "GPLISTAREQvista.asp?MODAL=PROCACT"
    Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",filas,"si")

    sHTML = sHTML & "<CENTER><BR>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","70% ","","0","","0","0","","",Formulario,"si")
    sHTML = sHTML & "</CENTER>" 
    WKFsHTML = sHTML


  END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''' PANEL DE LISTA DE VISTAS GENERAL ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if sMODAL = "LISVISTA"  then


    filas = ""


   celdas = ""

    sInput = "&nbsp;"

    celdas = celdas & ofv.GenCelda("","VTI colspan=3 style='color:#FFFFFF;' ","","","","","","LISTA DE VISTAS","si")

    CeldaS = CeldaS & ofv.GenCelda("","utI COLSPAN=2 style='BACKGROUND-COLOR:#DDDDDD;color:#888888;' ","RIGHT","","","","","VISTA&nbsp;ACTUAL:&nbsp;","si") 
    CeldaS = CeldaS & ofv.GenCelda("","utI COLSPAN=2 style='color:#FFFFFF;' ","","","","","",WKFVISTANOM,"si") 


    filas = filas & ofv.GenRow("","AAA BGCOLOR='#888888' ","","","25","","",celdas,"si")



        celdas = ""

        celdas = celdas & ofv.GenCelda("","Ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","","","","","","VISTA","si")

        sInput = "&nbsp;"

        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","","PROPIA","si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","","DEFAULT","si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","","PUBLICA","si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","","ACTIVA","si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","",sInput,"si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","",sInput,"si") 

        filas = filas & ofv.GenRow("","AAA BGCOLOR='#888888' ","","","20","","",celdas,"si")



     Strsql = ""
     Strsql = Strsql & "Select A.CODVISTA, A.NOMVISTA, A.ESDEFAULT, A.ESPUBLIC, A.ESTADO, A.USUARIO, "
     Strsql = Strsql & " CASE A.USUARIO "
     Strsql = Strsql & "    WHEN '" & SESSION("USUARIO") & "' THEN '0' "
     Strsql = Strsql & "    ELSE '1' "
     Strsql = Strsql & " END AS PROPIA  "
     Strsql = Strsql & "  From GPVISTAS A"
     Strsql = Strsql & " LEFT JOIN GPVISTASARC B ON (A.CODVISTA = B.CODVISTA) "
     Strsql = Strsql & " WHERE A.TIPO = 'REQ' AND "
     Strsql = Strsql & " ( "
     Strsql = Strsql & " A.ESPUBLIC = 'S' OR "
     Strsql = Strsql & " (A.ESPUBLIC = 'N' AND B.USUARIOSECTOR = '" & SESSION("USUARIO") & "' ) OR "
     Strsql = Strsql & " (A.ESPUBLIC = 'N' AND A.USUARIO = '" & SESSION("USUARIO") & "' )  "
     Strsql = Strsql & " ) "
     Strsql = Strsql & " ORDER BY  PROPIA, A.ESDEFAULT DESC, A.ESPUBLIC, A.NOMVISTA  "



     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     IF NOT RSX1.EOF THEN
        Do While (Not RSX1.EOF) 


             WKFFRMCODVISTA  = RSX1(0)
             WKFFRMESTADO    = RSX1(4)
             WKFFRMUSUARIO   = RSX1(5)
             WKFFRMNOMVISTA  = RSX1(1)
             WKFFRMESDEFAULT = RSX1(2)
             WKFFRMESPUBLIC  = RSX1(3)



             if  WKFFRMESTADO    = "S" THEN 
                 WKFFRMESTADO    = "<FONT COLOR='#AA0000'><B>SI</B></FONT>"
             ELSE
                 WKFFRMESTADO    = "NO"
             END IF 

             if  WKFFRMESDEFAULT    = "S" THEN 
                 WKFFRMESDEFAULT    = "<FONT COLOR='#AA0000'><B>SI</B></FONT>"
             ELSE
                 WKFFRMESDEFAULT    = "NO"
             END IF 

             if  WKFFRMESPUBLIC    = "S" THEN 
                 WKFFRMESPUBLIC    = "<FONT COLOR='#AA0000'><B>SI</B></FONT>"
             ELSE
                 WKFFRMESPUBLIC    = "NO"
             END IF 

             if  WKFFRMUSUARIO    = SESSION("USUARIO") THEN 
                 WKFFRMUSUARIO    = "<FONT COLOR='#AA0000'><B>SI</B></FONT>"
             ELSE
                 WKFFRMUSUARIO    = "NO"
             END IF 





        celdas = ""

        sInput = "<U>" & RSX1(1) & "</U>"


        sAccion = "GPLISTAREQvista.asp?MODAL=SETSEL&CVISTA=" & RSX1(0)

        sFUNCION = ""

        sInput = "<A CLASS=NT " & sFUNCION & "  TITLE='Seleccionar vista.' HREF='" & sAccion & "' >" & sInput & "</A>"



        celdas = celdas & ofv.GenCelda("","Ut  style='background-color:#F0F0F0;color:#888888;' ","","","","","",sInput,"si")

        sInput = "&nbsp;"

        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='background-color:#FAFAFA;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","",WKFFRMUSUARIO,"si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='background-color:#FAFAFA;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","",WKFFRMESDEFAULT,"si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='background-color:#FAFAFA;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","",WKFFRMESPUBLIC,"si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut STYLE='background-color:#FAFAFA;border-style:solid;border-width:1;border-color:#000000;' ","CENTER","","","","",WKFFRMESTADO,"si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut   ","CENTER","","","","",sInput,"si") 
        CeldaS = CeldaS & ofv.GenCelda("","ut   ","CENTER","","","","",sInput,"si") 

        filas = filas & ofv.GenRow("","AAA BGCOLOR='#F0F0F0' ","","","20","","",celdas,"si")

  

          RSX1.MOVENEXT
        LOOP
     END IF 

    celdas = ""

    wkfmodoact = "NORMAL"

    sInput = ofv.GenerarInput("MODOACTUAL",wkfmodoact,"hidden",sform,"10","dt","") & "&nbsp;"

    celdas = celdas & ofv.GenCelda("","Ut  ","","55%","","","",sInput,"si")

    sInput = "&nbsp;"
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","5%","","","",sInput,"si")
    celdas = celdas & ofv.GenCelda("","Ut  ","","20%","","","",sInput,"si")


    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


   wkproccod1 = "&nbsp;"
   wkproccod2 = "&nbsp;"
   wkproccod3 = "&nbsp;"
   wkproccod4 = "&nbsp;"

   wkproccod1S = ""
   wkproccod2S = ""
   wkproccod3S = ""
   wkproccod4S = ""

''''''alert('hola');



     sAccion = "javascript:void(0);"

     sFUNCION = "ONCLICK=" & CHR(34) & "document.all.MODOACTUAL.value='NORMAL'; document." & sform & ".submit(); return false;" & CHR(34)

     wkproccod1 = "<A CLASS=NT " & sFUNCION & " style='COLOR:#ffffff;' TITLE='Actualizar la vista y consultar.' HREF='" & sAccion & "' >VOLVER A VISTA ACTUAL</A>"

     wkproccod1S = " style='background-color:#336699;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "





    celdas = ""

    sInput = "&nbsp;"

    celdas = celdas & ofv.GenCelda("","Ut colspan=4 style='background-color:#dddddd;color:#888888;' ","","","","","",sInput,"si")

    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod2S,"CENTER","","","","",wkproccod2,"si") 
    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod3S,"CENTER","","","","",wkproccod3,"si") 
    CeldaS = CeldaS & ofv.GenCelda("","ut   " & wkproccod1S,"CENTER","","","","",wkproccod1,"si") 

    filas = filas & ofv.GenRow("","AAA BGCOLOR='#dddddd' ","","","25","","",celdas,"si")

    sAccion = "GPLISTAREQvista.asp?MODAL=SET"
    Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",filas,"si")

    sHTML = sHTML & "<CENTER><BR>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","95% ","","0","","0","0","","",Formulario,"si")
    sHTML = sHTML & "</CENTER>" 
    WKFsHTML = sHTML



  END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''' FINAL: OPCIONES DE SALIDA DEL MODULO ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  if sMODAL = "SAVEACT"  then

    sHTML = ""


    RESPONSE.REDIRECT "GPLISTAREQ.ASP?MODAL=SET&VISTA=" & WKFVISTA


  '''  sHTML = sHTML & "<br><p style='border-style:solid;border-width:1;line-height:20px;' >&nbsp;" &  WKTXTWHERE & "&nbsp;</p>"


  ELSEif sMODAL = "VISTATEMP"  then

    sHTML = ""


    RESPONSE.REDIRECT "GPLISTAREQ.ASP?MODAL=VISTATEMP&VISTA=" & WKFVISTA


  '''  sHTML = sHTML & "<br><p style='border-style:solid;border-width:1;line-height:20px;' >&nbsp;" &  WKTXTWHERE & "&nbsp;</p>"


  ELSE

     sHTML = ofv.FormHeader(session("FormN"))

     sHTML = sHTML & WKFsHTML

     Agregar = ""
     Buscar  = ""

     sAccion = "GPLISTAREQ.asp?MODAL=SET"
     Celda = ofv.BotonVolver(sAccion)


     sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,Buscar,celda)


  END IF




  call ofv.cerrarconn(cn)

end if

Response.Write sHTML



%>