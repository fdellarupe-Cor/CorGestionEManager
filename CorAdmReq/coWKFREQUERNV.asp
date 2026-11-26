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
  act     = request.querystring("act")
  sACCION = request.querystring("ACCION")
  smodoasg = request.querystring("MODOASG")
  smodoUSR = request.querystring("MODOUSR")


  sopx = sopcion
  sHTML = ofv.FormHeader("")


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_OPCION=" & SOPCION & "_ACT=S"

  WKFAUTO = "WKF0001693621131313343163"
  WKFAPRO = "WKF0001693673726313343163"
  DESLEVL = 9
  SUPLEVL = 9


  WKNIVSOLTIP = "USR"
  WKNIVSOLCOD = session("USUARIO")

  WKNIVSOL = 1
  WKNIVDES = 2

  ITRSP = 0
  TTRSP = 0

  dim ATRSPNIV(50,12)
  dim ATRSPNIVUSR(50,50)


  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
  if not rs.eof then
     SESSION("NOMREQ") = "REQUERIMIENTO " & RS(2) 
     WTIPOREQ =  RS(15)

     WKNIVSOLTIP = "USR"
     WKNIVSOLCOD = RS(13)

     WKNUSRORIG  = RS(13)
     WKNUSRDEST  = RS(14)

     Strsql = "Select * From TIPOREQNV where TIPOREQCOD = '" & WTIPOREQ & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        WKFAUTO = RS(5)
        DESLEVL = RS(12)
        SUPLEVL = RS(10)
        WKFDSC  = RS(2)

        WKNIVDESTIP = "SEC"
        WKNIVDESCOD = RS(11)

        WKREQCRPT   = RS(3)
        WKREQCTAD   = RS(4)
        WKREQPRY    = "N"

        IF WKREQCRPT = "S" THEN

           StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
           set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if NOT rsX1.eof then
              WKREQPRY    = "S"

           END IF

        END IF

     END IF

  END IF



  if nFecini < clng(cdate("01-01-1900")) then
     nFecini = clng(cdate("01-01-1900"))
  else
     nFecini = 0
  end if

  vNuevoValor = clng(DATE()) - nFecini








 DIM ATABLA(200,5) 

 DIM ATABSW(5,3) 

 TSHW = 0

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


    Strsql = "SELECT * FROM CKTVIEWnv WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsx.eof then

       session("WKFPSC")    = RSX(17)
       SESSION("SEGCKT")    = RSX(5)
       SESSION("WKF")       = RSX(0)
       SESSION("ORGUSR")    = RSX(26)


       set cn2 = ofv.conectar(ofv.strconn0)
       sInputx = "CORONTIME@corconsulting.com.ar"
       strsql = "Select EMAIL From USUARIOS where ID = '" & session("usuario") & "' "
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


  StrSql = "Select * From SEGNIVELESNV where SEGCKT = '" & SESSION("SEGCKT") & "' ORDER BY SEGNIVEL "
  set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
  if not rs.eof then
     DO WHILE NOT RS.EOF
        TTRSP = TTRSP + 1

        IF RS(07) = "SOL" THEN  WKNIVSOL = RS(04)
        IF RS(07) = "DES" THEN  WKNIVDES = RS(04)


        ATRSPNIV(TTRSP,01) = RS(04)  'COD. NIVEL
        ATRSPNIV(TTRSP,02) = RS(10)  'DESCRIPCION
        ATRSPNIV(TTRSP,03) = RS(05)  'MODO
        ATRSPNIV(TTRSP,04) = RS(06)  'ESTILO
        ATRSPNIV(TTRSP,05) = RS(09)  'SUPER.NIVEL
        ATRSPNIV(TTRSP,06) = "X"  'AUTOESCALA
        ATRSPNIV(TTRSP,07) = RS(07)  'TIPO RSP.
        ATRSPNIV(TTRSP,08) = RS(08)  'COD. RSP.


        ATRSPNIVUSR(TTRSP,00) = 0  'nro. de responsables del nivel.


     RS.MOVENEXT
     LOOP

  END IF

 ''' RESPONSE.WRITE SOPCION & "<BR>"

            SIGUEMR = "N" 

 IF SOPCION = "INITMP" OR SOPCION = "INITMR"  OR SOPCION = "INITMO" THEN

'''''''''''''''''''''''''''''' OPCIONES DE INICIO DEL CIRCUITO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 ''' RESPONSE.WRITE SOPCION & "<BR>"


   StrSql = "Select * From TIPOREQNIVNV where TIPOREQCOD = '" & WTIPOREQ & "'  ORDER BY ITEMCOD "
   set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
   if not rs.eof then
        TTRSP = 0
      DO WHILE NOT RS.EOF
        TTRSP = TTRSP + 1

        IF RS(10) = "SOL" THEN  WKNIVSOL = RS(04)
        IF RS(10) = "DES" THEN  WKNIVDES = RS(04)


        ATRSPNIV(TTRSP,01) = RS(04)  'COD. NIVEL
        ATRSPNIV(TTRSP,02) = RS(05)  'DESCRIPCION
        ATRSPNIV(TTRSP,03) = RS(06)  'MODO
        ATRSPNIV(TTRSP,04) = RS(07)  'ESTILO
        ATRSPNIV(TTRSP,05) = RS(08)  'SUPER.NIVEL
        ATRSPNIV(TTRSP,06) = RS(09)  'AUTOESCALA
        ATRSPNIV(TTRSP,07) = RS(10)  'TIPO RSP.
        ATRSPNIV(TTRSP,08) = RS(11)  'COD. RSP.


      RS.MOVENEXT
      LOOP

   END IF


   SESSION("WKF")    = WKFAUTO
   TITULO            = WKFDSC
   SESSION("WKFDSC") = WKFDSC


'''''''''''' GENERA EL PASO INICIAL

   StrSql = "Select * From WKFPASOSNV WHERE WKFCOD = '" & SESSION("WKF") & "' ORDER BY ORDEN"
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

   if not rs.eof then

      StrSql = "Select * From WKFPASOSPRTYNV WHERE WKFCOD = '" & SESSION("WKF") & "' AND "
      StrSql = StrSql & " WKFPASCOD = '" & RS(2) & "' AND ITEMCOD = 'INI'"
      set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rsX.eof then

         XPASCOD    = RSX(7)
         XPASACC    = RSX(9)
         XPASEST    = RSX(10)

      END IF


      IF SOPCION = "INITMR"   THEN



         StrSql = "Select * From WKFPASOSNV WHERE WKFCOD = '" & SESSION("WKF") & "' AND "
         StrSql = StrSql & " WKFACCION = 'AR' ORDER BY ORDEN "
         set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rsX.eof then

            act = "S"

            SIGUEMR = "S"  

            XPASCOD    = RSX(2)

         END IF
   
      else


      end if



''''' INICIA EL CIRCUITO 

      XMODEL     = SESSION("WKF")
      XMODELPSC  = RS(2)


      XSINST     = 1
      XSCICLO    = RS(22)
      XSNIVEL    = RS(21)
      XSSUBN     = 1

      XMODO      = "U"
      XESTILO    = "F"
      XNODO      = "INI"



      XRTIPO     = WKNIVSOLTIP
      XRCOD      = WKNIVSOLCOD
      XRASG      = WKNIVSOLCOD
      XRLVL      = SUPLEVL

      XPDSC      = RS(3) & ": " & RS(4) 
      XPDSC1     = RS(3) & ": " & RS(4)  

      if act = "S" then

         set cn2 = ofv.conectar(ofv.strconn0)
         sInputx = "COR ON-TIME"
         strsql = "Select EMAIL From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
         set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
         if not rsniv.eof then sInputx = rsniv(0)
         call ofv.cerrarconsulta(rsniv)

         SESSION("ORGMAIL") = SINPUTX

         GENCKTINI()
         SEGCKTPSC = GTCKTPSC()

      IF SOPCION = "INITMR"   THEN

         XOBSERV = "CIRCUITO INICIADO POR EL RESPONSABLE"
         XSEGCODIGO = WKNIVSOLCOD

      else

         XOBSERV = REQUEST.FORM("OBS")
         XSEGCODIGO = SESSION("USUARIO")

      end if


'''''''''''''''''''''''''''''''''''''''''''
         SEGCKTPSCSIG = GTCKTPSC() & "A"
         WXX = GENCKTPSCNV(SEGCKTPSC,SEGCKTPSCSIG,XMODEL,XMODELPSC,XSINST,XSCICLO,XSNIVEL,XSSUBN,XMODO,XESTILO,XNODO,"SI",XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XRLVL)
         SESSION("SEGCKTPSC") = SEGCKTPSC

         if XPASACC = "A" then  SSS = GENACCIONNV(XPASACC,XPASEST)


'''''''''''''''''''''''''''''''''''''''''''''
''''' ACTUALIZA USUARIO ASIGNADO PARA EL NIVEL
''''''''''''''''''''''''''''''''''''''''''''''


         wkrsporden = 10
         wkrspcarga = "S" 

	 	           Strsql = "SELECT * FROM SEGCKTRESPNV WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND "
 	                   StrSql = StrSql & " SEGNIVEL = " & XSNIVEL & " AND SEGCODIGO = '" & XSEGCODIGO & "' "
	 	           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	           if NOT rsx.eof then
                              wkrspcarga = "N" 

                           ELSE

	 	              Strsql = "SELECT * FROM SEGCKTRESPNV WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND "
 	                      StrSql = StrSql & " SEGNIVEL = " & XSNIVEL & " ORDER BY SEGRSPSECUEN DESC "
	 	              set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	              if NOT rsx.eof then
                                 wkrsporden = rsx(8) + 10
                              END IF

		              SEGCKTRSP = GTCKTPSC()
                              wkrspSECTOR = 0
                              wkrspNIVEL = 8

	                      set cn2 = ofv.conectar(ofv.strconn0)
	                      strsql = "Select SECTOR,NIVEL From USUARIOS where ID = '" & XSEGCODIGO & "' "
	                      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	                      if not rsniv.eof then
                                 wkrspSECTOR = rsniv(0)
                                 wkrspNIVEL  = rsniv(1)
                              END IF 
	                      call ofv.cerrarconsulta(rsniv)


                              strsql = "insert into SEGCKTRESPNV ("
                              STRSQL = STRSQL & "SEGCKT,SEGNIVEL,SEGRSPCOD,SEGTIPO,"
                              STRSQL = STRSQL & "SEGCODIGO,SEGNIVELCNF,SEGRSPORDEN,SEGRSPSECUEN,"
                              STRSQL = STRSQL & "SEGMODOASG,SEGRAZONASG,SEGCODIGOASG,SEGESTADORSP,"
                              STRSQL = STRSQL & "SEGRSPDSC,SEGRSPSECTOR"
                              STRSQL = STRSQL & ") values ("
                              STRSQL = STRSQL & "'" & SESSION("SEGCKT") & "','" & XSNIVEL & "','" & SEGCKTRSP & "','USR',"
                              STRSQL = STRSQL & "'" & XSEGCODIGO & "','" & wkrspNIVEL & "','10','" & wkrsporden & "',"
                              STRSQL = STRSQL & "'WKF','RP','','OK',"
                              STRSQL = STRSQL & "'RESPONSABLE NIVEL " & XSNIVEL & "','" & wkrspSECTOR & "'"
                              STRSQL = STRSQL & ")"
                      ''       response.write strsql & "<br>" & chr(13)
                              call ofv.crearconsultaEx(StrSql,cn,1,parametros)

                           END IF 


         TSHW           = TSHW + 1
         ATABSW(TSHW,1) = XPDSC
         ATABSW(TSHW,2) = 1
         ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)


'''''''' VERIFICA PASO SIGUIENTE '''''''''''' 1- MISMO PASO (REQUIERE AUTORIZACION) 2- SIGUIENTE PASO 3- SIGUIENTE PASO (PASO RECURRENTE) 


'''' CASO 1 '''''''''''''''''''''''MISMO PASO (REQUIERE AUTORIZACION)''''''''''''''''''''''''


'''' CASO 2 ''''''''''''''''''''''''SIGUIENTE PASO (UNICO USUARIO)''''''''''''''''''''''''



         StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & XPASCOD & "' "
         set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rs2.eof then

            XPASCOD2    = "NO"
            XMODEL2     = SESSION("WKF")
            XMODELPSC2  = RS2(2)


            XSINST2     = 1
            XSCICLO2    = RS2(22)
            XSNIVEL2    = RS2(21)
            XSSUBN2     = 1

            XMODO2      = "U"
            XESTILO2    = "A"
            XNODO2      = "NO"

            XRTIPO2     = ATRSPNIV(XSNIVEL2,07)   'TIPO RSP.
            XRCOD2      = ATRSPNIV(XSNIVEL2,08)   'COD. RSP.

            if XRTIPO2 = "DES" THEN

               XRTIPO2     = WKNIVDESTIP
               XRCOD2      = WKNIVDESCOD

            ELSEif XRTIPO2 = "SOL" THEN

               XRTIPO2     = WKNIVSOLTIP
               XRCOD2      = WKNIVSOLCOD
               XRASG2      = WKNIVSOLCOD

            ELSEif XRTIPO2 = "USR" THEN

               XRASG2      = ATRSPNIV(XSNIVEL2,08)   'COD. RSP.

            END IF

            XRASG2      = "NO"


            IF SOPCION = "INITMR"  AND SIGUEMR = "S" THEN

               XRTIPO2     = WKNIVDESTIP
               XRCOD2      = WKNIVDESCOD
               XRASG2      = SESSION("USUARIO")

            end if


            XRLVL2      = SUPLEVL

            XPDSC2      = RS2(3) & ": " & RS2(4) 
            XPDSC12     = RS2(3) & ": " & RS2(4)  

 
'''''''''''''''''''''''''''''''''''''''''''''
''''''' ACTUALIZA USUARIO ASIGNADO PARA EL NIVEL
''''''''''''''''''''''''''''''''''''''''''''''
            WXX = GENCKTPSCNV(SEGCKTPSCSIG,"NO",XMODEL2,XMODELPSC2,XSINST2,XSCICLO2,XSNIVEL2,XSSUBN2,XMODO2,XESTILO2,XNODO2,"NO",XRTIPO2,XRCOD2,XRASG2,XPDSC2,XOBSERV2,XRLVL2)
            SESSION("SEGCKTPSC") = SEGCKTPSCSIG

    
            XXXX = GENCKTMAIL(XRTIPO,XRCOD,XPDSC & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL)

 '''''''''''''''''''

            TSHW           = TSHW + 1
            ATABSW(TSHW,1) = XPDSC2
            ATABSW(TSHW,2) = 0
            ATABSW(TSHW,3) = GETRESPOCKT(XRCOD2,XRTIPO2)

        END IF
        call ofv.cerrarconsulta(rs2)


''''''' CASO 3 ''''''''''''''''''''''''SIGUIENTE PASO (PASO RECURRENTE)'''''''''''''''''''''''''


''''''''''''''''''' ACTUALIZA EL PROYECTO ASOCIADO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        IF WKREQCRPT = "S" THEN

                           XXXX = GENACTPRY()

                        END IF

''''''''''''''''''' FIN ACTUALIZA EL PROYECTO ASOCIADO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    end if
'    call ofv.cerrarconsulta(rs)

 
   IF SOPCION = "INITMR" AND act = "S" AND SIGUEMR = "S" THEN


 

      WKFMODULO = "editarPRY.asp?codigo=" & session("NCliente") & "&tipo=REQ"
      WKFMODULO = WKFMODULO & "&SEGCKTPSC=" & SESSION("SEGCKTPSC") & "&MODAL=WK"

      sHTML = sHTML & "<form style='display:none;' name='reqwkf' action='" & WKFMODULO & "'  TARGET='_top' "
      sHTML = sHTML & "   method=post>"
      sHTML = sHTML & "</form>"

      sHTML = sHTML & "<script type='text/javascript'>"
      sHTML = sHTML & "    document.reqwkf.submit();"
      sHTML = sHTML & "</script>"



   ELSE



    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4so = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"
    sImagen4sA = "<B><FONT FACE=WEBDINGS COLOR=#BBBBBB SIZE=2>&#0128;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0174;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>G</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenut = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>&#0192;</FONT></B>"
    sImagendu = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenfi = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>R</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenES = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>&#0232;</FONT></B>"
    sImagenP = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>2</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=3>&#0237;</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"

    sInput = "SIN ASIGNAR"

    IF session("NCliente") <> "NO" THEN

       StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       if not rs.eof then
          sInput = "<B>" & RS(2) & "</B>"
          SESSION("NOMREQ") = "REQUERIMIENTO " & RS(2) 
       END IF

    END IF

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
  
    celdas = ofv.GenCelda("","ttt colspan=2","","","","","",sImagen & "&nbsp;REQUERIMIENTO " & sInput & " ","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt colspan=2 ","","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    set cn2 = ofv.conectar(ofv.strconn0)

    StrSql = "Select * From USUARIOS WHERE ID = '" & Session("Usuario") & "' "
    set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
    if not rs.eof THEN

       sInput = "ORIGINANTE:"
       celdas = ofv.GenCelda("","vT","","30%","","","",sInput,"si")

       sInput = sImagen4so & "&nbsp;<b>" & Session("SEGUsuario") & "</b>"
       celdas = celdas & ofv.GenCelda("","vt","","70%","","","",sInput,"si")

       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "GRUPO:"
       celdas = ofv.GenCelda("","vT","","","","","",sInput,"si")

       sgrp = ""   

       StrSql = "Select * From grupos WHERE ID = '" & Session("SEGGRP") & "' "
       set rs5 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rs5.eof THEN
          sgrp = rs5(1)
       end if

       sInput = sImagen6S & "&nbsp;<B>" & SESSION("SEGGRP") & " - " & sgrp & "</b>"
       celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "SECTOR:"
       celdas = ofv.GenCelda("","vT","","","","","",sInput,"si")

       sInput = "SIN ASIGNAR"

       set cn2 = ofv.conectar(ofv.strconn2)
           ON ERROR RESUME NEXT
           strsql = "Select * From ECO where ECO = " & rs(21)
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
       call ofv.cerrarconn(cn2)

       celdas = celdas & ofv.GenCelda("","vt","","","","","",sImagen6SS & "&nbsp;" & sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "CIRCUITO: "
       celdas = ofv.GenCelda("","vT","","","","","",sInput,"si")

       sInput = session("WKFDSC")
       celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    END IF 



    celdas = ofv.GenCelda("","vt colspan=2 ","","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")
    TAB = TAB  & ofv.GenTabla("","aaa style='border-bottom-style:SOLID;border-bottom-width:1;border-bottom-color:#f7aa22;' ","","98%","","0","","0","0","","",filas,"si") & "<BR>"



    if act = "S" then

       filas = ""

       FOR ISHW = 1 TO TSHW

         sinput = sImagenPC  & "&nbsp;PASO  " & ATABSW(ISHW,1) 
         celdas = ofv.GenCelda("","vT ","","","50%","","",sInput,"si")

         IF ATABSW(ISHW,2) = 0 THEN 
            SINPUT3 = "<FONT SIZE=1 COLOR=#008000><U>INICIADO</U></FONT>"
         ELSE
            SINPUT3 = "<FONT SIZE=1 COLOR=#993333><U>CONFORMADO</U></FONT>"
         END IF 

         sinput = sImagenC & "&nbsp;ESTADO:&nbsp;" & SINPUT3 & "<BR>" & NOW()
         celdas = celdas & ofv.GenCelda("","vt","","","50%","","",sInput,"si")

         filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")


         celdas = ofv.GenCelda("","vt COLSPAN=2","","","50%","","","&nbsp;","si")
         filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")

         sinput = "DESTINO:"
         celdas = ofv.GenCelda("","vt","","","50%","","",sInput,"si")

         sinput =  sImagen6SS & "&nbsp;" & ATABSW(ISHW,3)
         celdas = celdas & ofv.GenCelda("","vt","","","50%","","",sInput,"si")

         filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")

         celdas = ofv.GenCelda("","vt COLSPAN=2 style='border-TOP-style:SOLID;border-TOP-width:1;border-TOP-color:#6699CC;' ","","","50%","","","&nbsp;","si")
         filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")

       NEXT


       sinput = "&nbsp;"
       celdas = ofv.GenCelda("","vt","","","50%","","",sInput,"si")

       sinput = "<a class=tt  href=EDREQUERNV.asp?OPCION=RELOAD"
       sinput = sinput  & " ><u><b>VOLVER A SOLICITUD</b></u></a>" 
       celdas = CELDAS & ofv.GenCelda("","tt","CENTER","","50%","","",sInput,"si")

       filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

       TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

    else

       sinput = sImagenPC  & "&nbsp;PASO  " & XPDSC1 
       celdas = ofv.GenCelda("","vT","","","50%","","",sInput,"si")

       sinput = sImagenC & "&nbsp;ESTADO:&nbsp;PENDIENTE DE INICIO"
       celdas = celdas & ofv.GenCelda("","vt","","","50%","","",sInput,"si")

       filas =  ofv.GenRow("","","","","","","",celdas,"si")

       TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


		    Celda = ofv.GenCelda("","vt colspan=3","CENTER","100%","","","","OBSERVACIONES","si")
		    FilaS = Ofv.GenRow("","","","","","","",Celda,"si")

		    nCol = 80

		    wdescri = "SIN OBSERVACIONES" 

		    t3b = "<font class=at><textarea class=dt "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=10 "
		    if bModificar then
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
		      sAccion = "COWKFREQUERnv.ASP?CantRegMover=0&OPCION=" & SOPCION & "&ACT=S"
		      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
		    else
		      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
		    end if


		   TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario0,"si") & "<BR>"


     end if

     Celda = ofv.GenCelda("","","center","","","","",TAB,"si")
     FilaS = ofv.GenRow("","","","","","","",Celda,"si")


     sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")








   END IF



  ELSE

     SHTML = SHTML & "<P>ERROR EN EL INICIO DEL CIRCUITO</P>"

  END IF


  call ofv.cerrarconsulta(rs)





''''''''''''''''''''' FIN OPCIONES DE INICIO DEL CIRCUITO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

 ELSEIF SOPCION = "OK" OR SOPCION = "RCH" THEN

  


'''''''''''''''''''''''''''''' nuevo administrador del workflow version 2 ''''''''''''''''''''''''''''''''''''''''''''''

 ELSEIF SOPCION <> "" THEN

	 IF ACT = "S" THEN
	    sobserv = request.form("OBS")




	    Strsql = "SELECT * FROM CKTVIEWnv WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	    if NOT rsx.eof then

	       psctipo  = RSx(22)
	       pscnivel = RSx(20)

            end if


            StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & SESSION("WKFPSC") & "' "
	    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	    if not rs2.eof then
               XPASACCADPAS    = rs2(25)
            END IF 


 	    StrSql = "Select * From WKFPASOSPRTYNV WHERE WKFCOD = '" & SESSION("WKF") & "' AND "
 	    StrSql = StrSql & " WKFPASCOD = '" & SESSION("WKFPSC") & "' AND ITEMCOD = '" & SOPCION & "'"
  	    set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	    if not rsXS.eof then

      	                     XPASCOD    = RSXS(7)
    	                     XPASACC    = RSXS(9)
    	                     XPASACCADPRTY    = RSXS(11)
      	                     XPASEST    = RSXS(10)
       	                     PSCMOT     = RSXS(4)
       	                     PSCTIPB    = RSXS(5)
        	             PSCTIPD    = RSXS(5)

         	             SOPCDESC = RSXS(4)

            END IF






		    SEGCKTPSC = GTCKTPSC()
		  '''  response.write segcktpsc




			IF PSCTIPO = "U"  THEN

'''''''''''''''''''''''''''''''''''''' por paso con responsable unico (unitario)


                            if XPASACC = "A" then

				    WXX = GENACCIONNV(XPASACC,XPASEST)

                            end if 

                            if XPASACCADPAS <> "D" then

				    WXX = GENACCIONADNV(XPASACCADPAS)

                            end if 

                            if XPASACCADPRTY <> "D" then

				    WXX = GENACCIONADNV(XPASACCADPRTY)

                            end if 




				      StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & XPASCOD & "' "
				      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				      if not rs2.eof then

				             XPASCOD2    = "NO"
				             XMODEL2     = SESSION("WKF")
				             XMODELPSC2  = RS2(2)


				              XSINST2     = 1
				              XSCICLO2    = RS2(22)
				              XSNIVEL2    = RS2(21)
				              XSSUBN2     = 1

				              XMODO2      = "U"
				              XESTILO2    = ATRSPNIV(XSNIVEL2,04)
				              XNODO2      = "NO"



				              XRTIPO2     = ATRSPNIV(XSNIVEL2,07)
				              XRCOD2      = ATRSPNIV(XSNIVEL2,08)
				              XRASG2      = "NO"
				              XRLVL2      = SUPLEVL



				              XPDSC2      = RS2(3) & ": " & RS2(4) 
				              XPDSC12     = RS2(3) & ": " & RS2(4)  
                                              XOBSERV2    = ""


                                              IF  XRTIPO2 = "SOL" THEN
				                  XRTIPO2     = WKNIVSOLTIP
				                  XRCOD2      = WKNIVSOLCOD
				                  XRASG2      = WKNIVSOLCOD

                                              END IF 

                                              IF  XRTIPO2 = "DES" THEN
				                  XRTIPO2     = WKNIVDESTIP
				                  XRCOD2      = WKNIVDESCOD
				                  XRASG2      = WKNUSRDEST

                                              END IF 


                                              IF smodoasg = "SU" THEN

                                                 IF (smodoUSR & "") <> "" THEN

				                    XRASG2      = smodoUSR
                                                 END IF


                                              END IF





 '''''''''''''''''''
 ' agregar envio de mail
                                       XXXX = GENCKTMAIL(XRTIPO2,XRCOD2,XPDSC2 & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL2)
 '''''''''''''''''''



				''' GENERAR EL PROXIMO PASO

    WXX = GENCKTPSCNV(SEGCKTPSC,"NO",XMODEL2,XMODELPSC2,XSINST2,XSCICLO2,XSNIVEL2,XSSUBN2,XMODO2,XESTILO2,XNODO2,"NO",XRTIPO2,XRCOD2,XRASG2,XPDSC2,XOBSERV2,XRLVL2)



				    END IF
				    call ofv.cerrarconsulta(rs2)

			ELSE

'''''''''''''''''''''''''''''''''''''' por paso con responsables segun lista (recurrente)




 '''''''''''''''''''


			END IF

			''''''''''''''''''''''''''''' ACTUALIZA PASO Y CIRCUITO DEL WORKFLOW CORRIENTE ''''''''''''''''''''''''''''''''''

                            GETFECH()

                            WKFFINAL = XLFECDAT - clng(cdate("01-01-1900"))
                            WKHFINAL = XLHORDAT


			    strsql = "UPDATE SEGPASONV SET  SEGPASOSIG = '" & SEGCKTPSC & "', observaciones = '" & sobserv & "', "
 	                    StrSql = StrSql & " SEGFFINAL= '" & WKFFINAL & "',SEGHFINAL= '" & WKHFINAL & "',SEGESTADO= 'SI',"
 	                    StrSql = StrSql & " SEGNODO= '" & SOPCION & "',RESPASG = '" & SESSION("USUARIO") & "'"
 	                    StrSql = StrSql & " WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
  ''  response.write strsql & "<br>" & chr(13)
			    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			    IF SOPCION = "FIN" THEN

				     strsql = "UPDATE SEGCIRCUITO SET FIN = '" & NOW() & "',CKTFFINAL = '" & WKFFINAL & "',CKTHFINAL = '" & WKHFINAL & "' "
 	                             StrSql = StrSql & " WHERE SEGCKT = '" & SESSION("SEGCKT") & "' "
				     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

			    END IF


''''''''''''''''''''''''''''''''''''''''' actualiza el responsable en los responsables del circuito '''''''''''''''''''''''''

			IF PSCTIPO = "U"  THEN

                           wkrsporden = 10

                           wkrspcarga = "S" 

	 	           Strsql = "SELECT * FROM SEGCKTRESPNV WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND "
 	                   StrSql = StrSql & " SEGNIVEL = " & pscnivel & " AND SEGCODIGO = '" & SESSION("USUARIO") & "' "
	 	           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	           if NOT rsx.eof then
                              wkrspcarga = "N" 

                           ELSE

	 	              Strsql = "SELECT * FROM SEGCKTRESPNV WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND "
 	                      StrSql = StrSql & " SEGNIVEL = " & pscnivel & " ORDER BY SEGRSPSECUEN DESC "
	 	              set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	              if NOT rsx.eof then
                                 wkrsporden = rsx(8) + 10
                              END IF

		              SEGCKTRSP = GTCKTPSC()
                              wkrspSECTOR = 0
                              wkrspNIVEL = 8

	                      set cn2 = ofv.conectar(ofv.strconn0)
	                      strsql = "Select SECTOR,NIVEL From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
	                      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	                      if not rsniv.eof then
                                 wkrspSECTOR = rsniv(0)
                                 wkrspNIVEL = rsniv(1)
                              END IF 
	                      call ofv.cerrarconsulta(rsniv)


                              strsql = "insert into SEGCKTRESPNV ("
                              STRSQL = STRSQL & "SEGCKT,SEGNIVEL,SEGRSPCOD,SEGTIPO,"
                              STRSQL = STRSQL & "SEGCODIGO,SEGNIVELCNF,SEGRSPORDEN,SEGRSPSECUEN,"
                              STRSQL = STRSQL & "SEGMODOASG,SEGRAZONASG,SEGCODIGOASG,SEGESTADORSP,"
                              STRSQL = STRSQL & "SEGRSPDSC,SEGRSPSECTOR"
                              STRSQL = STRSQL & ") values ("
                              STRSQL = STRSQL & "'" & SESSION("SEGCKT") & "','" & pscnivel & "','" & SEGCKTRSP & "','USR',"
                              STRSQL = STRSQL & "'" & SESSION("USUARIO") & "','" & wkrspNIVEL & "','10','" & wkrsporden & "',"
                              STRSQL = STRSQL & "'WKF','RP','','OK',"
                              STRSQL = STRSQL & "'RESPONSABLE NIVEL " & pscnivel & "','" & wkrspSECTOR & "'"
                              STRSQL = STRSQL & ")"
                      ''       response.write strsql & "<br>" & chr(13)
                              call ofv.crearconsultaEx(StrSql,cn,1,parametros)








                           END IF 



                        else


                        end if 


''''''''''''''''''' ACTUALIZA EL PROYECTO ASOCIADO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

                        IF WKREQCRPT = "S" THEN

                           XXXX = GENACTPRY()

                        END IF

''''''''''''''''''' FIN ACTUALIZA EL PROYECTO ASOCIADO '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	 	  Strsql = "SELECT * FROM CKTVIEWnv WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	 	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	  if NOT rsx.eof then
	 
	 

			  filas  = ""

			  celdas = ofv.GenCelda("","ut","","15%","","","","CIRCUITO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","ut","","15%","","","","PASO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenp & "</b> " & rsX(31),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			  
			  celdas = ofv.GenCelda("","ut","","15%","","","","OPCION:","si")
			     
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","",SOPCDESC,"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			  
			  IF RSX(37) <> "NO" THEN

			       celdas = ofv.GenCelda("","ut colspan=2 style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ","CENTER","","","","","<B>PASO FINALIZADO:</B>","si")
			       celdas2 = ofv.GenCelda("","ut colspan=2","CENTER","","","","",GTFechadorsh(RSX(35)) & "&nbsp;-&nbsp;" & GETGLBHMSW(RSX(36))  & "&nbsp;","si")

			       filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
			       Filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")


			  END IF
			  
			  

			  sHTML = sHTML  & ofv.GenTabla(""," ","","100%","","0","","-2","-1","","",filas,"si") & "<BR><BR>"
	 	  
	 	  END  IF





         else

'''''''''''''''''''''''''''''' actualizacion del paso - preliminar - modifica observaciones  ''''''''''''''''''''''''''''''''''''''''''''''


                  IF smodoasg = "SU" THEN

                     smodoUSR = request.form("DATUSRDEST") & ""

                     IF (smodoUSR & "") = "" OR (smodoUSR & "") = "0" THEN
                         smodoUSR = SESSION("USUARIO")
                     END IF

                  END IF


	 	  Strsql = "SELECT * FROM CKTVIEWnv WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	 	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	 	  if NOT rsx.eof then
	 
	 

			  filas  = ""

			  celdas = ofv.GenCelda("","ut","","15%","","","","CIRCUITO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

			  celdas = ofv.GenCelda("","ut","","15%","","","","PASO:","si")
			  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenP & "</b> " & rsX(31),"si")
			  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


 	                  StrSql = "Select * From WKFPASOSPRTYNV WHERE WKFCOD = '" & SESSION("WKF") & "' AND "
 	                  StrSql = StrSql & " WKFPASCOD = '" & rsX(17) & "' AND ITEMCOD = '" & SOPCION & "'"
  	                  set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	                  if not rsXS.eof then

      	                     XPASCOD    = RSXS(7)
    	                     XPASACC    = RSXS(9)
      	                     XPASEST    = RSXS(10)
       	                     PSCMOT     = RSXS(4)
       	                     PSCTIPB    = RSXS(5)
        	             PSCTIPD    = RSXS(5)

         	             SOPCDESC = RSXS(4)

          	          END IF

			  
			  celdas = ofv.GenCelda("","ut","","15%","","","","OPCION:","si")
			     
			  celdas = celdas & ofv.GenCelda("","vt","","80%","","","",SOPCDESC,"si")
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
		      sAccion = "COWKFREQUERNV.ASP?CantRegMover=0&OPCION=" & SOPCION & "&ACT=S&MODOASG=" & smodoasg & "&MODOUSR=" & smodoUSR
		      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
		    else
		      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
		    end if


		    sHTML = sHTML  & ofv.GenTabla("","aaa style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;' ","","100%","","0","","-2","-1","","",formulario0,"si") & "<BR><BR>"



         end if


'''''''''''''''''''''''''''''' entra desde la ventana de novedades del workflow  ''''''''''''''''''''''''''''''''''''''''''''''

 ELSE



	  Strsql = "SELECT * FROM CKTVIEWnv WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
	  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsx.eof then

	  filas  = ""

	  celdas = ofv.GenCelda("","ut","","15%","","","","CIRCUITO:","si")
	  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenC & "</b> " & rsX(7),"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","ut","","15%","","","","PASO:","si")
	  celdas = celdas & ofv.GenCelda("","ut","","80%","","","","<b>" & sImagenp & "</b> " & rsX(31),"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  sHTML = sHTML  & ofv.GenTabla(""," ","","98%","","0","","0","0","","",filas,"si") & "<BR>"




	  Strsql = "SELECT * FROM CKTVIEWnv WHERE SEGPASOSIG = '" & session("SEGCKTPSC") & "' "
	  set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxT.eof then

             PSCMODEL = RSXT(17)
             PSCMOT   = "A"
             PSCNODO  = RSXT(30)

	     Strsql = "SELECT * FROM WKFPASOSNV WHERE WKFPASCOD = '" & PSCMODEL & "' "
	     set rsxu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	     if NOT rsxU.eof then



                  StrSql = "Select * From WKFPASOSPRTYNV WHERE WKFCOD = '" & SESSION("WKF") & "' AND "
                  StrSql = StrSql & " WKFPASCOD = '" & PSCMODEL & "' AND ITEMCOD = '" & PSCNODO & "'"
                  set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if not rsXS.eof then

                     XPASCOD    = RSXS(7)
                     XPASACC    = RSXS(9)
                     XPASEST    = RSXS(10)
                     PSCMOT     = RSXS(4)
                     PSCTIPB    = RSXS(5)
                     PSCTIPD    = RSXS(5)

                  END IF




             END IF 



   

	  celdas = ofv.GenCelda("","uT bgcolor='#eeeeee' ","","70%","","","","ANTERIOR <B>" & rsXT(31) & "</B>&nbsp;-&nbsp;POR&nbsp;" & PSCMOT,"si")
	  celdas = celdas & ofv.GenCelda("","uT bgcolor='#eeeeee' ","RIGHT","30%","","","", GTFechadorsh(RSXT(35)) & "&nbsp;-&nbsp;" & GETGLBHMSW(RSXT(36)) & "&nbsp;","si")
	  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

			set cn2 = ofv.conectar(ofv.strconn0)

			sniv  = "DESCONOCIDO"
                        SNSEC = "DESCONOCIDO"

			strsql = "Select DESCRIPCION,SECTOR From USUARIOS where ID = '" & rsXT(26) & "' "
			set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			if not rsniv.eof then
                               sniv = rsniv(0)
                               SNSEC = GETRESPOCKT(rsniv(1),"SEC") 
                        END IF
			call ofv.cerrarconsulta(rsniv)

	  celdas = ofv.GenCelda("","ut colspan=2","","","","","","RESPONSABLE:&nbsp;" & sniv,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
	  celdas = ofv.GenCelda("","ut colspan=2 ","","","","","","AREA/SECTOR:&nbsp;" & SNSEC,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  celdas = ofv.GenCelda("","ut COLSPAN=2 STYLE='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;background-color:#f9f9f9;' ","CENTER","","","","","OBSERVACIONES","si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

		    nCol = 110
		    t3b = "<font class=at><textarea style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#888888;background-color:#ffffff;color:#000080;' "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=4 "
                    t3b = t3b & " readonly >"
		    t3b = t3b & rsXT(32) & "</textarea></font>"


	  celdas = ofv.GenCelda("","ut COLSPAN=2   ","center","","","","",t3b,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	  sHTML = sHTML  & ofv.GenTabla("","aaa style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#888888;' ","","98%","","0","","0","0","","",filas,"si") & "<br>"


          END IF


       session("WKFPSC")    = RSX(17)
       SESSION("SEGCKT")    = RSX(5)



	  IF RSX(37) <> "NO" THEN

	       celdas1 = ofv.GenCelda("","ut style='border-top-style:solid;border-top-width:2;border-top-color:#f7aa22;'  ","CENTER","","","","","<B>PASO FINALIZADO:</B>","si")
	       celdas2 = ofv.GenCelda("","ut","CENTER","","","","","<b>" & sImagenff & "</b> " & GTFechadorsh(RSX(35)),"si")


	       filas  = ofv.GenRow("","","","","10","","",celdas1,"si")
	       filas = filas & ofv.GenRow("","","","","10","","",celdas2,"si")



	       sHTML = sHTML  & ofv.GenTabla(""," ","","100%","","0","","-2","-1","","",filas,"si") & "<br><br>"



	  ELSE


            StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & session("WKFPSC") & "' "
	    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	    if not rs2.eof then



	         XPASCOD2    = "NO"
	         XMODEL2     = SESSION("WKF")
	         XMODELPSC2  = RS2(2)


	         XSINST2     = 1
	         XSCICLO2    = RS2(22)
	         XSNIVEL2    = RS2(21)
	         XSSUBN2     = 1

	         XMODO2      = "U"
	         XESTILO2    = "A"
	         XNODO2      = "NO"



	         XRTIPO2     = WKNIVDESTIP
	         XRCOD2      = WKNIVDESCOD
	         XRASG2      = "NO"
	         XRLVL2      = SUPLEVL

	         XPDSC2      = RS2(3) & ": " & RS2(4) 
	         XPDSC12     = RS2(3) & ": " & RS2(4)  




	       psctipo  = XMODO2
	       pscclase = "M"

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

                  DIM ATABWPSCSAL(15,10)
                  ITWSAL = 0
                  TTWSAL = 0

                  FOR ITWSAL = 1 to 12


                    IF RS2(7 + ITWSAL) = "S" THEN

                       if ITWSAL < 3 then

                          if  ITWSAL = 1 then  WKNOD = "SAL"
                          if  ITWSAL = 2 then  WKNOD = "CAN"

                       else 
                    
                          WKNOD = 9900 + (ITWSAL - 2)
                          WKNOD = REPLACE(CSTR(WKNOD),"99","N")

                       end if


                      StrSql = "Select * From WKFPASOSPRTYNV WHERE WKFCOD = '" & SESSION("WKF") & "' AND "
                      StrSql = StrSql & " WKFPASCOD = '" & session("WKFPSC") & "' AND ITEMCOD = '" & WKNOD & "'"
                      set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rsXS.eof then

                            TTWSAL = TTWSAL + 1

                            ATABWPSCSAL(TTWSAL,00)    = WKNOD     'nodo

                            ATABWPSCSAL(TTWSAL,01)    = RSXS(7)     'PASO SIGUIENTE
                            ATABWPSCSAL(TTWSAL,02)    = RSXS(9)     'ACC CAMBIO ESTADO
                            ATABWPSCSAL(TTWSAL,03)    = RSXS(10)    'NUEVO ESTADO
                            ATABWPSCSAL(TTWSAL,04)    = RSXS(4)     'DESCRIPCION BOTON
                            ATABWPSCSAL(TTWSAL,05)    = RSXS(5)     'TIPO PASO CODIGO
                            ATABWPSCSAL(TTWSAL,06)    = RSXS(5)     'TIPO PASO DECRIPCION

                            ATABWPSCSAL(TTWSAL,10)    = RSXS(11)    'DEAFULT - NO REQUIERE ACCION ADICIONAL


                            IF (RSXS(11) & "") = "SU" THEN
                               ATABWPSCSAL(TTWSAL,10)    = "SU"         'REQUIERE SOLICITAR EL USUARIO A DONDE VA

                            END IF

                            StrSql = "Select * From WKFPASOSNV WHERE WKFPASCOD = '" & RSXS(7) & "' "
                            set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if not rs3.eof then

                               ATABWPSCSAL(TTWSAL,07)    = RS3(3) & ": " & RS3(4)      'DESCRIPCION PROXIMO PASO
    
                               XSNIVEL2    = RS3(21)

                               ATABWPSCSAL(TTWSAL,08)    =  ATRSPNIV(XSNIVEL2,02)  'DESCRIPCION PROXIMO NIVEL

                            END IF

                            StrSql = "Select COLOR From WKFPSCTIPOS WHERE TIPO = '" & RSXS(5) & "'"
                            set rsXS2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                            if not rsXS2.eof then
                               ATABWPSCSAL(TTWSAL,09)    = rsXS2(0)      'COLOR
                            ELSE
                               ATABWPSCSAL(TTWSAL,09)    = "#6699CC"     'COLOR
                            END IF



                      END IF

                     END IF


                  NEXT


		  celdasB = ""


		  IF PSCTIPO = "U" THEN

''''''''''''''''''''' PASO CON RESPONSABLE UNITARIO '''''''''''''''''''''''''''''''''''''''''''''''''''''''

                      NsHTML = ""

                      FOR ITWSAL = 1 TO TTWSAL

                          sFormDSP = "frm" & ATABWPSCSAL(iTWSAL,00)

	                  filasB  = ""

          		  celdas  = ofv.GenCelda("","utI style='background-color:#888888;COLOR:#FFFFFF;' ","","10%","","","","SIGUE A:&nbsp;","si")
	           	  celdas  = celdas & ofv.GenCelda("","ut  ","","90%","","","",ATABWPSCSAL(ITWSAL,07),"si")
	                  filasB  = filasB & ofv.GenRow("","","","","20","","",celdas,"si")

          		  celdas  = ofv.GenCelda("","utI style='background-color:#888888;COLOR:#FFFFFF;' ","","","","","","DESTINO:&nbsp;","si")

                          IF ATABWPSCSAL(ITWSAL,10)    = "SU" THEN

                             NSELECT = ""
                             NSELECT = NSELECT & "<SELECT CLASS=AT NAME='DATUSRDEST' > " 
                             NSELECT = NSELECT & "<OPTION VALUE='0' >ASIGNAR USUARIO DESTINATARIO</OPTION>" 
                             set cn2 = ofv.conectar(ofv.strconn0)
                             scomK = scomK & "xxxx" & chr(9)
                             scomD = scomD & "ASIGNAR USUARIO DESTINATARIO" & chr(9)

                             strsql = "Select * From USUARIOS    "
                             strsql = strsql & " WHERE ID IN ("
                             wkfsep3 = "" 
                             strsql2 = "Select * From GPREQRESPONSABLES WHERE TIPOUSR = 'U' AND USUARIOSECTOR <> '" & SESSION("USUARIO") & "' "
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
                                NSELECT = NSELECT & "<OPTION VALUE='"  & rsniv(0) & "' >"  & rsniv(1) & "</OPTION>" 
                                scomK = scomK & rsniv(0) & chr(9)
                                scomD = scomD & rsniv(1) & chr(9)
                               RSNIV.MOVENEXT
                             LOOP
                             call ofv.cerrarconsulta(rsniv)

                             NSELECT = NSELECT & "</SELECT> " 


                      	  ''''   sInput2 = ofv.gencomboFX(sFormDSP,"DATUSRDEST","xxxx",scomK,scomD,"no")
                      	     sInput2 = NSELECT

                             NsHTML = ""
                             NsHTML = NsHTML & "  if (fr.name=='" & sFormDSP & "')"
                             NsHTML = NsHTML & "    {"
                             NsHTML = NsHTML & "      if (document.all.DATUSRDEST.value!='0')"
                             NsHTML = NsHTML & "        {"
                             NsHTML = NsHTML & "        }"
                             NsHTML = NsHTML & "      else"
                             NsHTML = NsHTML & "        {"
                             NsHTML = NsHTML & "         est=1;"
                             NsHTML = NsHTML & "         emsg= emsg + '- USUARIO DESTINATARIO\n';"
                             NsHTML = NsHTML & "        }"
                             NsHTML = NsHTML & "    }   "

                             call ofv.cerrarconn(cn2)

                             celdas  = celdas & ofv.GenCelda("","ut  ","","","","","",sInput2,"si")
                          ELSE 

                             celdas  = celdas & ofv.GenCelda("","ut  ","","","","","",ATABWPSCSAL(ITWSAL,08),"si")
                          END IF


	                  filasB  = filasB & ofv.GenRow("","","","","20","","",celdas,"si")

                          sfuncion = " onclick='document." & sFormDSP & ".submit();' "

                       ''   sfuncion = " onclick=' alert(document." & sFormDSP & ".name);' "

		          sinput = "" & chr(13)
		          sinput = sinput  & "<input type='button' size=10 name='bot" & ATABWPSCSAL(iTWSAL,00) & "' onclick=""disp_confirm('CONFIRMA " & ATABWPSCSAL(iTWSAL,04) & "???','" & sFormDSP & "');"" "
		          sinput = sinput  & " value='" & ATABWPSCSAL(iTWSAL,04) & "' style='cursor:hand;background-color:" & ATABWPSCSAL(iTWSAL,09) & ";color:#ffffff;' >"

          		  celdas  = ofv.GenCelda("","ut COLSPAN=2 style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#dddddd;' ","CENTER","","","","",sinput,"si")

	                  filasB  = filasB & ofv.GenRow("","","","","25","","",celdas,"si")


		          sinput = "" & chr(13)
		          sinput = sinput  & "<form name='" & SFORMDSP & "' action='cowkfrequerNV.asp?OPCION=" & ATABWPSCSAL(iTWSAL,00) & "&MODOASG=" & ATABWPSCSAL(ITWSAL,10) & "' "
		          sinput = sinput  & " method=post  >" & filasB
		          sinput = sinput  & "</form>" & chr(13)


	                  WCELTAB = ofv.GenTabla("","aaa style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#dddddd;background-color:#f0f0f0;' ","","250","","0","","2","-1","","",sinput,"si")

          		  celdasB  = celdasB & ofv.GenCelda("","ut style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#990000;' ","CENTER","","","","",WCELTAB,"si")




                      NEXT
                         
                      sHTML = sHTML & "<script type='text/javascript'>"
                      sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
                      sHTML = sHTML & "  {"
                      sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
                      sHTML = sHTML & "  var est=0;"
                      sHTML = sHTML & "  var emsg='DEBE INFORMAR:\n\n';"

                      sHTML = sHTML & NsHTML

                      sHTML = sHTML & "  if (est==1)"
                      sHTML = sHTML & "   {"
                      sHTML = sHTML & "    alert(emsg);"
                      sHTML = sHTML & "   }"
                      sHTML = sHTML & "  else"
                      sHTML = sHTML & "   {"

                      sHTML = sHTML & "     var r=confirm(disp_msg);"
                      sHTML = sHTML & "     if (r==true)"
                      sHTML = sHTML & "       {"
                    '  sHTML = sHTML & "        alert(fr.name);"
                      sHTML = sHTML & "         fr.submit();"
                      sHTML = sHTML & "       }"
                      sHTML = sHTML & "     else"
                      sHTML = sHTML & "       {"
                      sHTML = sHTML & "         alert('Se cancelo la accion!');"
                      sHTML = sHTML & "       }"
                      sHTML = sHTML & "     }"
                      sHTML = sHTML & "   }"
                      sHTML = sHTML & "</script>"



		  ELSE

''''''''''''''''''''' PASO CON RESPONSABLES RECURRENTE '''''''''''''''''''''''''''''''''''''''''''''''''''''''


		  END IF


	       filas  = ofv.GenRow("","","","","10","","",celdasB,"si")

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

FUNCTION GENCKTINI()

    SEGCKT = GTCKT()

    SESSION("SEGCKT") = SEGCKT

    GETFECH()

    WKFINICIO = XLFECDAT - clng(cdate("01-01-1900"))
    WKHINICIO = XLHORDAT

    strsql = "insert into SEGCIRCUITO ("
    STRSQL = STRSQL & "MODEL,COMIENZO,FIN,OBJTIPO,"
    STRSQL = STRSQL & "OBJCOD,SEGCKT,RESPONSABLE,CKTDSC,"
    STRSQL = STRSQL & "CKTMODAL,CKTVERSION,CKTFINICIO,CKTHINICIO,"
    STRSQL = STRSQL & "CKTFFINAL,CKTHFINAL"
    STRSQL = STRSQL & ") values ("
    STRSQL = STRSQL & "'" & SESSION("WKF") & "','" & NOW() & "','NO','REQ',"
    STRSQL = STRSQL & "'" & session("NCliente") & "','" & SEGCKT & "','" & SESSION("USUARIO") & "','" & session("WKFDSC") & "',"
    STRSQL = STRSQL & "'S','2','" & WKFINICIO & "','" & WKHINICIO & "',"
    STRSQL = STRSQL & "'0','0'"
    STRSQL = STRSQL & ")"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)


    FOR ITRSP=1 TO TTRSP



    SEGCKTNIV = GTCKT() & "N" & ITRSP

    wksupnivel = SUPLEVL

    if ATRSPNIV(ITRSP,05) <> "D" then wksupnivel = ATRSPNIV(ITRSP,05)


    strsql = "insert into SEGNIVELESNV ("
    STRSQL = STRSQL & "SEGCKT,SEGNIVCOD,MODELWKF,SEGNIVEL,"
    STRSQL = STRSQL & "SEGMODO,SEGESTILO,SEGTIPO,SEGCODIGO,"
    STRSQL = STRSQL & "SEGNIVELCNF,SEGNIVDSC"
    STRSQL = STRSQL & ") values ("
    STRSQL = STRSQL & "'" & SESSION("SEGCKT") & "','" & SEGCKTNIV & "','" & SESSION("WKF") & "','" & ATRSPNIV(ITRSP,01) & "',"
    STRSQL = STRSQL & "'" & ATRSPNIV(ITRSP,03) & "','" & ATRSPNIV(ITRSP,04) & "','" & ATRSPNIV(ITRSP,07) & "','" & ATRSPNIV(ITRSP,08) & "',"
    STRSQL = STRSQL & "'" & wksupnivel & "','" & ATRSPNIV(ITRSP,02) & "'"
    STRSQL = STRSQL & ")"
  ''  response.write strsql & "<br>" & chr(13)
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    NEXT







END FUNCTION



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


FUNCTION GENACCIONNV(WXACC,WXEST)

   
    strsql = "UPDATE REQUERIMIENTO SET ESTADO = " & WXEST & " WHERE codigo = '" & session("NCliente") & "' "
    call ofv.crearconsultaEx(StrSql,cn4,1,parametros)

 



END FUNCTION


FUNCTION GENACCIONADNV(WXACCAD)

       WXEXEACC = "N"
       WXHOY    = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))


       select case CSTR(WXACCAD & "")

           case "D"

              ''   sInput = "NO"
              WXEXEACC = "N"
              WXEXEACCRES = "OK"

           case "AR"

              ''   sInput = "ASIGNAR RESPONSABLE"
              WXEXEACC = "S"

              WXEXEACCDAT = " ASIGNADO = '" & SESSION("USUARIO") & "' "
              WXEXEACCRES = "OK"               

           case "IR"

              ''   sInput = "INICIAR REQUERIMIENTO"
              WXEXEACC = "S"

              WXEXEACCDAT = " FECHASOL = '" & WXHOY & "' "
              WXEXEACCRES = "OK"


           case "FR"

             ''    sInput = "FINALIZAR REQUERIMIENTO"
              WXEXEACC = "S"

              WXEXEACCDAT = " FECHAFIN = '" & WXHOY & "' "
              WXEXEACCRES = "OK"

           case "SU"

              ''   sInput = "SOLICITAR USUARIO"
              WXEXEACC = "N"
              WXEXEACCRES = "SU"

           case "EEP"

              ''   sInput = "ENVIAR AL EMISOR DEL PASO"
              WXEXEACC = "N"
              WXEXEACCRES = "EEP"

           case else 
           
              ''   sInput = "sin definir"
              WXEXEACC = "N"
              WXEXEACCRES = "OK"

       end select 

    IF  WXEXEACC = "S" THEN
   
        strsql = "UPDATE REQUERIMIENTO SET " & WXEXEACCDAT & " WHERE codigo = '" & session("NCliente") & "' "
        call ofv.crearconsultaEx(StrSql,cn4,1,parametros)

        StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
        set rsXX = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
        if not rsXX.eof then
           WKNUSRORIG  = RSXX(13)
           WKNUSRDEST  = RSXX(14)
        END IF 

        if CSTR(WXACCAD & "") = "AR" then

           strsql = "UPDATE PROYECTO SET " & WXEXEACCDAT & " WHERE codigo = '" & session("NCliente") & "' "
           call ofv.crearconsultaEx(StrSql,cn4,1,parametros)

           XXXX = GENPROYRESP(SESSION("USUARIO"),"AN",XSNIVEL2)


        end if



    END IF

    GENACCIONADNV = WXEXEACCRES



END FUNCTION


FUNCTION GENCKTPSCNV(XPSC,XPSIG,XMODEL,XMODELPSC,XSINST,XSCICLO,XSNIVEL,XSSUBN,XMODO,XESTILO,XNODO,XFIN,XRTIPO,XRRSP,XRASG,XPDSC,XOBSERV,XNIVEL)

    SEGCKTPSC = XPSC

    GETFECH()

    WKFINICIO = XLFECDAT - clng(cdate("01-01-1900"))
    WKHINICIO = XLHORDAT


    IF XFIN = "SI" THEN
       GETFECH()

       WKFFINAL = XLFECDAT - clng(cdate("01-01-1900"))
       WKHFINAL = XLHORDAT
       XFIN = "SI"
    ELSE
       WKFFINAL = 0
       WKHFINAL = 0
       XFIN = "NO"
    END IF

    strsql = "insert into SEGPASONV ("
    STRSQL = STRSQL & "SEGCKT,SEGCKTPSC,MODELWKF,MODELPSCWKF,"
    STRSQL = STRSQL & "SEGINSTANCIA,SEGCICLO,SEGNIVEL,SEGSUBNIVEL,"
    STRSQL = STRSQL & "SEGMODO,SEGESTILO,RESPTIPO,RESPONSABLE,"
    STRSQL = STRSQL & "RESPASG,SEGNIVELCNF,SEGESCALA,SEGPASOSIG,"
    STRSQL = STRSQL & "SEGNODO,PSCDSC,OBSERVACIONES,SEGFINICIO,"
    STRSQL = STRSQL & "SEGHINICIO,SEGFFINAL,SEGHFINAL,SEGESTADO"
    STRSQL = STRSQL & ") values ("
    STRSQL = STRSQL & "'" & SESSION("SEGCKT") & "','" & SEGCKTPSC & "','" & XMODEL & "','" & XMODELPSC & "',"
    STRSQL = STRSQL & "'" & XSINST & "','" & XSCICLO & "','" & XSNIVEL & "','" & XSSUBN & "',"
    STRSQL = STRSQL & "'" & XMODO & "','" & XESTILO & "','" & XRTIPO & "','" & XRRSP & "',"
    STRSQL = STRSQL & "'" & XRASG & "','" & XNIVEL & "','" & XESCALA & "','" & XPSIG & "',"
    STRSQL = STRSQL & "'" & XNODO & "','" & XPDSC & "','" & XOBSERV & "','" & WKFINICIO & "',"
    STRSQL = STRSQL & "'" & WKHINICIO & "','" & WKFFINAL & "','" & WKHFINAL & "','" & XFIN & "'"
    STRSQL = STRSQL & ")"
''    response.write strsql & "<br>" & chr(13)
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)



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


FUNCTION GENACTPRY()




''''''''''''''''''''''''''''''''''''''''' actualiza el paso en las tareas del proyecto asociado     '''''''''''''''''''''''''

                           IF WKREQPRY = "S" THEN

                              	    StrSql = "Select ETAPA,VPLAN From PROYETAPA where codigo = '" & session("NCliente") & "' "
                              	    StrSql = StrSql & "  AND ORDEN = 10 ORDER BY VPLAN DESC "
                              	    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                              	    if NOT rsx.eof then

                                       WKETAPA = RSX(0) 
                                       SESSION("VPLAN") = RSX(1)

                                    END IF

                              	    STRSQL = "SELECT * FROM SEGPASONV "
                                    STRSQL = STRSQL & " WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
                              	    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                              	    if NOT rsx.eof then

                                          IF RSx(7) = WKNIVSOL THEN

                                             WKTIPRSP = "OR"

                                          ELSEIF RSx(7) = WKNIVDES THEN

                                             WKTIPRSP = "AN"

                                          ELSE

                                             WKTIPRSP = "AF"

                                          END IF 


                                         XXXX = GENPROYRESP(RSx(13),WKTIPRSP,RSx(7))

                                         XORDEN = 0


                              	       STRSQL = "SELECT MAX(ORDEN) FROM PROYTAREAS WHERE codigo = '" & session("NCliente") & "' AND "
                                       STRSQL = STRSQL & " VPLAN = " & SESSION("VPLAN") & " AND ETAPA = '" & WKETAPA & "' "
                              	       set rsx2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                              	       if NOT rsx2.eof then

                                          XORDEN = RSX2(0)

                                       END IF



                                       XORDEN = XORDEN + 10


                                       SESSION("PRYCTA") = 1

                                       WKFSTART = RSX(20)
                                       WKHSTART = RSX(21)
                                       WKFEND   = RSX(22)
                                       WKHEND   = RSX(23)

                                       WKUSRASG = RSX(13)
                                       WKPSCDSC = RSX(18)
                                       WKPSCOBS = RSX(19)

                                       WKHSTARTw = GETGLBHMSW(WKHSTART)
                                       WKHENDw   = GETGLBHMSW(WKHEND)


                                       WKPSCTMP = rsX(22) - RSX(20)

                                       IF WKPSCTMP = 0 THEN 

                                     '''''' EN HORAS O MINUTOS

                                          WKPSCTMP = GETGLBHMSW(rsX(23) - RSX(21))

                                          WKPSCTMPH = CINT(LEFT(WKPSCTMP,2))
                                          WKPSCTMPM = CINT(MID(WKPSCTMP,4,2))
                                          WKPSCTMPS = CINT(RIGHT(WKPSCTMP,2))

                                          IF WKPSCTMPH > 0 THEN

                                             IF WKPSCTMPM > 19 THEN
                                                TADV   = WKPSCTMPH + 1
                                             ELSE
                                                TADV   = WKPSCTMPH
                                             END IF

                                             TTIPO  = 1

                                          ELSE
                                             IF WKPSCTMPS > 19 THEN
                                                TADV   = WKPSCTMPM + 1
                                             ELSE
                                                TADV   = WKPSCTMPM
                                             END IF

                                             TTIPO  = 4

                                          END IF


                                       ELSE

                                     '''''' EN DIAS

                                          TADV   = WKPSCTMP
                                          TTIPO  = 2
   
                                       END IF 


                                       Xtsk = OWKF.GENPRYTSK(session("NCliente"),SESSION("VPLAN"),XORDEN,WKETAPA,WKPSCDSC,WKFSTART,WKFEND,TADV,TTIPO,"0","1")

                                       StrSql = "UPDATE PROYTAREAS SET ESTADO = 2,MODCREATE = 2,OBSERVACIONES = '" & WKPSCOBS & "',"
                                       StrSql = StrSql & " MASIG = 1, CASIG = TDUR, TASIG = UDUR,"
                                       StrSql = StrSql & " HORAIRESOL = '" & WKHSTARTw & "', HORAFRESOL = '" & WKHENDw & "',"
                                       StrSql = StrSql & " FECHASREAL = " & WKFSTART & ",FECHAEREAL = " & WKFEND & ", USRRESP = '" & WKUSRASG & "' "
                                       StrSql = StrSql & " WHERE  codigo = '" & session("NCliente") & "' AND TAREA = '" & Xtsk & "' "
                                       Call ofv.crearconsultaEx(StrSql,cn,1,parametros)



                                       Xtsk2 = OWKF.GENPRYTSK(session("NCliente"),"0",XORDEN,WKETAPA,WKPSCDSC,WKFSTART,WKFEND,TADV,TTIPO,"0","1")

                                       StrSql = "UPDATE PROYTAREAS SET ESTADO = 2,MODCREATE = 2,OBSERVACIONES = '" & WKPSCOBS & "',"
                                       StrSql = StrSql & " MASIG = 1, CASIG = TDUR, TASIG = UDUR,"
                                       StrSql = StrSql & " HORAIRESOL = '" & WKHSTARTw & "', HORAFRESOL = '" & WKHENDw & "',"
                                       StrSql = StrSql & "FECHASREAL = " & WKFSTART & ",FECHAEREAL = " & WKFEND & ", USRRESP = '" & WKUSRASG & "', "
                                       StrSql = StrSql & " TAREA = '" & Xtsk & "' "
                                       StrSql = StrSql & " WHERE  codigo = '" & session("NCliente") & "' AND TAREA = '" & Xtsk2 & "' "
                                       Call ofv.crearconsultaEx(StrSql,cn,1,parametros)



'''''''' agregar afectacion a cuentas ''''''''''''''''''''''''

                                     XFDCUENTA = OWKF.GENTSKCUENTA(WKFSTART,TADV, TTIPO, SESSION("PRYCTA"))


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                                      STRSQL = "INSERT INTO PROYTAREASADV (CODIGO,TSK,FUPD,TADV,TPRC,USRUPD,TTIPO,FCARGA,CODINPUT) VALUES ("  
                                      STRSQL = STRSQL & "'" & session("NCliente") & "','" & Xtsk & "'," & WKFEND & ","
                                      STRSQL = STRSQL & "" & TADV & ",'100','" & WKUSRASG & "','" & TTIPO & "','" & WKFEND & "','" & XFDCUENTA & "')" 
                                      CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)






                                   end if









                           else

                                    XXXX = GENPROYASOC()

                              	    StrSql = "Select ETAPA,VPLAN From PROYETAPA where codigo = '" & session("NCliente") & "' "
                              	    StrSql = StrSql & "  AND ORDEN = 10 ORDER BY VPLAN DESC "
                              	    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                              	    if NOT rsx.eof then

                                       WKETAPA = RSX(0) 
                                       SESSION("VPLAN") = RSX(1)


                                    END IF

                              	    STRSQL = "SELECT DISTINCT RESPASG,SEGNIVEL,SEGSUBNIVEL FROM CKTVIEWNV WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND "
                                    STRSQL = STRSQL & " SEGESTADO = 'SI' ORDER BY SEGNIVEL,SEGSUBNIVEL "
                              	    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                              	    if NOT rsx.eof then
                                       DO WHILE NOT RSX.EOF

                                          IF RSx(1) = WKNIVSOL THEN

                                             WKTIPRSP = "OR"

                                          ELSEIF RSx(1) = WKNIVDES THEN

                                             WKTIPRSP = "AN"

                                          ELSE

                                             WKTIPRSP = "AF"

                                          END IF 


                                         XXXX = GENPROYRESP(RSx(0),WKTIPRSP,RSx(1))

                                       RSX.MOVENEXT
                                       LOOP

                                       XORDEN = 0

                              	    STRSQL = "SELECT * FROM SEGPASONV WHERE SEGCKT = '" & SESSION("SEGCKT") & "' AND "
                                    STRSQL = STRSQL & " SEGESTADO = 'SI' "
                                    STRSQL = STRSQL & " ORDER BY SEGFINICIO,SEGHINICIO,SEGCKTPSC "
                              	    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                                    DO WHILE NOT RSX.EOF

                                       XORDEN = XORDEN + 10


                                       SESSION("PRYCTA") = 1

                                       WKFSTART = RSX(20)
                                       WKHSTART = RSX(21)
                                       WKFEND   = RSX(22)
                                       WKHEND   = RSX(23)

                                       WKUSRASG = RSX(13)
                                       WKPSCDSC = RSX(18)
                                       WKPSCOBS = RSX(19)

                                       WKHSTARTw = GETGLBHMSW(WKHSTART)
                                       WKHENDw   = GETGLBHMSW(WKHEND)

                                       WKPSCTMP = rsX(22) - RSX(20)

                                       IF WKPSCTMP = 0 THEN 

                                     '''''' EN HORAS O MINUTOS

                                          WKPSCTMP = GETGLBHMSW(rsX(23) - RSX(21))

                                          WKPSCTMPH = CINT(LEFT(WKPSCTMP,2))
                                          WKPSCTMPM = CINT(MID(WKPSCTMP,4,2))
                                          WKPSCTMPS = CINT(RIGHT(WKPSCTMP,2))

                                          IF WKPSCTMPH > 0 THEN

                                             IF WKPSCTMPM > 19 THEN
                                                TADV   = WKPSCTMPH + 1
                                             ELSE
                                                TADV   = WKPSCTMPH
                                             END IF

                                             TTIPO  = 1

                                          ELSE
                                             IF WKPSCTMPS > 19 THEN
                                                TADV   = WKPSCTMPM + 1
                                             ELSE
                                                TADV   = WKPSCTMPM
                                             END IF

                                             TTIPO  = 4

                                          END IF


                                       ELSE

                                     '''''' EN DIAS

                                          TADV   = WKPSCTMP
                                          TTIPO  = 2
   
                                       END IF 




                                       Xtsk = OWKF.GENPRYTSK(session("NCliente"),SESSION("VPLAN"),XORDEN,WKETAPA,WKPSCDSC,WKFSTART,WKFEND,TADV,TTIPO,"0","1")

                                       StrSql = "UPDATE PROYTAREAS SET ESTADO = 2,MODCREATE = 2,OBSERVACIONES = '" & WKPSCOBS & "',"
                                       StrSql = StrSql & " MASIG = 1, CASIG = TDUR, TASIG = UDUR,"
                                       StrSql = StrSql & " HORAIRESOL = '" & WKHSTARTw & "', HORAFRESOL = '" & WKHENDw & "',"
                                       StrSql = StrSql & " FECHASREAL = " & WKFSTART & ",FECHAEREAL = " & WKFEND & ", USRRESP = '" & WKUSRASG & "' "
                                       StrSql = StrSql & " WHERE  codigo = '" & session("NCliente") & "' AND TAREA = '" & Xtsk & "' "
                                       Call ofv.crearconsultaEx(StrSql,cn,1,parametros)



                                       Xtsk2 = OWKF.GENPRYTSK(session("NCliente"),"0",XORDEN,WKETAPA,WKPSCDSC,WKFSTART,WKFEND,TADV,TTIPO,"0","1")

                                       StrSql = "UPDATE PROYTAREAS SET ESTADO = 2,MODCREATE = 2,OBSERVACIONES = '" & WKPSCOBS & "',"
                                       StrSql = StrSql & " MASIG = 1, CASIG = TDUR, TASIG = UDUR,"
                                       StrSql = StrSql & " HORAIRESOL = '" & WKHSTARTw & "', HORAFRESOL = '" & WKHENDw & "',"
                                       StrSql = StrSql & " FECHASREAL = " & WKFSTART & ",FECHAEREAL = " & WKFEND & ", USRRESP = '" & WKUSRASG & "', "
                                       StrSql = StrSql & " TAREA = '" & Xtsk & "' "
                                       StrSql = StrSql & " WHERE  codigo = '" & session("NCliente") & "' AND TAREA = '" & Xtsk2 & "' "
                                       Call ofv.crearconsultaEx(StrSql,cn,1,parametros)



'''''''' agregar afectacion a cuentas ''''''''''''''''''''''''

                                     XFDCUENTA = OWKF.GENTSKCUENTA(WKFSTART,TADV, TTIPO, SESSION("PRYCTA"))


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


                                      STRSQL = "INSERT INTO PROYTAREASADV (CODIGO,TSK,FUPD,TADV,TPRC,USRUPD,TTIPO,FCARGA,CODINPUT) VALUES ("  
                                      STRSQL = STRSQL & "'" & session("NCliente") & "','" & Xtsk & "'," & WKFEND & ","
                                      STRSQL = STRSQL & "" & TADV & ",'100','" & WKUSRASG & "','" & TTIPO & "','" & WKFEND & "','" & XFDCUENTA & "')" 
                                      CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)



                                    RSX.MOVENEXT
                                    LOOP

                                   end if







                           END IF




 END FUNCTION




FUNCTION GENPROYASOC()

       prycod = OWKF.GENPROY(session("NCliente"))

     '  response.write prycod & "<BR>"


       Session("VPLAN")    = 0 

       OWKF.GENPROP()

    '   response.write prycod & "-2<BR>"

       STRSQL = "UPDATE PROYPRTY SET CONF = 1 WHERE CODIGO = '" & session("NCliente") & "'  "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       OWKF.GENINITPRY()

    ''   YXETP = OWKF.GENPRYETP(Session("NCliente"), Session("VPLAN"), "10", "INCIDENCIAS REALIZADAS", "0", "0")

     '  response.write prycod & "-3<BR>"

    '   StrSql = "UPDATE PROYECTO SET TITULO = '" & PRYTIT & "',CUENTA = 2, MODAL = 0  WHERE CODIGO = '" & prycod & "'"
    '   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

       wdescri = "CIRCUITO DE REQUERIMIENTO"

       XXXX = OWKF.PRYLNCH(prycod,wdescri,"S")

       StrSql = "DELETE From PROYTAREAS WHERE CODIGO = '" & Session("NCliente") & "' "
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

''       XXXX =  GENSECTORORAS() 
''       XXXX =  GENUSRCONTAS()


     '  response.write prycod & "-4<BR>"






END FUNCTION


'''''''''''''''''''''''''
FUNCTION GENPROYRESP(XUSRRSP,XTIPRSP,XNIVRSP)


  StrSql = "Select MAX(ORDEN) From PROYRESP where codigo = '" & session("NCliente") & "'  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     XORDEN = 10

     XXXX = GENSECTORACC("999999","GL")

  ELSE

     XORDEN = RS(0)

     IF XORDEN <> "" THEN
        XORDEN = XORDEN + 10
     ELSE
        XORDEN = 30
     END IF

  END IF


  StrSql = "Select TIPO From PROYRESP where codigo = '" & session("NCliente") & "' and USRAFC = '" & XUSRRSP & "'  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

       set cn2 = ofv.conectar(ofv.strconn0)
       set cn3 = ofv.conectar(ofv.strconn2)



     StrSql = "Select * From USUARIOS where ID = '" & XUSRRSP & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG = rsX2(21)
        XSECTORIGB = XSECTORIG
        XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTORIGB
          set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             EXIT DO
          Else
              XSECTGEREN = rsX3(0)
              XSECTORIGB = rsX3(0)
              StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTORIGB
              set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
              If rsX4.EOF Then
                 EXIT DO
              Else
                 XJERAR = RSX4(3)
                 XJERAR = INSTR(XJERAR,"GEREN")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
              End If
          End If
        
        LOOP 

        STRSQL = "INSERT INTO PROYRESP ("
        STRSQL = STRSQL & "CODIGO,ORDEN,SECTOR,GERENCIA,"
        STRSQL = STRSQL & "TIPO,SUBTIPO,USRAFC,SKR,"
        STRSQL = STRSQL & "CLASE,SK1,SK2,SK3,"
        STRSQL = STRSQL & "SK4,SK5,FUNCION"
        STRSQL = STRSQL & ") VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & XORDEN & "','" & XSECTORIG & "','" & XSECTGEREN & "',"
        STRSQL = STRSQL & "'" & XTIPRSP & "','" & XNIVRSP & "','" & XUSRRSP & "','CT',"

     IF XTIPRSP = "OR" THEN

        STRSQL = STRSQL & "'U','1','0','1',"
        STRSQL = STRSQL & "'1','0','3'"

     ELSEIF XTIPRSP = "AN" THEN

        STRSQL = STRSQL & "'U','1','1','1',"
        STRSQL = STRSQL & "'1','0','2'"

     ELSE

        STRSQL = STRSQL & "'U','1','0','1',"
        STRSQL = STRSQL & "'1','0','1'"

     END IF  

        STRSQL = STRSQL & ")" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

        XXXX = GENSECTORACC(XSECTORIG,"SA")
        XXXX = GENSECTORACC(XSECTGEREN,"RA")

     IF XTIPRSP = "AN" THEN

        XXXX = GENUSRCONTAS(XSECTORIG,XSECTGEREN)

     END IF

  END IF

 END IF






  



END FUNCTION





'''''''''''''''''''''''''
FUNCTION GENUSRCONTAS(XSECTANA,XSECTGEREN)





''''''''''''''''''''''''''''''''''''
'' agregar usuarios de control '''''
''''''''''''''''''''''''''''''''''''

        StrSql = "Select * From PROYRESP where codigo = '0' and SUBTIPO IN ('" & XSECTANA & "','" & XSECTGEREN & "','999999')"
        set rsXp = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rsXp.eof then

           DO WHILE NOT RSXP.EOF


              XP1 = 0
              XP2 = 0
              XP3 = 0
              XP4 = 0
              XP5 = 0

              IF RSXP(10) THEN XP1 = 1
              IF RSXP(11) THEN XP2 = 1
              IF RSXP(12) THEN XP3 = 1
              IF RSXP(13) THEN XP4 = 1
              IF RSXP(14) THEN XP5 = 1

              STRSQL = "INSERT INTO PROYRESP (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC,SKR,CLASE,SK1,SK2,SK3,SK4,SK5,FUNCION) VALUES ("  
              STRSQL = STRSQL & "'" & session("NCliente") & "','100','" & RSXP(3) & "','" & RSXP(4) & "','UC','','" & RSXP(7) & "','" & RSXP(8) & "','" & RSXP(9) & "','" & XP1 & "','" & XP2 & "','" & XP3 & "','" & XP4 & "','" & XP5 & "','" & RSXP(15) & "')" 
              CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

           RSXP.MOVENEXT
           LOOP


        END IF  




END FUNCTION





%>