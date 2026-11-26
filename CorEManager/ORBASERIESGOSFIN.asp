<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sMODAL       = request.querystring("MODAL")
  sMODALMAIN   = request.querystring("MODALMAIN")
  sCODTRACER   = request.querystring("CODTRACER")
  sFECHA       = request.querystring("FECHA")
  sOPCION      = request.querystring("OPCION")
  sOPCSET      = request.querystring("OPCSET")


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  dim taberror (100,5)
  ttaberror = 0

  dim tabDATOS (100,1)
  ttabDATOS = 0

  tabDATOS (000,1) = "CODIGO"
  tabDATOS (001,1) = "CODIGO"
  tabDATOS (002,1) = "F.REGISTRO"
  tabDATOS (003,1) = "F.INICIO"
  tabDATOS (004,1) = "F.DESCUBRIMIENTO"
  tabDATOS (005,1) = "F.R.CONTABLE"
  tabDATOS (006,1) = "PROCESO"
  tabDATOS (007,1) = "SUBPROCESO"
  tabDATOS (008,1) = "LINEA DE NEGOCIOS"
  tabDATOS (009,1) = "ACTIVIDAD"
  tabDATOS (010,1) = "RIESGO"
  tabDATOS (011,1) = "DESCRIPCION"
  tabDATOS (012,1) = "MONTO BRUTO"
  tabDATOS (013,1) = "CODIGO"
  tabDATOS (014,1) = "OTROS MONTOS"
  tabDATOS (015,1) = "MONTO RECUPERADO"
  tabDATOS (016,1) = "MONTO NETO"
  tabDATOS (017,1) = "ESTADO"
  tabDATOS (018,1) = "CODIGO"
  tabDATOS (019,1) = "CODIGO"
  tabDATOS (020,1) = "AREA/SUCURSAL"
  tabDATOS (021,1) = "COBERTURA SEGURO"
  tabDATOS (022,1) = "CODIGO"
  tabDATOS (023,1) = "F.FINALIZACION"
  tabDATOS (024,1) = "CARACTER"
  tabDATOS (025,1) = "PRODUCTO/SERVICIO"
  tabDATOS (026,1) = "RIESGOS ASOCIADOS"
  tabDATOS (027,1) = "AREA SOPORTE"
  tabDATOS (028,1) = "CLIENTE"
  tabDATOS (029,1) = "CUENTA CONTABLE"
  tabDATOS (030,1) = "TIPO DE EVENTO"
  tabDATOS (031,1) = "CATEGORIA"
  tabDATOS (039,1) = "COSTO OPORTUNIDAD"
  tabDATOS (040,1) = "DEPENDENCIA/SUCURSAL"
  tabDATOS (041,1) = "ASIENTO CONTABLE"
  tabDATOS (042,1) = "IMPORTE SIN IMPACTO CONTABLE"

    Celda = ""


 '''''' IF SMODAL    = "" THEN SMODAL = "AAG"

  Strsql = "  select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "
 ' RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rs.eof then

     response.write "**** error evento de perdida no encontrado !!! *****"


  else

     sFECREGIS  = CLNG(DATE()) - clng(cdate("01-01-1900"))
     sCODTRACERB = "DOBLIGA"




'''''''''''''''' ver estado ''''''''''''''''''''''''''''''

            IF SOPCION = "XXXXXXXXXXXXPENDIENTE" then

               wkestado = "ENRF"


            ELSEIF SOPCION = "XXXXXXXXXXXXXFIN" then

               wkestado = "FIN"

            ELSEIF SOPCION = "XXXXXXXXXXXXXXXXXXCANCEL" then

               wkestado = "CAN"

            ELSE

               wkestado = SOPCSET


            END IF




   '''''''' wkestado = rs(17) & ""

    Strsql = "Select ENVIA From ORESTADOS WHERE CODIGO = '" & wkestado & "' "
    set rsXX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT rsXX.eof THEN

       wkestadoENVIA = rsxx(0) & ""

    END IF
'''''''''''''''''''''''''''''''''''''''''''''


  IF  wkestadoENVIA = "S" OR wkestadoENVIA = "X" THEN


     strsql = "SELECT PATRON From ORTIPOCARACTER where OBJID = 0" & rs(24)
   '  RESPONSE.WRITE STRSQL
     set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsniv.eof then
        sInputx = rsniv(0) & ""
        IF  sInputx = "ESP" THEN sCODTRACERB = "DOBLIGA" & rs(24)
     END IF
     call ofv.cerrarconsulta(rsniv)



     Strsql = "  select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACERB & "' "
 '    RESPONSE.WRITE STRSQL
     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rsx1.eof then


        strsql = "Insert into ORTRACER "
        strsql = strsql & " (CODTRACER,FECREGIS,FECORIG,FECIDENT, FECCONTA,"
        strsql = strsql & " TRCPROCESO, TRCPRODSER, TRCNEGOLINE, TRCIDENTACT, TRCEVENTO, "
        strsql = strsql & " TRCOBSERV, TRCIMPBRUTO, TRCIMPACRECEN, TRCIMPOTROS, TRCIMPRECUPERA, "
        strsql = strsql & " TRCIMPNETO, TRCIMPSEGURO,TRCESTADO, TRCUSRREG,"
        strsql = strsql & " fecfinal, trccaracter, trcproducto, trcriesgoasoc, trcsecsoporte,"
        strsql = strsql & " TRCEVENTORIESGO1, TRCEVENTORIESGO2, TRCLOTEING, TRCTIPOING, TRCINFOING,"
        strsql = strsql & " TRCCLIENTE, TRCCTACBLE, trcecoresp, TRCIMPDESAFECTA,TRCDISTRIB,TRCPORCDISTRIB,TRCCOSTOPOR,trcAsientoCont"



        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & sCODTRACERB & "','" & SFECREGIS & "','0',"
        strsql = strsql &  "'0','0',"
        strsql = strsql & "'0','0','0','0','0','0','0','0','0','0','0','0','DOBLIGA',"
        strsql = strsql & "'0',"
        strsql = strsql & "'0','1','0','0','0',"
        strsql = strsql & "'0','0','0','0','0',"
        strsql = strsql & "'0','0','0','0','0','0','0','0'"
        strsql = strsql & ")"
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        Strsql = "  select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACERB & "' "
        set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      END IF



     II=0
     for ii=0 to (rsx1.Fields.Count - 1)


   '   IF UCASE(rs(ii).NAME)<> "TRCEVENTO" THEN
         if ii > 1 then

           if (rsx1(ii) & "") = "1" then

             if (rs(ii) & "") <> "0" and (rs(ii) & "") <> "" then

             else

                ttaberror = ttaberror + 1
                taberror (ttaberror,0) = ii & "&nbsp-&nbsp;" & tabDATOS (ii,1)' & II & "---" & rs(ii).NAME & ":  " & rs(ii) & "///" & rsx1(II).NAME

             end if


           end if

         end if

   '   END IF

     next

''''''''''''''''''''''''''''''''''''''''''''''

'  tabDATOS (002,1) = "F.REGISTRO"
'  tabDATOS (003,1) = "F.INICIO"
'  tabDATOS (004,1) = "F.DESCUBRIMIENTO"
'  tabDATOS (005,1) = "F.R.CONTABLE"
'  tabDATOS (023,1) = "F.FINALIZACION"

'''''''''''''''''''''' VALIDACIONES PRE-BCRA..................

        WKFORFECREG = 0
        WKFORFECINI = 0
        WKFORFECDES = 0
        WKFORFECCON = 0
        WKFORFECFIN = 0

''                ttaberror = ttaberror + 1
''                taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;NO INFORMADA"


    ''' VER F.REGISTRO

        II=2

        if (rs(ii) & "") <> "0" and (rs(ii) & "") <> "" then
            WKFORFECREG = rs(ii)
        END IF


    ''' VER F.INICIO

        II=3

        if (rs(ii) & "") <> "0" and (rs(ii) & "") <> "" then
            WKFORFECINI = rs(ii)
        END IF

    ''' VER F.DESCUBRIMIENTO

        II=4

        if (rs(ii) & "") <> "0" and (rs(ii) & "") <> "" then
            WKFORFECDES = rs(ii)
        END IF

    ''' VER F.DESCUBRIMIENTO

        II=5

        if (rs(ii) & "") <> "0" and (rs(ii) & "") <> "" then
            WKFORFECCON = rs(ii)
        END IF


    ''' VER F.FINALIZACION

        II=23

        if (rs(ii) & "") <> "0" and (rs(ii) & "") <> "" then
            WKFORFECFIN = rs(ii)
        END IF



    ''' VER F.INICIO

        II=3

        if (rsx1(ii) & "") = "1" then

           IF WKFORFECINI < (WKFORFECREG + 1) THEN

           ELSE

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;FECHA NO CORRESPONDE CON F.DE CARGA."

           END IF
        END IF


    ''' VER F.DESCUBRIMIENTO

        II=4

        if (rsx1(ii) & "") = "1" then

           IF WKFORFECDES < (WKFORFECREG + 1) THEN

           ELSE

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;FECHA NO CORRESPONDE CON F.DE CARGA."

           END IF



           IF WKFORFECINI < (WKFORFECDES + 1) THEN

           ELSE

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;FECHA NO CORRESPONDE CON F.DE INICIO."

           END IF





        END IF


    ''' VER F.FINALIZACION

        II=23

        if (rsx1(ii) & "") = "1" then


           IF WKFORFECINI < (WKFORFECFIN + 1) THEN

           ELSE

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;FECHA NO CORRESPONDE CON F.DE INICIO."

           END IF


        END IF






''  tabDATOS (012,1) = "MONTO BRUTO"
''  tabDATOS (021,1) = "COBERTURA SEGURO"


    ''' VER MONTO BRUTO

        II=12

        if (rsx1(ii) & "") = "1" then

            WKFORMONTO = cstr( rs(ii))

           IF IsNumeric(WKFORMONTO)  THEN

           ELSE

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;DATO NO NUMERICO."

           END IF
        END IF

    ''' VER COBERTURA SEGURO

        II=21

        if (rsx1(ii) & "") = "1" then

            WKFORCOBSEG = cstr("0" & rs(ii))

          ''  response.write "cs: " & WKFORCOBSEG & "- in:" & IsNumeric(WKFORCOBSEG) & "- <0:" & (clng(WKFORCOBSEG) < 0) & "<br>"

           IF IsNumeric(WKFORCOBSEG)  THEN

              if  clng(WKFORCOBSEG) < 0 then

                  ttaberror = ttaberror + 1
                  taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;DATO NO NUMERICO."

              end if

           ELSE

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;DATO NO NUMERICO."

           END IF
        END IF

''  tabDATOS (011,1) = "DESCRIPCION"

    ''' VER DESCRIPCION

        II=11

        if (rsx1(ii) & "") = "1" then

            WKFORDESCRI    = cstr("" & rs(ii))
            WKFORDESCRIDSC = ""

           IF LEN(WKFORDESCRI) < 301 THEN
              WKFORDSCL = LEN(WKFORDESCRI)
           ELSE
              WKFORDSCL = 300
              WKFORDESCRI = LEFT(WKFORDESCRI,300)
              WKFORDESCRIDSC = "LONGITUD EXCEDE DE 300 CARACTERES,&nbsp;"
           END IF

           IF (INSTR(WKFORDESCRI,CHR(34)) > 0 )  OR (INSTR(WKFORDESCRI,CHR(35)) > 0 )  OR (INSTR(WKFORDESCRI,CHR(38)) > 0 )  OR (INSTR(WKFORDESCRI,CHR(39)) > 0 )  OR (INSTR(WKFORDESCRI,CHR(64)) > 0 )  OR (INSTR(WKFORDESCRI,CHR(95)) > 0 ) THEN

              ttaberror = ttaberror + 1
              taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;" & WKFORDESCRIDSC & "CARACTERES NO VALIDOS PARA BCRA (01)."

           ELSE

              WWWCCC = 0

              FOR CCC=1 TO LEN(WKFORDESCRI)

                  WKFORDESCRICCC = MID(WKFORDESCRI,CCC,1)

                  IF WKFORDESCRICCC > CHR(125) or WKFORDESCRICCC < CHR(32)  THEN
                     WWWCCC = 1
                     EXIT FOR
                  END IF

              NEXT

              IF WWWCCC > 0 THEN

                 ttaberror = ttaberror + 1
                 taberror (ttaberror,0) = ii & "&nbsp-&nbsp;ADV:" & tabDATOS (ii,1) & "&nbsp-&nbsp;" & WKFORDESCRIDSC & "CARACTERES NO VALIDOS PARA BCRA (02)."

              END IF

           END IF
        END IF


''''''''''''''''''''''''''''''''''''''''''''''
    END IF

     if ttaberror > 0 then
        sHTML = ofv.FormHeader(session("FormN"))

       Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO - CONTROL DE PATRON DE DATOS","si")
       Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","","&nbsp;","si")
       Fila  = ofv.GenRow("","","","","","","",Celda,"si")
       sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

       Filax  = ""
       Celda = ofv.gencelda("","VTI style='background-color:#888888;color:#FFFFFF;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DATOS OBLIGATORIOS FALTANTES","si")
       Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

        for ii=1 to ttaberror

            Celda = ofv.gencelda("","VTI style='background-color:#F0F0F0;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",taberror (ii,0),"si")
            Filax  = Filax & ofv.genrow("","","","","30","","",Celda,"si")



        next

        sAccion = "ORBASERIESGOSADD.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & ""
        FILA2   = "<INPUT TYPE='SUBMIT' VALUE='ACEPTAR' CLASS='BT'> "
        Formu2  = ofv.GenForm("FACC","",sAccion,"","post","",Fila2,"si")

        Celda = ofv.gencelda("","VTI style='background-color:#888888;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",Formu2,"si")
        Filax  = Filax & ofv.genrow("","","","","20","","",Celda,"si")


        sHTML = sHTML & "<CENTER>"
        sHTML = sHTML & "<BR>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:3;border-color:#888888;' ","","50% ","","0","","0","0","","",Filax,"si")
        sHTML = sHTML & "</CENTER>"

       'Botonera
        Agregar = ""
        Buscar  = ""
        Celda   = ""

        sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

        call ofv.cerrarconsulta(rs)
        call ofv.cerrarconn(cn)

        sHTML = sHTML & "</body></html>"



     else



        IF SOPCION = "SET" THEN

           sFECREGIS  = CLNG(DATE()) - clng(cdate("01-01-1900"))

           IF SOPCSET = "CAN" OR SOPCSET = "FIN" THEN

            ''  strsql = "UPDATE ORTRACER SET TRCESTADO = '" & SOPCSET & "', FECFINENR = '" & sFECREGIS & "', "
            ''  strsql = STRSQL & " TRCMARCAACTP = 'E', TRCFECACTP = '" & sFECREGIS & "' "
            ''  strsql = strsql & " WHERE CODTRACER = '" & SCODTRACER & "' "
          ' response.write strsql & "<br>"
            ''  call ofv.crearconsultaEx(StrSql,cn,1,volver)

              strsql = "UPDATE ORTRACER SET TRCESTADO = '" & SOPCSET & "', FECFINENR = '" & sFECREGIS & "' "
              strsql = strsql & " WHERE CODTRACER = '" & SCODTRACER & "' "
          ' response.write strsql & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


           ELSE

            ''  strsql = "UPDATE ORTRACER SET TRCESTADO = '" & SOPCSET & "', "
            ''  strsql = STRSQL & " TRCMARCAACTP = 'E', TRCFECACTP = '" & sFECREGIS & "' "
            ''  strsql = strsql & " WHERE CODTRACER = '" & SCODTRACER & "' "
         '  response.write strsql & "<br>"
            ''  call ofv.crearconsultaEx(StrSql,cn,1,volver)


              strsql = "UPDATE ORTRACER SET TRCESTADO = '" & SOPCSET & "' "
              strsql = strsql & " WHERE CODTRACER = '" & SCODTRACER & "' "
         '  response.write strsql & "<br>"
              call ofv.crearconsultaEx(StrSql,cn,1,volver)


           END IF




           XXXX = ORRECALCEVENTO(sCODTRACER,"NO")

           RESPONSE.REDIRECT "ORBASERIESGOSADD.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & ""




        ELSE



           RESPONSE.REDIRECT "ORBASERIESGOSADD.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & ""



        end if

     END IF



   end if




end if

Response.Write sHTML








%>