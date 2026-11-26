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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")
 sMODAL  = request.querystring("MODAL")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

''RESPONSE.WRITE SOPCION & " " & SACCION

 SUSRRESPON = "SIN ASIGNAR"
 SORGUSUARIO = "SIN ASIGNAR"

 CARGAR = "S"

 IF SMODAL = "MP" THEN
    SORGUSUARIO = SESSION("USUARIO")
 ELSE
    SORGUSUARIO = REQUEST.FORM("DATUSUARIOORG") & ""
    IF SMODAL = "MR" THEN
       SUSRRESPON = SESSION("USUARIO")
    end if

    IF SORGUSUARIO = "0" OR SORGUSUARIO = "1" OR SORGUSUARIO = "" THEN
       CARGAR = "N"
    END IF


 END IF

     IF (SMODAL & "") = "MR" THEN

         StrSql1 = "Select SECTOR From GPREQRESPONSABLES WHERE USUARIOSECTOR = '" & SESSION("USUARIO") & "' "
         set rsX = ofv.crearconsultaEx(StrSql1,cn,1,parametros)
         IF Not rsX.EOF THEN

            snivSECTOR = rsX(0)

         END IF

     END IF


 SELECT CASE SOPCION

        CASE "RELOAD"
          CARGAR = "S"
          SACCION = "S"



        CASE "CREAR"

              SUSRRESPON  = REQUEST.FORM("DATRESPON") & ""
              SORGUSUARIO = REQUEST.FORM("DATORIGIN") & ""

              IF SUSRRESPON   = "SIN ASIGNAR" THEN SUSRRESPON  = ""
              IF SORGUSUARIO  = "SIN ASIGNAR" THEN SORGUSUARIO  = ""

              SORGCONTACTO = REQUEST.FORM("DATCONTACTO") & ""
              SORGTIPOREQ  = REQUEST.FORM("DATTIPOREQ") & ""
              SORGREQUER   = REQUEST.FORM("DATREQUER") & ""
              SORGSISTEMA  = REQUEST.FORM("DATSISTEMA") & ""
              SORGDESCRI   = REQUEST.FORM("DATDESCRI") & ""

              SCODREQ = GENREQ(SORGTIPOREQ,SORGUSUARIO)
              SESSION("TCLIENTE") = "REQ"
              SESSION("NCLIENTE") = SCODREQ

              RESPONSE.REDIRECT "EDREQUERNV.asp?SKIND=1"


 END SELECT



    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4so = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"
    sImagen4sA = "<B><FONT FACE=WEBDINGS COLOR=#AAAAAA SIZE=2>&#0128;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenut = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>&#0192;</FONT></B>"
    sImagendu = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenfi = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>R</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0236;</FONT></B>"
    sImagenS = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenES = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>&#0232;</FONT></B>"
    sImagenP = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>2</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=3>&#0237;</FONT></B>"
    sImagenPR = "<B><FONT FACE=SYMBOL COLOR=#AA0000 SIZE=4>!</FONT></B>"
    sImagenin = "<B><FONT FACE=WebDINGS COLOR=#aa0000 SIZE=4>s</FONT></B>"





sHTML = ""




   sHTML = ofv.FORMHeader(SESSION("FORMN")) & corstyler2()


 IF CARGAR = "N" THEN

    filas  = ""

    SFORM = "GPCREAREQN"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    celdas = ""



    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;FECHA:","si")


    dFecha = GTFechadorsh(CLNG(DATE()) - CLNG(CDATE("01-01-1900")))


    sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")





    celdas = ""

    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ESTADO:","si")

    wkfestado = "1"

      sInstrumento = "---"

        Strsql = "Select descri From reqestados Where id = " & wkfestado
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento


    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")

 IF (SMODAL & "") <> "MP" THEN


    celdas = ""



    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;MODO CARGA:","si")

   WKFMODOCARGA = (SMODAL & "")

   SELECT CASE WKFMODOCARGA

          CASE "MP"
                sInput = "ORIGINANTE"

          CASE "MR"
                sInput = "RESPONSABLE"

          CASE "MO"
                sInput = "OTRO USUARIO"

          CASE "AC"
                sInput = "CARGA AUTOMATIZADA"

          CASE "AV"
                sInput = "VUELCO"

          CASE ELSE
                sInput = "OTROS"


   END SELECT

      celdas = celdas & ofv.GenCelda("","uT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")

    celdas = ""


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;USR.CARGA:","si")

      sInput = SESSION("USUARIO")


      strsql = "Select * From USUARIOS where ID = '" & sInput & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then
            sniv = rsniv(1)
            sInput = sniv
        END IF
        call ofv.cerrarconsulta(rsniv)



      celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")



 END IF

    celdas = ""


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ORIGINANTE:","si")



     strsql = ""

     strsql = strsql & " select DISTINCT c.usuariosector from gpreqresponsables c "
     strsql = strsql & "     left join gpreqxusr b on (c.usuariosector = b.usuariosector) "
     strsql = strsql & "     left join tiporeqnv a on (a.tiporeqcod = b.tiporeqcod or a.permiso = 'N') "
     strsql = strsql & "     where c.usuariosector <> 'TODOS' AND c.usuariosector <> '" & SESSION("USUARIO") & "' AND (A.ACTIVO = 'S' ) "

     IF (SMODAL & "") = "MR" THEN

         strsql = strsql & " AND ( A.DESTSECTOR = 0" & snivSECTOR & " ) "

     END IF


     SCOMSEP = ""
     scomK   = ""
     scomD   = ""



  ''''   StrSql = "Select * From GPREQRESPONSABLES WHERE USUARIOSECTOR <> 'TODOS' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     Do While (Not rsX.EOF)

         ''''   sCOD = rsX(2)
            sCOD = rsX(0)
            sniv = sCOD

        strsql = "Select * From USUARIOS where ID = '" & sCOD  & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then
            sniv = rsniv(1)
        END IF
        call ofv.cerrarconsulta(rsniv)

         scomD = scomD & SCOMSEP & sniv
         scomK = scomK & SCOMSEP & sCOD
         SCOMSEP = chr(9)


     RSX.MOVENEXT
     LOOP


    sInput = ofv.gencomboFX(sForm,"DATUSUARIOORG","1",scomK,scomD,"no0")


    celdas = celdas & ofv.GenCelda("","uT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")




   celdas = ""
   celdas = celdas & ofv.GenCelda("","UT ","","60%","","","","&nbsp;","si")
   celdas = celdas & ofv.GenCelda("","Ut onclick='document." & SFORM & ".submit();' style='background-color:#F7AA22;color:#ffffff;CURSOR:HAND;border-style:OUTSET;border-width:2;' TITLE='Ingresar solicitante.' ","CENTER","10%","","","","&nbsp;INGRESAR&nbsp;SOLICITANTE&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")


    if bModificar then
      sAccion = "GPCREAREQ.asp?MODAL=" & SMODAL
      formularioX = ofv.GenForm(SFORM,"",sAccion,"","post","",filas,"si")
    else
      formularioX = ofv.GenForm("","","","","post","",filas,"si")
    end if



    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#FFFFFF;BACKGROUND-color:#F0F0F0;' ","","60%","","0","","0","0","","",formularioX,"si")
    sHTML = sHTML & "</center>"

  END IF


  IF CARGAR = "S" THEN

     filas  = ""

    SFORM = "GPCREAREQ"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    celdas = ""



    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;FECHA:","si")


    dFecha = GTFechadorsh(CLNG(DATE()) - CLNG(CDATE("01-01-1900")))


    sInput = dFecha

    sInput = sInput & ofv.GenerarInput("DATRESPON",SUSRRESPON,"HIDDEN",SFORM,"1","dt","")
    sInput = sInput & ofv.GenerarInput("DATORIGIN",SORGUSUARIO,"HIDDEN",SFORM,"1","dt","")

    celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ORIGINANTE:","si")

        sniv = "SIN ASIGNAR"

      strsql = "Select * From USUARIOS where ID = '" & SORGUSUARIO & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
        sInput = sniv

      celdas = celdas & ofv.GenCelda("","uT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""

    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;RESPONSABLE:","si")

        sniv = "SIN ASIGNAR"

      strsql = "Select * From USUARIOS where ID = '" & SUSRRESPON & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sniv


      celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""

    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ESTADO:","si")

    wkfestado = "1"

      sInstrumento = "---"

        Strsql = "Select descri From reqestados Where id = " & wkfestado
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento


    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

 IF (SMODAL & "") <> "MP" THEN


    celdas = ""



    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;MODO CARGA:","si")

   WKFMODOCARGA = (SMODAL & "")

   SELECT CASE WKFMODOCARGA

          CASE "MP"
                sInput = "ORIGINANTE"

          CASE "MR"
                sInput = "RESPONSABLE"

          CASE "MO"
                sInput = "OTRO USUARIO"

          CASE "AC"
                sInput = "CARGA AUTOMATIZADA"

          CASE "AV"
                sInput = "VUELCO"

          CASE ELSE
                sInput = "OTROS"


   END SELECT

      celdas = celdas & ofv.GenCelda("","uT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;USR.CARGA:","si")

      sInput = SESSION("USUARIO")


      strsql = "Select * From USUARIOS where ID = '" & sInput & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then
            sniv = rsniv(1)
            sInput = sniv
        END IF
        call ofv.cerrarconsulta(rsniv)



      celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;CONTACTO:","si")

     strsql = "Select * From GPMODOCONTACTO Order By DESCRI"
     sInput = ofv.GenerarCombo(SFORM,"DATCONTACTO","1",strsql,cn,0,1,"no")


    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")








 END IF



    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

    celdas = ofv.GenCelda("","VtI style='background-color:#888888;color:#ffffff;'  colspan=2","","","","","","SOLICITUD DEL REQUERIMIENTO","si")
    filas  = filas & ofv.GenRow("","","","","25","","",celdas,"si")

    celdas = ""


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;TIPO:","si")

     strsql = ""

     strsql = strsql & " Select a.*, b.usuariosector From TIPOREQNV a "
     strsql = strsql & " left join GPREQXUSR b on (a.tiporeqcod = b.tiporeqcod) "
     strsql = strsql & " where ((a.permiso = 'N') or (a.permiso = 'S' and b.usuariosector = '" & SORGUSUARIO & "' )) AND (A.ACTIVO = 'S' ) "


     IF (SMODAL & "") = "MR" THEN

         strsql = strsql & " AND ( A.DESTSECTOR = 0" & snivSECTOR & " ) "

     END IF

     strsql = strsql & " Order By a.DESCRIPCION "

    ' response.write strsql & "<br>"

  ''''''   strsql = "Select * From TIPOREQNV Order By DESCRIPCION"


     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     scomK1   = "0"
     scomD1   = ""

     GPTOTREQ = 0

     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF NOT RSX.EOF THEN
        scomK1   = rsX(1)
        scomD1   = rsX(1) & "&nbsp;-&nbsp;" & rsX(2)

     END IF
     Do While (Not rsX.EOF)

      GPTOTREQ = GPTOTREQ + 1

         scomD = scomD & SCOMSEP & rsX(1) & "&nbsp;-&nbsp;" & rsX(2)
         scomK = scomK & SCOMSEP & rsX(1)
         SCOMSEP = chr(9)


     RSX.MOVENEXT
     LOOP


     IF GPTOTREQ = 1 THEN

       sInput = ofv.GenerarInput("DATTIPOREQ",scomK1,"HIDDEN",SFORM,"1","dt","") & scomD1

     ELSE

       sInput = ofv.gencomboFX(sForm,"DATTIPOREQ",scomK1,scomK,scomD,"no0")



     END IF

   ''''  sInput = ofv.GenerarCombo(SFORM,"DATTIPOREQ","1",strsql,cn,1,2,"no")


    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;REQUERIMIENTO:","si")


     sInput = ofv.GenerarInput("DATREQUER","NUEVO REQUERIMIENTO","agregar",SFORM,"70","dt","")

    celdas = celdas & ofv.GenCelda("","utI bgcolor='#f7eedd' style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;SISTEMA:","si")


      strsql = "Select * From GPSISTEMAS Order By DESCRI"
      sInput = ofv.GenerarCombo(SFORM,"DATSISTEMA","1",strsql,cn,1,2,"no")


    celdas = celdas & ofv.GenCelda("","UtI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","AA COLSPAN=2 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#bbbbbb;font-size:6pt;' ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    Celda = ofv.GenCelda("","VtI style='background-color:#888888;color:#ffffff;' colspan=2 ","","100%","","","","DESCRIPCION DEL PEDIDO","si")
    FilaS =  filaS & ofv.GenRow("","","","","25","","",Celda,"si")

    wdescri = "DESCRIPCION DETALLADA"

    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;FONT-SIZE:7PT;' "
    t3b = t3b & "name=DATDESCRI cols=120 rows=8 "
    if bModificar then
  '''    t3b = t3b & " onblur=Usuarios.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = ofv.GenCelda("","UT colspan=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","center","100%","","","",t3b,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")




    if bModificar then
      sAccion = "GPCREAREQ.asp?OPCION=CREAR&MODAL=" & SMODAL
      formularioX = ofv.GenForm(SFORM,"",sAccion,"","post","",filas,"si")
    else
      formularioX = ofv.GenForm("","","","","post","",filas,"si")
    end if



    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#FFFFFF;BACKGROUND-color:#F0F0F0;' ","","100%","","0","","0","0","","",formularioX,"si")
    sHTML = sHTML & "</center>"







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






 '  celdas = celdas & ofv.GenCelda("","Ut onclick='document.reqobs.submit();' style='background-color:#888888;color:#ffffff;CURSOR:HAND;' TITLE='Agregar observaciones.' ","CENTER","10%","","","","&nbsp;OBSERVACIONES&nbsp;","si")


 '  celdas = celdas & ofv.GenCelda("","Ut onclick='document.reqobst.submit();' style='background-color:#888888;color:#ffffff;CURSOR:HAND;' TITLE='visualizar todas las observaciones.' ","CENTER","10%","","","","&nbsp;VER&nbsp;OBSERVACIONES&nbsp;","si")

 '  celdas = celdas & ofv.GenCelda("","vtI style='background-color:#F0F0F0;color:#993333;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","80%","","","",reqmsg,"si")


'

   celdas = celdas & ofv.GenCelda("","Ut onclick='disp_confirm();' style='background-color:#F7AA22;color:#ffffff;CURSOR:HAND;border-style:OUTSET;border-width:2;' TITLE='Confirmar el pedido para su tratamiento.' ","CENTER","10%","","","","&nbsp;CONFIRMAR&nbsp;","si")

'

   filas = ofv.GenRow("","","","","20","","",celdas,"si")





    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#FFFFFF;' ","","100%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"

''EDWKFSTART.ASP?OPCION=AUTO


    sHTML = sHTML & "<form style='display:none;' name='reqadj' action='edarchivosnv.asp'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "<form style='display:none;' name='reqobs' action='edobservanv.asp'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "<form style='display:none;' name='reqobst' action='edobservatodonv.asp'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"


    sHTML = sHTML & "<form style='display:none;' name='reqwkf' action='GPCREAREQ.ASP?OPCION=CREAR'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "<form style='display:none;' name='reqhist' action='EDHISTWKFNV.asp'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"


 '' end if

 END IF




    sHTML = sHTML & "<form style='display:none;' name=init action='requerindexnv.asp' target='WK3' "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "</center>"



    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm()"
    sHTML = sHTML & "  {"

    sHTML = sHTML & "  var est=0;"
    sHTML = sHTML & "  var emsg='DEBE INFORMAR:\n\n';"

   if smodal <> "MP" then

    sHTML = sHTML & "  if (document.all.DATCONTACTO.value!='1')"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "   est=1;"
    sHTML = sHTML & "   emsg= emsg + '- CONTACTO\n';"
    sHTML = sHTML & "    }"

  end if



    sHTML = sHTML & "  if (document.all.DATTIPOREQ.value!='0')"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "   est=1;"
    sHTML = sHTML & "   emsg= emsg + '- TIPO DE REQUERIMIENTO\n';"
    sHTML = sHTML & "    }"

    sHTML = sHTML & "  if (document.all.DATREQUER.value!='NUEVO REQUERIMIENTO')"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "   est=1;"
    sHTML = sHTML & "   emsg= emsg + '- TITULO DEL REQUERIMIENTO\n';"
    sHTML = sHTML & "    }"

    sHTML = sHTML & "  if (document.all.DATSISTEMA.value!='1')"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "   est=1;"
    sHTML = sHTML & "   emsg= emsg + '- SISTEMA\n';"
    sHTML = sHTML & "    }"

    sHTML = sHTML & "  if (document.all.DATDESCRI.value!='DESCRIPCION DETALLADA')"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "   est=1;"
    sHTML = sHTML & "   emsg= emsg + '- DESCRIPCION\n';"
    sHTML = sHTML & "    }"


    sHTML = sHTML & "  if (est==1)"
    sHTML = sHTML & "     {"
    sHTML = sHTML & "    alert(emsg);"
    sHTML = sHTML & "     }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "     {"
    sHTML = sHTML & "  var r=confirm('Confirma el pedido?');"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    document." & SFORM & ".submit();"
 ''   sHTML = sHTML & "    alert('todo ok!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Cancelo el envio!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "     }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"



 '' RESPONSE.WRITE CARGARF & " " & SACCIONF


' IF CARGARF = "S" AND SACCIONF = "S" THEN

    shtml = shtml & "<script  language=javascript>" & chr(13)
    shtml = shtml & " document.init.submit(); " & chr(13)
    shtml = shtml & " </script>" & chr(13)

' END IF


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

'  Agregar = Celt1


 ' xxxx    = CorGenBotonXp(3, "NO", "NO", "NO", "NO", "Historial de circuitos", "NO", "Historial")
 ' sAccion = "EDHISTWKFNV.asp"
 ' Agregar = ofv.GenBoton(3, "esp1Bot", "BT", "S", "", "S", "S", "N", "esp1Frm", sAccion & "")

  Agregar = ""

  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''

  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENREQ(XTIPOREQ,XTORIGIN)

    CODREQ = GTREQ()


    vNuevoValor = clng(DATE()) - clng(cdate("01-01-1900"))

     XSECTORIG  = "0"
     XSECTORIGN = 0
     XCODINTER  = "0"

     StrSql = "Select * From USUARIOS where ID = '" & XTORIGIN & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTORIGN = rsX2(21)
        XSECTORIG  = CSTR(rsX2(21))
     END IF


     IF XSECTORIGN > 0 THEN

     StrSql = "Select CODINTER From ECO where ECO = " & XSECTORIGN
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        XCODINTER = rsX2(0)
     END IF

     END IF


     IF XCODINTER = "" THEN XCODINTER = "0"


     XYEAR = datepart("yyyy",date())



     StrSql = "Select SECUEN From REQCODINT WHERE COD1 = '" & XCODINTER & "' AND COD2 = '" & XYEAR & "' "
   ''  response.write strsql & "<br>"
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsX2.eof then

        XSECUEN = rsX2(0) + 1

        strsql = "UPDATE REQCODINT SET SECUEN = " & XSECUEN
        strsql = strsql & " WHERE COD1 = '" & XCODINTER & "' AND COD2 = '" & XYEAR & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     ELSE

        XSECUEN = 1

        strsql = "insert into REQCODINT (COD1,COD2,COD3,COD4,COD5,SECUEN) "
        strsql = strsql & "values ('" & XCODINTER & "','" & XYEAR & "','NO','NO','NO','1')"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     END IF

    FOR YYY=LEN(XSECUEN) TO 5

        XSECUEN = " " & XSECUEN


    NEXT


    XSECUEN = REPLACE(XSECUEN," ","0")


    XCODINT = XCODINTER & "-" & XSECUEN & "-" & XYEAR


           '   SUSRRESPON  = REQUEST.FORM("DATRESPON") & ""
           '   SORGUSUARIO = REQUEST.FORM("DATORIGIN") & ""
           '   SORGCONTACTO = REQUEST.FORM("DATCONTACTO") & ""
           '   SORGTIPOREQ  = REQUEST.FORM("DATTIPOREQ") & ""
           '   SORGREQUER   = REQUEST.FORM("DATREQUER") & ""
           '   SORGSISTEMA  = REQUEST.FORM("DATSISTEMA") & ""
           '   SORGDESCRI   = REQUEST.FORM("DATDESCRI") & ""


    strsql = "insert into REQUERIMIENTO ("
    strsql = strsql & "CODIGO,TITULO,FECHAINI,FECHASOL,"
    strsql = strsql & "FECHAENT,FECHAFIN,ORIGEN,TIPO,"
    strsql = strsql & "TEMASUBTEMA,DESCRIPCION,ESTADO,DESCTIPO,"
    strsql = strsql & "ORIGINANTE,ASIGNADO,TIPOREQ,FECHATEN,"
    strsql = strsql & "PRIORIDAD,CODINT,VERSION,CODSIST,"
    strsql = strsql & "MODOCARGA,USRCARGA,MODCONTACTO"
    strsql = strsql & ") "
    strsql = strsql & "values ("
    strsql = strsql & "'" & CODREQ & "','" & SORGREQUER & "','" & vNuevoValor & "','0',"
    strsql = strsql & "'0','0','0','0',"
    strsql = strsql & "'0','" & SORGDESCRI & "','1','NO',"
    strsql = strsql & "'" & SORGUSUARIO & "','" & SUSRRESPON & "','" & SORGTIPOREQ & "','0',"
    strsql = strsql & "'3','" & XCODINT & "','2','" & SORGSISTEMA & "',"
    strsql = strsql & "'" & SMODAL & "','" & SESSION("USUARIO") & "','" & SORGCONTACTO & "'"
    strsql = strsql & ")"
   '' response.write strsql & "<br>"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)



''' crea subdirectorio de la carpeta

  Set fso = CreateObject("Scripting.FileSystemObject")

  sPath = server.mappath("./Legajos")

  Set Folder = fso.getfolder(sPath)
  set sfolder = Folder.SubFolders
  sfolder.add(CODREQ)
 ''' xxx = GENSECTOROR(CODREQ)

    GENREQ = CODREQ

END FUNCTION


Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



%>