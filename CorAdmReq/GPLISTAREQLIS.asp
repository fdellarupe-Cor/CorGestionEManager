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


  CantRegMostrarx = 5
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL       = request.querystring("MODAL")
  sMODALMAIN   = request.querystring("MODALMAIN")
  sCODTRACER  = request.querystring("CODTRACER")
  sFECHA        = request.querystring("FECHA")
  sOPCION        = request.querystring("OPCION")

  if smodal = "" then

     smodal = session("smodal")

  else
     session("smodal") = smodal 

  end if



  if sMODAL = "" then
     sMODAL = "REG"
     session("smodal") = smodal
  end if 

  sHTML = ofv.FormHeader(session("FormN"))

       WSPSESION = "ORWHERE"

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "status=yes,resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

    SCONECTOR = ""
    SWHERE    = ""

    if sMODAL = "REG"  then 
       SCONECTOR = " AND "
    ELSE
       SWHERE    = " WHERE "
    END IF


 ' if sopcion = "SI" then
 '    GENREQ("500")
 ' end if



  if sMODAL = "NO" then



        strsql = "DELETE FROM ORTRACER WHERE TRCESTADO = 'NO' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


     SMODAL = WKVISION

  ELSEif sMODAL = "ACC" then

        strsql = "UPDATE ORTRACER SET TRCESTADO = 'REG' WHERE CODTRACER = '" & SCODTRACER & "' "
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


     SMODAL = WKVISION
  END IF




   wkproccod0 = "&nbsp;"
   wkproccod1 = "&nbsp;"
   wkproccod2 = "&nbsp;"
   wkproccod3 = "&nbsp;"
   wkproccod4 = "&nbsp;"

   wkproccod0S = " STYLE='background-color:#AAAAAA;color:#FFFFFF;' "
   wkproccod1S = ""
   wkproccod2S = ""
   wkproccod3S = ""
   wkproccod4S = ""


   WKFMODULO = "editarPRY.asp?codigo=NUEVOREQ&tipo=REQ"


  if sMODAL <> "WKF"  then


     sAccion = "GPLISTAREQ.asp?MODAL=WKF"

     wkproccod4 = "<A CLASS=NT style='COLOR:#ffffff;' TITLE='Bandeja de pendientes' HREF='" & sAccion & "' >B.PENDIENTES</A>"

     wkproccod4S = " style='background-color:#336699;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "



     Strsql = "Select PERMITEMR,PERMITEMO,PERMITEVT,PERMITEVN From GPREQRESPONSABLES WHERE TIPOUSR = 'U' AND USUARIOSECTOR = '" & SESSION("USUARIO") & "' "
     set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rs2.eof then

        sAccion = "GPLISTAREQ.asp?MODAL=WKF"

        WKFMODULO2 =  WKFMODULO & "&MODAL=MP"

    ''''   wkproccod3 = "<A CLASS=NT TITLE='Nuevo requerimiento' HREF='" & sAccion & "' >NUEVO REQUERIMIENTO</A>"

        wkproccod3 = "NUEVO REQUERIMIENTO"

        wkproccod3S = " style='background-color:#F7EEDD;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "
        wkproccod3S = wkproccod3S & " onclick=" & chr(34) & "AbrirCons('" & WKFMODULO2 & "'); " & chr(34) & " title='Nuevo requerimiento propio.' " 

        IF (RS2(0) & "") = "S" THEN


            strsql = ""

            strsql = strsql & " Select a.* From TIPOREQNV a "
            strsql = strsql & " where  "

             StrSql1 = "Select SECTOR From GPREQRESPONSABLES WHERE USUARIOSECTOR = '" & SESSION("USUARIO") & "' "
             set rsX = ofv.crearconsultaEx(StrSql1,cn,1,parametros)
             IF Not rsX.EOF THEN

                 snivSECTOR = rsX(0)
  
             END IF 

            strsql = strsql & "  ( A.DESTSECTOR = 0" & snivSECTOR & " ) AND (A.ACTIVO = 'S' )"


            strsql = ""

            strsql = strsql & " select DISTINCT A.tiporeqcod from gpreqresponsables c "
            strsql = strsql & "     left join gpreqxusr b on (c.usuariosector = b.usuariosector) "
            strsql = strsql & "     left join tiporeqnv a on (a.tiporeqcod = b.tiporeqcod or a.permiso = 'N') "
            strsql = strsql & "     where c.usuariosector <> 'TODOS' AND c.usuariosector <> '" & SESSION("USUARIO") & "' AND (A.ACTIVO = 'S' ) "
            strsql = strsql & " AND ( A.DESTSECTOR = 0" & snivSECTOR & " ) "

 



            set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)

            IF Not rsX.EOF THEN

               sAccion = "GPLISTAREQ.asp?MODAL=WKF"

               WKFMODULO2 =  WKFMODULO & "&MODAL=MR"


          '''     wkproccod2 = "<A CLASS=NT TITLE='Nuevo requerimiento desde responsable' HREF='" & sAccion & "' >NUEVO REQ.RESPONSABLE</A>"
               wkproccod2 = "NUEVO REQ.RESPONSABLE"

               wkproccod2S = " style='background-color:#F7EEDD;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "
               wkproccod2S = wkproccod2S & " onclick=" & chr(34) & "AbrirCons('" & WKFMODULO2 & "'); " & chr(34) & " title='Nuevo requerimiento desde responsable.' " 


            END IF 


        END IF 



        IF (RS2(1) & "") = "S" THEN

            sAccion = "GPLISTAREQ.asp?MODAL=WKF"

            WKFMODULO2 =  WKFMODULO & "&MODAL=MO"


         '''   wkproccod1 = "<A CLASS=NT TITLE='Nuevo requerimiento desde otro usuario' HREF='" & sAccion & "' >NUEVO REQ.M.AYUDA</A>"
            wkproccod1 = "NUEVO REQ.M.AYUDA"


            wkproccod1S = " style='background-color:#F7EEDD;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "
            wkproccod1S = wkproccod1S & " onclick=" & chr(34) & "AbrirCons('" & WKFMODULO2 & "'); " & chr(34) & " title='Nuevo requerimiento desde otro usuario.' " 

        END IF 


        IF (RS2(2) & "") = "S" OR (RS2(3) & "") = "S" THEN

            sAccion = "GPLISTAREQvista.asp?MODAL=SET"

            wkproccod0 = "<A CLASS=NT TITLE='Modificar / crear nuevas vistas' HREF='" & sAccion & "' >VISTAS</A>"

            wkproccod0S = " style='background-color:#F7EEDD;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "


        END IF 




     END IF









    sForm = "formACT"
    Formulario = ""
    sInput = "&nbsp;"


        WKFVISTAN = "NO EXISTEN VISTAS HABILITADAS PARA ESTE RESPONSABLE"


     sMODALVISTA     = SESSION("CVISTA")

     IF SMODALVISTA = "VISTATEMP" THEN

        WKFVISTAN = "VISTA TEMPORAL"

     ELSE

        Strsql = "Select NOMVISTA From GPVISTAS WHERE CODVISTA = '" & sMODALVISTA & "' "
        set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs2.eof then
           WKFVISTAN = RS2(0)
        END IF
     END IF


     Strsql = ""
     Strsql = Strsql & "Select A.CODVISTA, A.NOMVISTA From GPVISTAS A"
     Strsql = Strsql & " LEFT JOIN GPVISTASARC B ON (A.CODVISTA = B.CODVISTA) "
     Strsql = Strsql & " WHERE A.TIPO = 'REQ' AND "
     Strsql = Strsql & " ( "
     Strsql = Strsql & " A.ESPUBLIC = 'S' OR "
     Strsql = Strsql & " (A.ESPUBLIC = 'N' AND B.USUARIOSECTOR = '" & SESSION("USUARIO") & "' ) OR "
     Strsql = Strsql & " (A.ESPUBLIC = 'N' AND A.USUARIO = '" & SESSION("USUARIO") & "' )  "
     Strsql = Strsql & " ) "
     Strsql = Strsql & " ORDER BY A.NOMVISTA "

     sInput = ofv.GenerarCombo(sForm,"NVISTA",SESSION("CVISTA"),strsql,cn,0,1,"")






    Formulario = ""
    Celda = ""

    Celda = Celda & ofv.gencelda("","UT STYLE='background-color:#AAAAAA;color:#FFFFFF;' ","","10%","","","","VISTA:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;BORDER-COLOR:#888888;' ","","30%","","","",WKFVISTAN,"si")



    sInput0 = ofv.GenerarInput("NROREQ","0","agregar","","6","dtn","si")
    sInput2 = "BUSCAR"
    Celda = Celda & ofv.gencelda("","UT STYLE='background-color:#AAAAAA;color:#FFFFFF;' ","","10%","","","","NUMERO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;BORDER-COLOR:#888888;' ","","5%","","","",sInput0,"si")
    Celda = Celda & ofv.gencelda("","vT onclick=" & chr(34) & "TomaReq(); " & chr(34) & " title='ver Requerimiento.' style='cursor:hand;border-style:SOLID;border-width:1;BORDER-COLOR:#888888;background-color:#ddddff;border-style:outset;border-width:2;' ","CENTER","5%","","","",sInput2,"si")

    Celda = Celda & ofv.gencelda("","UT STYLE='background-color:#AAAAAA;color:#FFFFFF;' ","","10%","","","","MAS VISTAS:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;BORDER-COLOR:#888888;' ","","30%","","","",SINPUT,"si")
    Celda = Celda & ofv.GenCelda("","ut   " & wkproccod0S,"CENTER","10%","","","",wkproccod0,"si") 

    sAccion = "GPLISTAREQ.ASP?MODAL=SET"

    FILAX = ofv.genrow("","","","","25","","",Celda,"si")
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")

  sHTML = sHTML & "" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#000000;' ","","100% ","","0","","0","0","","",Formulario,"si")

  wwwstr = ""

  wwwstr = wwwstr & "<script type='text/javascript'> "

  wwwstr = wwwstr & " function TomaReq() { "
  wwwstr = wwwstr & "  var str='editarPRY.asp?codigo=XXXXXX&tipo=REQ&MODAL=KED'; "
  wwwstr = wwwstr & "  var strx=str.replace('XXXXXX', document.all.NROREQ.value); "
 ' wwwstr = wwwstr & "  alert(document.all.NROREQ.value); "
 ' wwwstr = wwwstr & "  alert(strx); "
  wwwstr = wwwstr & "  AbrirCons(strx); "
  wwwstr = wwwstr & " } "
  wwwstr = wwwstr & "</script> "

  sHTML = sHTML & wwwstr

 ELSE

   sAccion = "GPLISTAREQ.asp?MODAL=SET"

   wkproccod4 = "<A CLASS=NT style='COLOR:#ffffff;' TITLE='Bandeja de entrada' HREF='" & sAccion & "' >B.ENTRADA</A>"

   wkproccod4S = " style='background-color:#336699;border-style:outset;border-width:2;BORDER-COLOR:#ffffff;CURSOR:HAND;' "


 END IF 



     StrSql = "Select A.*"

  if sMODAL = "WKF" then

     StrSql = StrSql & ", B.* " 

  end if 


     StrSql = StrSql & " From REQUERIMIENTO A  "

  if sMODAL = "WKF" then

     StrSql = StrSql & " INNER JOIN CKTVIEWNV B ON (B.OBJTIPO = 'REQ' AND B.OBJCOD = A.CODIGO) " 

  end if 

     StrSql = StrSql & " WHERE (A.VERSION = '2' ) " & SESSION("ORWHERE")
     StrSql = StrSql & " ORDER BY PRIORIDAD DESC, ID DESC  "


 ''' response.write strsql & "<br>"

 ''' sHTML = sHTML & "<br><p style='border-style:solid;border-width:1;line-height:20px;' >&nbsp;" &  SESSION("ORWHERE") & "&nbsp;</p>"


     WKACCION = " - BANDEJA DE ENTRADA"

     if sMODAL = "WKF"  then      WKACCION = " - BANDEJA DE PENDIENTES"
     if sMODAL = "CSN"  then      WKACCION = " - CONSULTA DE EVENTOS DE PERDIDA"


   Celda = ""


    
   Celda = Celda & ofv.GenCelda("","uti style='color:#ffffff;'  ","","40%","","","","LISTA DE REQUERIMIENTOS  " & WKACCION,"si")
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod1S,"CENTER","15%","","","",wkproccod1,"si")  
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod2S,"CENTER","15%","","","",wkproccod2,"si") 
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod3S,"CENTER","15%","","","",wkproccod3,"si") 
   Celda = Celda & ofv.GenCelda("","ut   " & wkproccod4S,"CENTER","15%","","","",wkproccod4,"si") 
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","aaa bgcolor='#888888' style='border-style:solid;border-width:1;BORDER-COLOR:#000000;' valign='middle' align='center' ","","100% ","30","0","","2","1","","",Fila,"si")

 
 ' sHTML = sHTML & "<br>"



  'Encabezados

   Formulario = ""
   FILAX = ""
   Celda = ""

     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        rs.move(CantRegMoverx)
  posi = 0
        formulario = formulario & GENLISTA(RS)



     ELSE

        celdas = ofv.GenCelda("","VT colspan=7","center","","","","","<BR><BR><BR>NO HAY REQUERIMIENTOS PARA LA SELECCION<BR><BR><BR>","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     END IF





  sHTML = sHTML  & ofv.GenTabla("","","","100%","","0","","1","1","","",formulario,"si")


  'Botonera
  Agregar = ""
'  if bAgregar and sMODAL <> "WKF" then
'    sAccion = "GPLISTAREQLIS.asp?Opcion=SI&OPC=1&TIPO=REQ&RESP=" & SRESP & "&MODAL=" & SMODAL
'    Agregar = REPLACE(ofv.BotonAgregar(sAccion),"Agregar","NUEVO " & STABLA)
'  end if

  Buscar = ""
  
  VOLVER = ""

'  xxxx    = CorGenBotonXp(3, "NO", "NO", "NO", "NO", "Bandeja de Pendientes", "NO", "Pendientes")
'  sAccion = "GPLISTAREQ.asp?MODAL=WKF"
'  VOLVER = ofv.GenBoton(3, "esp1Bot", "BT", "S", "", "S", "S", "N", "esp1Frm", sAccion & "")
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




end if

Response.Write sHTML



FUNCTION GENREQ(XTIPOREQ)

    CODREQ = GTREQ()



if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

     XSECTORIG  = "0"
     XSECTORIGN = 0
     XCODINTER  = "0"

     StrSql = "Select * From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
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


    strsql = "insert into REQUERIMIENTO (CODIGO,TITULO,FECHAINI,FECHASOL,FECHAENT,FECHAFIN,ORIGEN,TIPO,"
    strsql = strsql & "TEMASUBTEMA,DESCRIPCION,ESTADO,DESCTIPO,ORIGINANTE,ASIGNADO,TIPOREQ,FECHATEN,PRIORIDAD,CODINT,VERSION) "
    strsql = strsql & "values ('" & CODREQ & "','NUEVO REQUERIMIENTO','" & vNuevoValor & "','0','0','0','0','0','0','DESCRIPCION','1','NO','" & SESSION("USUARIO") & "','NO','" & XTIPOREQ & "','0','3','" & XCODINT & "','2')"
   '' response.write strsql & "<br>"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)



''' crea subdirectorio de la carpeta

  Set fso = CreateObject("Scripting.FileSystemObject")

  sPath = server.mappath("./Legajos")

  Set Folder = fso.getfolder(sPath)
  set sfolder = Folder.SubFolders
  sfolder.add(CODREQ)
  xxx = GENSECTOROR(CODREQ)



END FUNCTION

FUNCTION GENSECTOROR(XCODREQ)



  StrSql = "Select * From REQUERIMIENTO where codigo = '" & XCODREQ & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rsX.eof then
     XTIPOREQ = RSX(15)
     XUSRORIG = RSX(13)
     XSECTORIG = 0

     StrSql = "Select * From USUARIOS where ID = '" & XUSRORIG & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG = rsX2(21)
        XSECTORIGB = XSECTORIG
        XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTORIGB
''RESPONSE.WRITE STRSQL
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

        STRSQL = "INSERT INTO REQSECTORES (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC) VALUES ("  
        STRSQL = STRSQL & "'" & XCODREQ & "','10','" & XSECTORIG & "','" & XSECTGEREN & "','OR','','" & XUSRORIG & "')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

     END IF


     StrSql = "Select * From TIPOREQ where TIPOREQ = '" & XTIPOREQ & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG  = rsX2(34)
        XSECTGEREN = rsX2(33)

        STRSQL = "INSERT INTO REQSECTORES (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC) VALUES ("  
        STRSQL = STRSQL & "'" & XCODREQ & "','20','" & XSECTORIG & "','" & XSECTGEREN & "','AN','','NO')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

     END IF







  END IF



END FUNCTION





FUNCTION GENLISTA(RSX)



   xTrColor = "#FFFFFF"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "ut"

    if bSuprimir  AND RSX(11) = 1 then
      sHref = "PRYREQDELETE.asp?ID=" & rsX(1) & "&CantRegMover=" & CantRegMoverx
      CeldaS = ofv.BotonEliminar( "REQUERIMIENTO: " & rsX(2) & " - (ID:" & RSX(1) & ")",sHref)
    else
      CeldaS = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if



 






    Filax  = ""

    Celda = ""

 

    Celda = ""


    'CODIGO
    SINPUT = rsx(0)


   WKFMODULO = "editarPRY.asp?codigo=" & rsx(1) & "&tipo=REQ"

    IF SMODAL = "WKF" THEN
       WKFMODULO = WKFMODULO & "&SEGCKTPSC=" & RSX("SEGCKTPSC") & "&MODAL=WK"
    ELSE
       WKFMODULO = WKFMODULO & "&MODAL=ED"

    END IF


    Celda = Celda & ofv.gencelda("","vT onclick=" & chr(34) & "AbrirCons('" & WKFMODULO & "'); " & chr(34) & " title='ver Requerimiento.' style='cursor:hand;background-color:#ddddff;border-style:outset;border-width:2;' ROWSPAN=4  ","center","","","","",sInput,"si")









    'titulo

   ''' sInputx = "REQUERIMIENTO:&nbsp;" & rsx(2)
    sInputx = rsx(2)


    sInput = sInputx
  ''  Celda = Celda & ofv.gencelda("","VTi style='background-color:#F7EEDD;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","REQUERIMIENTO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","VTi COLSPAN=2 style='background-color:#F7EEDD;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'F.INICIO
    if clng(rsx(3)) = 0 then
      dFechan = "---"
    else
      dFechan = GTFechadorSH(rsx(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","INICIO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","25%","","","",sInput,"si")


    'F.FIN
    if clng(rsx(4)) = 0 then
      dFechan = "---"
    else
      dFechan = GTFechadorSH(rsx(4))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","FIN:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")




    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")





    Celda = ""
    'CODIGO INTERNO
       sinput = "<FONT COLOR=#336699 SIZE=1><U>" &  rsx(19) & "</U></FONT>"
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","COD.INTERNO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'ORIGINANTE
      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(13) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","ORIGINANTE:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



    'RESPONSABLE
      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(14) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","RESPONSABLE:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""
    'SISTEMA
      sniv = "---"

                strsql = "Select DESCRI From GPSISTEMAS where CODSIST = '" & RSX(21) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then sniv = rsniv(0)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","SISTEMA:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'MODULO
      sniv = "---"

                strsql = "Select DESCRI From GPMODULOS where CODMOD = '" & RSX(22) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then sniv = rsniv(0)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","MODULO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'ESTADO
      sniv = "---"

                strsql = "Select * From GPESTADOS where ID = 0" & RSX(11)
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

    sInput = sniv

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","ESTADO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","uTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""
    'DESCRIPCION
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' COLSPAN=3 ","","","","","","DESCRIPCION","si")

  '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
    'OBSERVACIONES


    SINPUT = RSX(10) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DESCRIPCION:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


  if SMODAL = "WKF" THEN


    Celda = ""

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#F7AA22;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","WORKFLOW&nbsp;","si")

    'PASO


      sInput = RSX("PSCDSC")
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PASO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'VIENE DE
      sniv = "---"
      sRSP = "---"

                strsql = "Select PSCDSC,RESPASG From CKTVIEWNV where SEGPASOSIG = '" & RSX("SEGCKTPSC") & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then
                   sniv = rsniv(0)
                   sRSP = rsniv(1)
                END IF
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","VIENE DE:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'EMISOR
      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & sRSP & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)


    sInput = sniv

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#EEEEEE;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","EMISOR:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","uTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF




    sInput = "&nbsp;"
    Celda = ofv.gencelda("","UT colspan=7 style='border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor='#fff9f9' ","","","25","","",Celda,"si")




    formu = formu & Filax

 ''     celdas = ofv.GenCelda("","AAA COLSPAN=7 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;FONT-SIZE:5pt' ","","1%","","","","&nbsp;","si")

 ''   formu = formu & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF

    posi = posi + 1
    rsX.Movenext
  loop



  GENLISTA = FORMU

END FUNCTION










%>