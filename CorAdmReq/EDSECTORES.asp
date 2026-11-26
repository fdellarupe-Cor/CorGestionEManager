<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



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

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

 sopcion = request.querystring("opcion")

 if sopcion = "SI" THEN
    sECTOR = request.querystring("SECTOR")
     GENSECTORAF(sECTOR)
 END IF

     ESTADO = 99 

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

     ESTADO = RS(11) 

  END IF


  StrSql = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' ORDER BY ORDEN "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     GENSECTOROR()

     StrSql = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' ORDER BY ORDEN "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF

    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0128;</FONT></B>"
    sImagenA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;GERENCIAS Y SECTORES AFECTADOS","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","ttr  ","","","","","","","si")
  celdas = celdas & ofv.GenCelda("","ttr  ","","","","","","GERENCIA","si")
  celdas = celdas & ofv.GenCelda("","ttr  ","","","","","","SECTOR","si")
  celdas = celdas & ofv.GenCelda("","ttr  ","","","","","","TIPO","si")
  filas = filas & ofv.GenRow("","","","","22","","",celdas,"si")


  DO UNTIL RS.EOF


    'Eliminar
    if bSuprimir AND RS(5) = "AF" AND ESTADO = 1 then

          SELIM = "S" 
    else
          SELIM = "N" 
    end if


       IF SELIM = "S" THEN  
         sHref = "asp/EliminarRegistro.asp?ID=" & rs(0)
         sHref = sHref & "&Tabla=REQSECTORES&volver=" & vuelta & "&Eliminar=NSNC"
         CeldaS = ofv.BotonEliminar(" GERENCIA/SECTOR AFECTADO ",sHref)
         CeldaS = REPLACE(CeldaS,"<TD","<TD WIDTH=1%  ")
       ELSE
         CeldaS = ofv.GenCelda("","UT  ","","1%","","","","&nbsp;","si")
       END IF



    if RS(4) = 0 then
      celdas = celdas & ofv.GenCelda("","ut","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(4)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")


    end if

    if RS(3) = 0 then
      celdas = CELDAS & ofv.GenCelda("","ut","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(3)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = CELDAS & ofv.GenCelda("","ut","","","","","",sInput,"si")


    end if

    SELECT CASE RS(5)
           CASE "OR"
                 sInput = sImagenO & "&nbsp;<b>ORIGINANTE</b>"
           CASE "AN"
                 sInput = sImagenR & "&nbsp;<b>INTERVINIENTE</b>"
           CASE "AF"
                 sInput = sImagenA & "&nbsp;<b>AFECTADO</b>"
    END SELECT

     celdas = CELDAS & ofv.GenCelda("","ut","","","","","",sInput,"si")





      filas = filas & ofv.GenRow("","","","","20","","",celdas,"si")





    RS.MOVENEXT
    LOOP





    Celda = ofv.GenCelda("","abc colspan=4 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","-1","-1","","",filas,"si")

  
    sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1

  Buscar = ""
  if bAgregar then
     sAccion = "REQsectoresBUS.asp?VASP=ES"
     Buscar = ofv.BotonAgregar(sAccion)
  end if


  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENSECTOROR()



  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
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
        STRSQL = STRSQL & "'" & session("NCliente") & "','10','" & XSECTORIG & "','" & XSECTGEREN & "','OR','','" & XUSRORIG & "')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

     END IF


     StrSql = "Select * From TIPOREQ where TIPOREQ = '" & XTIPOREQ & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XSECTORIG  = rsX2(34)
        XSECTGEREN = rsX2(33)

        STRSQL = "INSERT INTO REQSECTORES (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','20','" & XSECTORIG & "','" & XSECTGEREN & "','AN','','NO')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql

     END IF







  END IF



END FUNCTION

FUNCTION GENSECTORAF(XSECTOR)


  StrSql = "Select MAX(ORDEN) From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'AF' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     XORDEN = 30
  ELSE

     XORDEN = RS(0)

     IF XORDEN <> "" THEN
        XORDEN = XORDEN + 10
     ELSE
        XORDEN = 30
     END IF

  END IF



     XSECTORIG = XSECTOR
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

        STRSQL = "INSERT INTO REQSECTORES (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & XORDEN & "','" & XSECTORIG & "','" & XSECTGEREN & "','AF','','NO')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql











END FUNCTION

%>