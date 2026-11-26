<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.formHeader(" ")



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn4 = ofv.conectar(ofv.strconn6)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  WKFAUTO = "WKF0001693621131313343163"
  WKFAPRO = "WKF0001693673726313343163"
  DESLEVL = 9
  SUPLEVL = 9

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
  if not rs.eof then
     Strsql = "Select * From TIPOREQ where TIPOREQ = '" & RS(15) & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        WKFAUTO = RS(37)
        WKFAPRO = RS(38)
        DESLEVL = RS(35)
        SUPLEVL = RS(36)
     END IF

  END IF





 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")
 act = request.querystring("act")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

''RESPONSE.WRITE SOPCION & " " & SACCION


 SELECT CASE SOPCION

        CASE "AUTO"
             SESSION("WKF") = WKFAUTO
             TITULO = "INICIO DE CIRCUITO DE AUTORIZACION "



        CASE "APRO"
             SESSION("WKF") = WKFAPRO
             TITULO = "INICIO DE CIRCUITO DE APROBACION "




 END SELECT

 DIM ATABLA(200,5) 

 DIM ATABSW(5,3) 
 
 TSHW = 0




''' GENERA EL PASO INICIAL

  StrSql = "Select * From WKFCKTMASTER WHERE INI=1 AND WKFCOD = '" & SESSION("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     session("WKFDSC") = rs(1)
''' INICIA EL CIRCUITO 
     XMODEL  = RS(5)

     XRTIPO  = RS(15)
     XRCOD   = RS(16)
     XPASCOD = RS(12)
     XPDSC   = RS(6) & ": " & RS(7) 
     XPDSC1   = RS(6) & ": " & RS(7) 

    if act = "S" then

       set cn2 = ofv.conectar(ofv.strconn0)
       sInputx = "COR ON-TIME"
       strsql = "Select EMAIL From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
       set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       SESSION("ORGMAIL") = SINPUTX

     GENCKTINI()

     GENCKTRESPO(SOPCION)

     SEGCKTPSC = GTCKTPSC()

     XOBSERV = REQUEST.FORM("OBS")

     XRTIPO  = ATABLA(1,2)
     XRCOD   = ATABLA(1,1)
     XRASG   = ATABLA(1,1)
     XRLVL   = ATABLA(1,4)
     XRSCD   = ATABLA(1,5)

    SEGCKTPSCSIG = GTCKTPSC() & "A"
    WXX = GENCKTPSC(SEGCKTPSC,SEGCKTPSCSIG,XMODEL,"SI",XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XRLVL,XRSCD)
    
    SSS = GENACCION(RS(5),"INI")

    TSHW           = TSHW + 1
    ATABSW(TSHW,1) = XPDSC
    ATABSW(TSHW,2) = 1
    ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)


 ''''   SINPUT2 = GETRESPOCKT(ATABLA(1),"SEC")


    StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       XMODEL  = RS2(5)


       IF  ATABLA(0,1) =  0 THEN

              XRTIPO  = ATABLA(2,2)
              XRCOD   = ATABLA(2,1)
              XPASCOD = RS2(12)
              XPDSC   = RS2(6) & ": " & RS2(7) 
              XRASG   = "NO"
              IF XRTIPO = "USR" THEN  XRASG   = XRCOD

              XRLVL   = ATABLA(2,4)
              XRSCD   = ATABLA(2,5)

''' GENERA EL PASO FINAL AL DESTINATARIO OK
              WXX = GENCKTPSC(SEGCKTPSCSIG,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD)

              TSHW           = TSHW + 1
              ATABSW(TSHW,1) = XPDSC
              ATABSW(TSHW,2) = 0
              ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

 '''''''''''''''''''
 ' agregar envio de mail

              XXXX = GENCKTMAIL(XRTIPO,XRCOD,XPDSC & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL)

 '''''''''''''''''''

       ELSE

           IF ATABLA(3,1) =  SESSION("USUARIO") THEN

              XRTIPO  = ATABLA(3,2)
              XRCOD   = ATABLA(3,1)
              XPASCOD = RS2(12)
              XPDSC   = RS2(6) & ": " & RS2(7) 
              XRASG   = ATABLA(3,1)
              XRLVL   = ATABLA(3,4)
              XRSCD   = ATABLA(3,5)

              SEGCKTPSCSIG2 = GTCKTPSC() & "B"

''' GENERAR EL PRIMER PASO OK - NO REQUIERE AUTORIZANTES PARA EL ORIGINANTE
              WXX = GENCKTPSC(SEGCKTPSCSIG,SEGCKTPSCSIG2,XMODEL,"SI",XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XRLVL,XRSCD)

              TSHW           = TSHW + 1
              ATABSW(TSHW,1) = XPDSC
              ATABSW(TSHW,2) = 1
              ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

              IF  ATABLA(0,1) > 1 THEN

                  XRTIPO  = ATABLA(4,2)
                  XRCOD   = ATABLA(4,1)
                  XPASCOD = RS2(12)
                  XPDSC   = RS2(6) & ": " & RS2(7) 
                  XRASG   = "NO"
                  IF XRTIPO = "USR" THEN  XRASG   = XRCOD        
                  XRLVL   = ATABLA(4,4)
                  XRSCD   = ATABLA(4,5)     

''' SI HAY LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS ACTIVA PASO RECURRENTE PARA EL PRIMERO

              ELSE

''' SI NO HAY LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS ACTIVA PASO FINAL AL DESTINATARIO

                  SSS = GENACCION(XMODEL,"APR")

                  StrSql = "Select * From WKFCKTMASTER WHERE WKFPASCOD = '" & XPASCOD & "' "
                  set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                  if not rs3.eof then
                     XMODEL  = RS3(5)
      
                     XRTIPO  = ATABLA(2,2)
                     XRCOD   = ATABLA(2,1)
                     XPASCOD = RS3(12)
                     XPDSC   = RS3(6) & ": " & RS3(7) 
                     XRASG   = "NO"
                     IF XRTIPO = "USR" THEN  XRASG   = XRCOD
                     XRLVL   = ATABLA(2,4)
                     XRSCD   = ATABLA(2,5)     

                  END IF
                  call ofv.cerrarconsulta(rs3)

              END IF 

''' GENERAR EL SIGUIENTE PASO OK
              WXX = GENCKTPSC(SEGCKTPSCSIG2,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD)

              TSHW           = TSHW + 1
              ATABSW(TSHW,1) = XPDSC
              ATABSW(TSHW,2) = 0
              ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

 '''''''''''''''''''
 ' agregar envio de mail
              XXXX = GENCKTMAIL(XRTIPO,XRCOD,XPDSC & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL)

 '''''''''''''''''''



           ELSE

              XRTIPO  = ATABLA(3,2)
              XRCOD   = ATABLA(3,1)
              XPASCOD = RS2(12)
              XPDSC   = RS2(6) & ": " & RS2(7) 
              XRASG   = "NO"
              IF XRTIPO = "USR" THEN  XRASG   = XRCOD
              XRLVL   = ATABLA(3,4)
              XRSCD   = ATABLA(3,5)     

''' GENERAR EL PRIMER PASO AUTORIZACION PARA EL ORIGINANTE OK
              WXX = GENCKTPSC(SEGCKTPSCSIG,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC,"",XRLVL,XRSCD)

              TSHW           = TSHW + 1
              ATABSW(TSHW,1) = XPDSC
              ATABSW(TSHW,2) = 0
              ATABSW(TSHW,3) = GETRESPOCKT(XRCOD,XRTIPO)

 '''''''''''''''''''
 ' agregar envio de mail
              XXXX = GENCKTMAIL(XRTIPO,XRCOD,XPDSC & " POR " & SESSION("NOMREQ"),SESSION("ORGMAIL"),XRLVL)

 '''''''''''''''''''
        

           END IF  

       END IF

    END IF
    call ofv.cerrarconsulta(rs2)

   end if
   call ofv.cerrarconsulta(rs)

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

       sinput = "<a class=tt  href=EDREQUER.asp?OPCION=RELOAD"
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
		 '''     t3b = t3b & " onblur=Usuarios.submit() "
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
		      sAccion = "EDWKFSTART.ASP?CantRegMover=0&OPCION=" & SOPCION & "&ACT=S"
		      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
		    else
		      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
		    end if


		   TAB = TAB  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario0,"si") & "<BR>"


   end if



    Celda = ofv.GenCelda("","","center","","","","",TAB,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")




  ELSE
   SHTML = SHTML & "<P>ERROR EN EL INICIO DEL CIRCUITO</P>"

  END IF





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1
  Buscar  = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''



  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML




FUNCTION GENCKTINI()

    SEGCKT = GTCKT()

    SESSION("SEGCKT") = SEGCKT

    strsql = "insert into SEGCIRCUITO (MODEL,COMIENZO,FIN,OBJTIPO,OBJCOD,SEGCKT,RESPONSABLE,CKTDSC) values "
    STRSQL = STRSQL & " ('" & SESSION("WKF") & "','" & NOW() & "','NO','REQ','" & session("NCliente") & "','" & SEGCKT & "','" & SESSION("USUARIO") & "','" & session("WKFDSC") & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION


FUNCTION GENCKTRESPO(XTIPO)

''' session("SegNivel")

    ATABLA(0,0) =  2
    ATABLA(0,1) =  0

''''''''' 1. GENERA ORIGINANTE DEL CIRCUITO --> USUARIO DE LA SESION. 

    ATABLA(1,0) =  "O"
    ATABLA(1,1) =  SESSION("USUARIO")
    ATABLA(1,2) =  "USR"
    ATABLA(1,3) =  "NO"
    ATABLA(1,4) =  session("SegNivel")
    ATABLA(1,5) =  "0"

    STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
    STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','O','USR','" & SESSION("USUARIO") & "','10','" & session("SegNivel") & "','0')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

''''''''' 2. GENERA DESTINATARIO DEL CIRCUITO --> A: AUTO > SECTOR INTERVINIENTE - NIVEL DESTINO. B: APRO > USUARIO SOLICITANTE - NIVEL 9 

    IF XTIPO = "AUTO" THEN

       STRSQL = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'AN' "
       set rs1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       IF NOT rs1.eof THEN

          STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
          STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','D','SEC','" & RS1(3) & "','999','"
          STRSQL = STRSQL & DESLEVL & "','" & RS1(4) & "')"
          call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          ATABLA(2,0) =  "D"
          ATABLA(2,1) =  RS1(3)
          ATABLA(2,2) =  "SEC"
          ATABLA(2,3) =  "NO"
          ATABLA(2,4) =  DESLEVL
          ATABLA(2,5) =  RS1(4)

       END IF


    ELSE

       STRSQL = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'OR' "
       set rs1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       IF NOT rs1.eof THEN

          STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
          STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','NO','D','USR','" & RS1(7) & "','999','"
          STRSQL = STRSQL & SUPLEVL & "','" & RS1(3) & "')"
          call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          ATABLA(2,0) =  "D"
          ATABLA(2,1) =  RS1(7)
          ATABLA(2,2) =  "USR"
          ATABLA(2,3) =  "NO"
          ATABLA(2,4) =  SUPLEVL
          ATABLA(2,5) =  RS1(3)
       END IF



    END IF






    STRSQL = "Select * From WKFCKTMASTER WHERE TIPO = 'R' AND WKFCOD = '" & SESSION("WKF") & "' ORDER BY ORDEN "
    set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    DO UNTIL rs3.eof 

    XORDEN = 10

    XINDT = 2

    ATABLA(0,0) =  2
    ATABLA(0,1) =  0

 ' DESLEVL = 9
 ' SUPLEVL = 9


''''''''' 3. GENERA AUTORIZANTE DEL ORIGINANTE DEL CIRCUITO --> A: AUTO > SECTOR SOLICITANTE - NIVEL MINIMO SUPERVISION.
'                                                               B: APRO > SECTOR INTERVINIENTE - NIVEL MINIMO SUPERVISION. 
'                                                                 (SOLO SI HAY PASOS RECURRENTES DE AUTORIZACION)



    IF XTIPO = "AUTO" THEN
     IF SUPLEVL < session("SegNivel") THEN
       STRSQL = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'OR' "
       set rs1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       IF NOT rs1.eof THEN
             XORDEN = XORDEN + 10 
             XINDT = XINDT + 1

             ATABLA(XINDT,0) =  "L"
             ATABLA(XINDT,1) =  RS1(3)
             ATABLA(XINDT,2) =  "SEC"
             ATABLA(XINDT,3) =  RS3(5)
             ATABLA(XINDT,4) =  SUPLEVL
             ATABLA(XINDT,5) =  RS1(4)

             strsql = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
             STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','SEC','" & RS1(3) & "','" & XORDEN & "','" & SUPLEVL & "','" & RS1(4) & "')"
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)
       END IF
     ELSE

       XORDEN = XORDEN + 10 
       XINDT = XINDT + 1

       STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
       STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" & SESSION("USUARIO") & "','" & XORDEN & "','" & session("SegNivel") & "','0')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          ATABLA(XINDT,0) =  "L"
          ATABLA(XINDT,1) =  SESSION("USUARIO")
          ATABLA(XINDT,2) =  "USR"
          ATABLA(XINDT,3) =  RS3(5)
          ATABLA(XINDT,4) =  session("SegNivel")
          ATABLA(XINDT,5) =  "0"
     END IF  

    ELSE

     IF SUPLEVL < session("SegNivel") THEN
       STRSQL = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' AND TIPO = 'AN' "
       set rs1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       IF NOT rs1.eof THEN
             XORDEN = XORDEN + 10 
             XINDT = XINDT + 1

             ATABLA(XINDT,0) =  "L"
             ATABLA(XINDT,1) =  RS1(3)
             ATABLA(XINDT,2) =  "SEC"
             ATABLA(XINDT,3) =  RS3(5)
             ATABLA(XINDT,4) =  SUPLEVL
             ATABLA(XINDT,5) =  RS1(4)

             strsql = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
             STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','SEC','" & RS1(3) & "','" & XORDEN & "','" & SUPLEVL & "','" & RS1(4) & "')"
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)

       END IF
     ELSE

       XORDEN = XORDEN + 10 
       XINDT = XINDT + 1

       STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
       STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','USR','" & SESSION("USUARIO") & "','" & XORDEN & "','" & session("SegNivel") & "','0')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          ATABLA(XINDT,0) =  "L"
          ATABLA(XINDT,1) =  SESSION("USUARIO")
          ATABLA(XINDT,2) =  "USR"
          ATABLA(XINDT,3) =  RS3(5)
          ATABLA(XINDT,4) =  session("SegNivel")
          ATABLA(XINDT,5) =  "0"

     END IF  

    END IF


''''''''' 4. GENERA LISTA DE AUTORIZANTES PARA SECTORES AFECTADOS DEL CIRCUITO --> AUTO Y APRO >
'                                                                 SECTOR AFECTADO - NIVEL MINIMO SUPERVISION.
'                                                                  (SOLO SI HAY PASOS RECURRENTES DE AUTORIZACION)



       XBY = " AND  (tipo <> 'AN' AND tipo <> 'OR' ) ORDER BY  ORDEN "

       STRSQL = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' " & XBY
       set rs1 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       DO UNTIL rs1.eof 
             XORDEN = XORDEN + 10 
             XINDT = XINDT + 1

             ATABLA(XINDT,0) =  "L"
             ATABLA(XINDT,1) =  RS1(3)
             ATABLA(XINDT,2) =  "SEC"
             ATABLA(XINDT,3) =  RS3(5)
             ATABLA(XINDT,4) =  SUPLEVL
             ATABLA(XINDT,5) =  RS1(4)

             STRSQL = "insert into SEGCKTRESP (SEGCKT,SEGCKTPSC,MODO,TIPO,CODIGO,ORDEN,NIVEL,SUBCODIGO) values "
             STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & RS3(5) & "','L','SEC','" & RS1(3) & "','" & XORDEN & "','" & SUPLEVL & "','" & RS1(4) & "')"
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)

       RS1.MOVENEXT
       LOOP

    RS3.MOVENEXT
    LOOP

    IF XINDT > 0 THEN

       ATABLA(0,0) =  XINDT
       ATABLA(0,1) =  XINDT - 2

    END IF

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



FUNCTION GENCKTPSC(XPSC,XSIG,XMODEL,XFIN,XRTIPO,XRCOD,XRASG,XPDSC,XOBSERV,XNIVEL,XSUBCOD)

    SEGCKTPSC = XPSC

    SESSION("SEGCKTPSC") = SEGCKTPSC

    XCOM = NOW()
    IF XFIN = "SI" THEN
       XFIN = NOW()
    ELSE
       XFIN = "NO"
    END IF

    strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC,OBSERVACIONES,NIVEL,SUBCODIGO) values "
    STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & SESSION("SEGCKTPSC") & "','" & XMODEL & "','" & NOW() & "','" & XFIN & "','" & XSIG & "','" & XRTIPO & "','" & XRCOD & "','" & XRASG & "','" & XPDSC & "','" & XOBSERV & "','" & XNIVEL & "','" & XSUBCOD & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

FUNCTION GENACCION(XPASO,XTPP)


    StrSql = "Select * From WKFPSCTPC WHERE WKFPASCOD = '" & XPASO & "' AND TIPO = '" & XTPP & "' "
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS1.EOF THEN
       IF RS1(4) = "A" THEN     
           strsql = "UPDATE REQUERIMIENTO SET ESTADO = " & RS1(5) & " WHERE codigo = '" & session("NCliente") & "' "
           call ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       END IF
    
    
    END IF
 



END FUNCTION






%>