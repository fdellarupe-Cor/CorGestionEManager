<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.formHeader("")



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
 act = request.querystring("act")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

''RESPONSE.WRITE SOPCION & " " & SACCION


 SELECT CASE SOPCION

        CASE "DELEGA"

        IF SACCION = "AFC" THEN


               WDSEC = REQUEST.querystring("WDSEC")
               WDGER = 0
               WDUSR = REQUEST.querystring("WDUSR")

        ELSE
               WDSEC = 0
               WDGER = 0
               WDUSR = "NO"

        END IF 



        CASE "OKDELEGA"

            StrSql = "Select * From REQDELEGA WHERE CODIGO = '" & session("NCliente") & "' ORDER BY ORDEN DESC "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rs.eof THEN
               WORD = RS(2) + 10

            ELSE

               WORD = 10 


            END IF

               WASEC = 0
               WAGER = 0
               WAUSR = "NO"

               WDSEC = REQUEST.FORM("WDSEC")
               WDGER = REQUEST.FORM("WDGER")
               WDUSR = REQUEST.FORM("WDUSR")
               WDENV = REQUEST.FORM("WDENV")
               WDMSG = REQUEST.FORM("WDMSG")


            StrSql = "Select * From REQUERIMIENTO WHERE CODIGO = '" & session("NCliente") & "' "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rs.eof THEN
               WAUSR = RS(14)
               strsql = "UPDATE REQUERIMIENTO SET ESTADO = 14, ASIGNADO = '" & WDUSR & "', DELEGANTE = '" & RS(14) & "' WHERE ID = " & RS(0)
               call ofv.crearconsultaEx(StrSql,cn,1,parametros)
            END IF
            call ofv.cerrarconsulta(RS)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SI ESTA PARA ENVIAR MAIL ==> MANDAR MAIL DE DELEGACION

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrSql = "Select * From REQSECTORES WHERE TIPO = 'AN' AND CODIGO = '" & session("NCliente") & "' "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rs.eof THEN

               WASEC = RS(3)
               WAGER = RS(4)


               strsql = "UPDATE REQSECTORES SET SECTOR = '" & WDSEC & "' "
               strsql = STRSQL & ", GERENCIA = '" & WDGER & "' "
               strsql = STRSQL & ", USRAFC = '" & WDUSR & "' "
               strsql = STRSQL & " WHERE ID = " & RS(0)
               call ofv.crearconsultaEx(StrSql,cn,1,parametros)

            END IF
            call ofv.cerrarconsulta(RS)

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini



            STRSQL = "INSERT INTO REQDELEGA (CODIGO,ORDEN,SECTOR,GERENCIA,TIPO,SUBTIPO,USRAFC,"
            STRSQL = STRSQL & "WMSG,ENVIA,DSECTOR,DGERENCIA,DUSRAFC,ESTADO,FECHA,HORA) VALUES ("  
            STRSQL = STRSQL & "'" & session("NCliente") & "','" & WORD & "','" & WASEC & "','" & WAGER & "','DL','','" & WAUSR & "'"
            STRSQL = STRSQL & ",'" & WDMSG & "','" & WDENV & "','" & WDSEC & "','" & WDGER & "','" & WDUSR & "','0','" & vNuevoValor & "','" & NOW() & "')" 
            CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)





            
            RESPONSE.REDIRECT "EDREQUER.asp?OPCION=RELOAD"



        CASE "CANCEL"
            StrSql = "Select * From REQDELEGA WHERE CODIGO = '" & session("NCliente") & "' ORDER BY ORDEN DESC "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rs.eof THEN

''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SI ESTA PARA ENVIAR MAIL ==> MANDAR MAIL DE CANCELACION

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''

               strsql = "UPDATE REQDELEGA SET ESTADO = 1 WHERE ID = " & RS(0)
               call ofv.crearconsultaEx(StrSql,cn,1,parametros)

               strsql = "UPDATE REQSECTORES SET SECTOR = '" & RS(3) & "' "
               strsql = STRSQL & ", GERENCIA = '" & RS(4) & "' "
               strsql = STRSQL & ", USRAFC = '" & RS(7) & "' "
               strsql = STRSQL & " WHERE TIPO = 'AN' AND CODIGO = '" & session("NCliente") & "' "
               call ofv.crearconsultaEx(StrSql,cn,1,parametros)



            END IF



            StrSql = "Select * From REQUERIMIENTO WHERE CODIGO = '" & session("NCliente") & "' "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rs.eof THEN
               strsql = "UPDATE REQUERIMIENTO SET ESTADO = 4, ASIGNADO = '" & RS(17) & "', DELEGANTE = 'NO' WHERE ID = " & RS(0)
               call ofv.crearconsultaEx(StrSql,cn,1,parametros)
            END IF
            call ofv.cerrarconsulta(RS)



            
            RESPONSE.REDIRECT "EDREQUER.asp?OPCION=RELOAD"

 END SELECT



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;DELEGACION DEL REQUERIMIENTO","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ""

  celdas = celdas & ofv.GenCelda("","ttr","","4%","","","","USUARIO","si")
  celdas = celdas & ofv.GenCelda("","ttr","","32%","","","","NOMBRE","si")
  celdas = celdas & ofv.GenCelda("","ttr","","32%","","","","AREA/SECTOR","si")
  celdas = celdas & ofv.GenCelda("","ttr","","32%","","","","GERENCIA","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    sInput = "REQUERIMIENTO: "
    celdas = ofv.GenCelda("","VT","","10%","","","",sInput,"si")

    sInput = "SIN ASIGNAR"



  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sInput = "<B>" & RS(2) & "</B>"
    WAUSR = RS(14)
  END IF


    celdas = celdas & ofv.GenCelda("","VT COLSPAN=2 ","","","","","",sInput,"si")

 ''''   filas = filas & ofv.GenRow("","","","","","","",celdas,"si")




    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","vt colspan=4","","100%","","","",sLink,"si")
'    filas = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sinput = "RESPONSABLE DELEGANTE"
    celdas = ofv.GenCelda("","VT COLSPAN=4 ","CENTER","","","","",sInput,"si")

 '   filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    celdas = ""

    sinput = "DELEGANTE"
    celdas = ofv.GenCelda("","UT style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#F7AA22;background-color:#f9f9f9;' ","","","","","",sInput,"si")

    StrSql = "Select * From REQSECTORES WHERE TIPO = 'AN' AND CODIGO = '" & session("NCliente") & "' "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof THEN

    if WAUSR = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","uT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & WAUSR & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput =  sniv 
      celdas = celdas & ofv.GenCelda("","Ut","","","","","",sInput,"si")
    end if


    if RS(3) = 0 then
      celdas = celdas & ofv.GenCelda("","Ut","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(3)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
        Session("Sletra") = RS(3)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","Ut","","","","","",sInput,"si")


    end if

    if RS(4) = 0 then
      celdas = CELDAS & ofv.GenCelda("","Ut","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(4)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = CELDAS & ofv.GenCelda("","Ut","","","","","",sInput,"si")


    end if






      filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


   END IF
   call ofv.cerrarconsulta(RS)

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","vt colspan=4","","100%","","","",sLink,"si")
 '''   filas = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=EDDELEGABUS.asp?Sopcion=Y"
     sinput = sinput  & "  TITLE='Afectar/Modificar'  ><b>" & sImagenI3 & "&nbsp;(Afectar/Modificar)</b></a>" 

 '   sinput = "NUEVO RESPONSABLE&nbsp;&nbsp;" & sinput
 '   celdas = ofv.GenCelda("","VT COLSPAN=4 ","CENTER","","","","",sInput,"si")

 '''   filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

    sLink = "<hr width=100% size=1 color=#888888 >"
    Celda = ofv.GenCelda("","vt colspan=4","","100%","","","",sLink,"si")
    filas = filas & ofv.GenRow("","","","","","","",Celda,"si")


celdas = ""

    sinput = "DELEGADO"
    celdas = ofv.GenCelda("","UT style='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#008000;background-color:#f9f9f9;' ","","","","","",sInput,"si")


    if WDUSR = "NO" then
      sInput = "SIN RESPONSABLE" & ofv.GenerarInput("WDUSR",WDUSR,"HIDDEN","Usuarios","70","dt","")
      celdas = celdas & ofv.GenCelda("","uT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & WDUSR & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput =  sniv  & ofv.GenerarInput("WDUSR",WDUSR,"HIDDEN","Usuarios","70","dt","")
      celdas = celdas & ofv.GenCelda("","Ut","","","","","",sInput,"si")
    end if


    if WDSEC = 0 then
      celdas = celdas & ofv.GenCelda("","Ut","","","","","","SIN AFECTAR" & ofv.GenerarInput("WDSEC",WDSEC,"HIDDEN","Usuarios","70","dt",""),"si")

    else

     StrSql = "Select * From ECO where ECO = " & WDSEC
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","Ut","","","","","",sInput & ofv.GenerarInput("WDSEC",WDSEC,"HIDDEN","Usuarios","70","dt",""),"si")


    end if

    if WDSEC = 0 then
      celdas = CELDAS & ofv.GenCelda("","Ut","","","","","","SIN AFECTAR" & ofv.GenerarInput("WDGER",WDGER,"HIDDEN","Usuarios","70","dt",""),"si")

    else



     WDGER = GENSECTORAF(WDSEC)

    if WDGER = 0 then
      celdas = CELDAS & ofv.GenCelda("","Ut","","","","","","SIN AFECTAR" & ofv.GenerarInput("WDGER",WDGER,"HIDDEN","Usuarios","70","dt",""),"si")

    else

     StrSql = "Select * From ECO where ECO = " & WDGER
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = CELDAS & ofv.GenCelda("","Ut","","","","","",sInput & ofv.GenerarInput("WDGER",WDGER,"HIDDEN","Usuarios","70","dt",""),"si")

    END IF

    end if






      filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


    sLink = "<hr width=100% size=1 color=#993333 ><BR>"
    Celda = ofv.GenCelda("","vt colspan=4","","100%","","","",sLink,"si")
    filas = filas & ofv.GenRow("","","","","","","",Celda,"si")


		    Celda = ofv.GenCelda("","TTR colspan=4","center","100%","","","","MENSAJE MAIL" & ofv.GenerarInput("WDENV","0","HIDDEN","Usuarios","70","dt",""),"si")
		    FilaS = filas & Ofv.GenRow("","","","","","","",Celda,"si")

		    nCol = 80

		    wdescri = "SIN OBSERVACIONES" 

		    t3b = "<font class=at><textarea class=dt "
		    t3b = t3b & "name=WDMSG cols=" & nCol & " rows=8 "
		    if bModificar then
		 '''     t3b = t3b & " onblur=Usuarios.submit() "
		      t3b = t3b & " onkeypress=" & chr(34)
		      t3b = t3b & " if (this.value.length > 950) {alert('El texto es demasiado"
		      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
		    else
		      t3b = t3b & " readonly >"
		    end if
		    xdatos =  wdescri
		    t3b = t3b & xdatos & "</textarea></font>"
		    Celda = ofv.GenCelda("","abc colspan=4","center","100%","","","",t3b,"si")
		    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


		    if bModificar then
		      sAccion = "EDDELEGA.ASP?OPCION=OKDELEGA"
		      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
		    else
		      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
		    end if


                    sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario0,"si") & "<BR>"







		    Celda = ""



		  sinput = ""

		  sinput = sinput  & "<form name='YY' style='display:none;'  action='NONE'"
		  sinput = sinput  & " method=post>"
		  sinput = sinput  & "<input  type='BUTTON' size=10 name=botonOK value=' Aceptar ' "
		  sinput = sinput  & " onclick='Usuarios.submit();return false;' "
		  sinput = sinput  & " style='background-color:#00CC00;color:#ffffff;' >"
		  sinput = sinput  & "</form>"

		    Celda = celda & ofv.GenCelda("","UT","","","","","",sinput,"si")




		  sinput = ""

		  sinput = sinput  & "<form name='YX' style='display:BLOCK;' action='edrequer.asp'"
		  sinput = sinput  & " method=post>"
		  sinput = sinput  & "<input type='SUBMIT' size=10 name=botonrch value=' Cancelar ' "
		  sinput = sinput  & "  "
		  sinput = sinput  & " style='background-color:#cc0000;color:#ffffff;' >"
		  sinput = sinput  & "</input>" & chr(13)
		  sinput = sinput  & "</form>"


		    Celda = celda & ofv.GenCelda("","Ut","","","","","",sinput,"si")


		    FilaS =  ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","10%","","0","","0","0","","",filaS,"si")
    sHTML = sHTML & "</center><br><br>"

    sHTML = sHTML & "<SCRIPT LANGUAGE='JAVASCRIPT' >"
    sHTML = sHTML & " if (document.all.Usuarios) { "
    sHTML = sHTML & "    if (document.all.WDSEC.value>=" & chr(34) & "1" & chr(34) & ") { "
    sHTML = sHTML & "        document.all.YY.style.display=" & chr(34) & "block" & chr(34) & "} "
    sHTML = sHTML & "                            } "
    sHTML = sHTML & "</SCRIPT>"



  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1

  Buscar = ""
  if bAgregar then
     sAccion = "EDDELEGABUS.asp?Sopcion=Y"
     Buscar = ofv.BotonAgregar(sAccion)
  end if

  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



FUNCTION GENSECTORAF(XSECTOR)




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



GENSECTORAF = XSECTGEREN









END FUNCTION



%>