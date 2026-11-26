<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sHTML = "<HTML><BODY>"

  sHTML = sHTML & "<center>"

    Celda = ofv.GenCelda("","ttr colspan=2","","100%","","","","PANEL DE CONTROL DEL REQUERIMIENTO","si")
 '   Fila = ofv.GenRow("","","","","","","",Celda,"si")

    set cn = ofv.conectar(ofv.strconn6)

   StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
   set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if NOT rs.eof then
      estado  = RS(11) 
      origin  = RS(13) 
      asigna  = RS(14) 
      DELEGA  = RS(17) 
      TIPOREQ = RS(15) 

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","vt colspan=2","","100%","","","",sLink,"si")
'    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

   '' sImagen = ofv.GenImage("","aa style='border-style:outset;border-color:#336699;' ","","","imagenes/botrep.jpg","","")

    sImagen = "<B><FONT FACE=WebDINGS COLOR=#336699 SIZE=2 style='border-width:1;border-style:outset;' >4</FONT></B>"

    sHref = "EDREQUERNV.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","SOLICITUD","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    sHref = "EDALCANCENV.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","ALCANCE","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    sHref = "EDSECTORESNV.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","GERENCIAS/SECTORES AFECTADOS","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    StrSql = "Select * From REQDELEGA WHERE CODIGO = '" & session("NCliente") & "' "
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsX.eof THEN


    sHref = "EDVIEWDELEGANV.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","DELEGACIONES DEL REQUERIMIENTO","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    END IF
   call ofv.cerrarconsulta(RSX)

    if ESTADO > 1 THEN


    sHref = "EDHISTWKFNV.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","HISTORIAL DE CIRCUITOS","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    END IF


    sHref = "EDARCHIVOS.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","ARCHIVOS ASOCIADOS","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    sHref = "EDCORMANANGER.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","DOCUMENTOS COREMANAGER RELACIONADOS/AFECTADOS","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
  '  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    sHref = "EDCONTACTOS.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","CONTACTOS","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    sHref = "EDresumen.asp"
    sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
    Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","nt",sHref,"","","","RESUMEN","si")
    Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","vt colspan=2","","100%","","","",sLink,"si")
  '  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","aaa bgcolor='#f7f7fe' rules=rows style='border-style:SOLID;border-width:1;border-color:#888888;' ","","90%","","0","","-2","1","","",fila,"si")

  sHTML = sHTML & "<hr width=100% size=1 color=#993333 >"

  '  sImagen = ofv.GenImage("","aa style='border-style:outset;border-color:#F7AA22;' ","","","imagenes/botrep.jpg","","")
  '  sImagen2 = ofv.GenImage("","aa style='border-style:outset;border-color:#ff0000;' ","","","imagenes/botrep.jpg","","")


    sImagen = "<B><FONT FACE='WebDINGS' COLOR='#F7AA22' SIZE=2 style='border-width:1;border-style:outset;' >4</FONT></B>"
    sImagen2 = "<B><FONT FACE='WebDINGS' COLOR='#ff0000' SIZE=2 style='border-width:1;border-style:outset;' >4</FONT></B>"

    Fila = ""



    Celda = ofv.GenCelda("","VT colspan=2 bgcolor='#DDDDDD' style='border-width:1;border-style:outset;' ","","","","","","ACCIONES DEL REQUERIMIENTO","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


      SELECT CASE ESTADO

             CASE 1

                IF SESSION("USUARIO") = ORIGIN THEN


                   sHref = "EDREQUER.asp?OPCION=FECHATEN&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","ESTABLECER FECHA ESPERADA DE ENTREGA","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                   sHref = "EDREQUER.asp?OPCION=AUTO&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","INICIAR CIRCUITO AUTORIZACION","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


                   sHref = "EDREQUER.asp?OPCION=CANCEL&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","CANCELAR REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF

             CASE 3



                StrSql = "Select DESTSECTOR,DESTGERENCIA,DESTNIVEL From TIPOREQ where TIPOREQ =  '" & TIPOREQ & "' "
                set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsX.eof then
                   SECTORAN = RSX(0)
                   GERENCAN = RSX(1)
                   NIVELCAN = RSX(2)
                ELSE

                   SECTORAN = 0
                   GERENCAN = 0
                   NIVELCAN = 0
                END IF
                call ofv.cerrarconsulta(RSX)

                set cn2 = ofv.conectar(ofv.strconn0)

                sniv = 0

                strsql = "Select SECTOR From USUARIOS where ID = '" & SESSION("USUARIO") & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(0)
                call ofv.cerrarconsulta(rsniv)

              '''  response.write sectoran & " " & gerencan & " " & session("segnivel") & " " & sniv & " " & nivelcan

                IF (SECTORAN = sniv OR GERENCAN = SNIV) AND (clng(NIVELCAN) >= clng(SESSION("SEGNIVEL"))) THEN



                   sHref = "EDREQUER.asp?OPCION=TOMAR&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","ASIGNAR REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF



             CASE 4

                IF SESSION("USUARIO") = ORIGIN THEN


                   sHref = "EDREQUER.asp?OPCION=FECHATEN&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","ESTABLECER FECHA ESPERADA DE ENTREGA","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF

                IF SESSION("USUARIO") = ASIGNA THEN


                   sHref = "EDREQUER.asp?OPCION=FECHA&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","ESTABLECER FECHA PLANIFICADA DE ENTREGA","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


                   XREQTSK = "N"
                   XREQPRY = "N"

                   sHref = "PROYREQUER.asp?opcion=CP"
                   sLink = ofv.GenLink("","vt",sHref,"_top","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                StrSql = "Select * From PROYECTO where codigo = '" & session("NCliente") & "' "
                set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsX.eof then
                   XREQPRY = "S"
                   sLink = ofv.GenLink("","nt",sHref,"_top","","","ACCEDER AL PROYECTO ASOCIADO","si")
                ELSE

                   sLink = ofv.GenLink("","nt",sHref,"_top","","","ESTABLECER PROYECTO ASOCIADO","si")
                END IF
                call ofv.cerrarconsulta(RSX)
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   FilaP = ofv.GenRow("","","","","","","",Celda,"si")

                   sHref = "PROYREQUER.asp?opcion=CT"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                StrSql = "Select * From PROYTAREAS where PROYTSKOBJ = '" & session("NCliente") & "' "
                set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if NOT rsX.eof then
                   XREQTSK = "S"
                   sLink = ofv.GenLink("","nt",sHref,"","","","ACCEDER A LA TAREA DEL PROYECTO ASOCIADA","si")
                ELSE

                   sLink = ofv.GenLink("","nt",sHref,"","","","ESTABLECER LA TAREA DEL PROYECTO ASOCIADA","si")
                END IF
                call ofv.cerrarconsulta(RSX)
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   FilaT = ofv.GenRow("","","","","","","",Celda,"si")


                   if XREQTSK = "N" and XREQPRY = "N" then
                      Fila = Fila & FilaP & FilaT
                   elseif XREQTSK ="S" then
                      Fila = Fila & FilaT
                   elseif XREQPRY ="S" then
                      Fila = Fila & FilaP
                   end if

                   sHref = "EDREQUER.asp?OPCION=DELEGA&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","DELEGAR EL REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


                   sHref = "EDREQUER.asp?OPCION=APRO&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","INICIAR CIRCUITO APROBACION","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


                   sHref = "EDREQUER.asp?OPCION=RECHAZAR&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","RECHAZAR REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



                END IF


             CASE 8

                IF SESSION("USUARIO") = ORIGIN THEN


                   sHref = "EDREQUER.asp?OPCION=AUTO&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","REINICIAR CIRCUITO APROBACION","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



                   sHref = "EDREQUER.asp?OPCION=CANCEL&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","CANCELAR REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF

             CASE 9

                IF SESSION("USUARIO") = ASIGNA THEN



                   sHref = "EDREQUER.asp?OPCION=REACTIVAR&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","REACTIVAR REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF

             CASE 14

                IF SESSION("USUARIO") = ASIGNA THEN


                   sHref = "EDREQUER.asp?OPCION=TOMAR&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","ASIGNAR REQUERIMIENTO","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF

                IF SESSION("USUARIO") = DELEGA THEN


                   sHref = "EDREQUER.asp?OPCION=CANDELEGA&ACCION=C"
                   sLink = ofv.GenLink("","vt",sHref,"","","",sImagen2,"si")
                   Celda = ofv.GenCelda("","vt","","1%","","","",sLink,"si")
                   sLink = ofv.GenLink("","nt",sHref,"","","","CANCELAR DELEGACION","si")
                   Celda = Celda & ofv.GenCelda("","ut","","99%","","","",sLink,"si")
                   Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

                END IF


          END SELECT

      end if


  sHTML = sHTML & ofv.GenTabla("","aaa bgcolor='#f7f7fe' rules=rows style='border-style:SOLID;border-width:1;border-color:#888888;' ","","90%","","0","","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"



  sHTML = sHTML & "</body></html>"
end if

  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)

Response.Write sHTML

%>