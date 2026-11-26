<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

''RESPONSE.WRITE SOPCION & " " & SACCION


 SELECT CASE SOPCION

        CASE "RELOAD"
          CARGAR = "S"
          SACCION = "S"

        CASE "AUTO"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "CONFIRMA INICIO DE CIRCUITO DE AUTORIZACION ???"
          ELSE
             CARGAR = "A"

             RESPONSE.REDIRECT "EDWKFSTART.ASP?OPCION=AUTO"

          END IF

        CASE "DELEGA"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "VA A DELEGAR EL REQUERIMIENTO A UN NUEVO RESPONSABLE ???"
          ELSE
             CARGAR = "A"

             RESPONSE.REDIRECT "EDDELEGA.ASP?OPCION=DELEGA"

          END IF

        CASE "CANDELEGA"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "VA A CANCELAR LA DELEGACION DEL REQUERIMIENTO A UN NUEVO RESPONSABLE ???"
          ELSE
             CARGAR = "A"

             RESPONSE.REDIRECT "EDDELEGA.ASP?OPCION=CANCEL"

          END IF

        CASE "CANCEL"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "CONFIRMA LA CANCELACION DEL REQUERIMIENTO ???"
          ELSE
             CARGAR = "S"

             strsql = "UPDATE REQUERIMIENTO SET ESTADO = 11, FECHAFIN = " & vNuevoValor & " WHERE codigo = '" & session("NCliente") & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '''            RESPONSE.WRITE STRSQL

          END IF

        CASE "TOMAR"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "CONFIRMA LA ASIGNACION DEL REQUERIMIENTO ???"
          ELSE
             CARGAR = "S"

             strsql = "UPDATE REQUERIMIENTO SET ESTADO = 4,ASIGNADO = '" & SESSION("USUARIO") & "' WHERE codigo = '" & session("NCliente") & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)

             strsql = "UPDATE REQSECTORES SET USRAFC = '" & SESSION("USUARIO") & "' WHERE codigo = '" & session("NCliente") & "' AND TIPO = 'AN' "
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)


          END IF

        CASE "FECHA"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             RESPONSE.REDIRECT "EDFECENT.ASP"
 
          END IF

        CASE "FECHATEN"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             RESPONSE.REDIRECT "EDFECTEN.ASP"
 
          END IF

        CASE "APRO"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "CONFIRMA INICIO DE CIRCUITO DE APROBACION ???"
          ELSE
             CARGAR = "A"

             RESPONSE.REDIRECT "EDWKFSTART.ASP?OPCION=APRO"

          END IF

        CASE "RECHAZAR"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "CONFIRMA EL RECHAZO DEL REQUERIMIENTO ???"
          ELSE
             CARGAR = "S"

             strsql = "UPDATE REQUERIMIENTO SET ESTADO = 7, FECHAFIN = " & vNuevoValor & " WHERE codigo = '" & session("NCliente") & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          END IF


        CASE "REACTIVAR"
          IF SACCION = "C" THEN   
             CARGAR = "N"
             TITULO = "CONFIRMA LA REACTIVACION REQUERIMIENTO ???"
          ELSE
             CARGAR = "S"

             strsql = "UPDATE REQUERIMIENTO SET ESTADO = 4, FECHAFIN = 0, FECHAENT = 0 WHERE codigo = '" & session("NCliente") & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)

          END IF


        CASE "SI"

             StrSql = "INSERT INTO REQINFO (CODIGO,TIPO,DESCRI,USUARIO,FECHA) VALUES ('" & session("NCliente") & "','1','DESCRIPCION','" & SESSION("USUARIO") & "','0') "
   ''' response.write strsql & "<br>" 

             CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

             CARGAR = "S"


        CASE ELSE 
             CARGAR = "S"



 END SELECT

 CARGARF  = CARGAR
 SACCIONF = SACCION


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



' FUNCION = " onload='uno();' "

sHTML = ""




   sHTML = ofv.FORMHeader(SESSION("FORMN"))


 IF CARGAR = "N" THEN

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
  
    celdas = ofv.GenCelda("","TTT colspan=3","","","","","",sImagen & "&nbsp;REQUERIMIENTO (version nueva)","si")
    filas  = ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","VT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#993333;background-color:#fbfbfb;' COLSPAN=3","CENTER","","50","","","<BIG><B>" & TITULO & "</B></BIG>","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     sinput = "<a  class=tt href=EDREQUERnv.asp?ACCION=S&OPCION=" & SOPCION
     sinput = sinput  & " ><b>&nbsp;CONFIRMA&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

     sinput = "<a  class=tt href=EDREQUERnv.asp"
     sinput = sinput  & " ><b>&nbsp;CANCELA&nbsp;</b></a>" 
     celdas = CELDAS & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

    celdas = CELDAS & ofv.GenCelda("","TTT","right","","","","","COR&nbsp;ON-TIME&nbsp;&nbsp;&nbsp;","si")
    filas = filas & ofv.GenRow("","aaa bgcolor='#f8f8f8' ","","","30","","",celdas,"si")
    sHTML = sHTML  & "<CENTER>" & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#AAAAAA;' ","","70%","","0","","3","0","","",filas,"si")  & "</CENTER>"

  END IF


  IF CARGAR = "S" THEN

  filas  = ""

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 
     xMODAL   = RS(24) & ""

     IF XMODAL = "" THEN XMODAL = "MP"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    celdas = ""

    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;'  ","","","","","","&nbsp;CODIGO:","si")
    if rs(19) <> "" then

        XTIPOREQ  = rs(19)

    else

        XTIPOREQ  = rs(1)

    end if
  
    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",XTIPOREQ,"si")


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;GENERADO:","si")

    if clng(rs(3)) = 0 then
      dFecha = "---"
    else
      dFecha = GTFechadorsh(rs(3))
    end if

    sInput = dFecha

    celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


   celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ENVIADO:","si")

     if clng(rs(4)) = 0 then
      dFecha = "---"
    else
      dFecha = GTFechadorsh(rs(4))
    end if
    sInput = dFecha

    celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ""



    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ORIGINANTE:","si")

    if RS(13) = "NO" then
       sInput = "SIN ORIGINANTE"

    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
        sInput = sniv
    end if

      celdas = celdas & ofv.GenCelda("","uT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;RESPONSABLE:","si")
    if RS(14) = "NO" then
      sInput = "SIN ASIGNAR"

    else
      strsql = "Select * From USUARIOS where ID = '" & rs(14) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sniv
    end if

      celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;ESTADO:","si")

    wkfestado = rs(11) & ""

      sInstrumento = "---"

        Strsql = "Select descri From GPestados Where id = " & rs(11)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento


    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

 IF (RS(24) & "") <> "MP" THEN


    celdas = ""



    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;MODO CARGA:","si")

   WKFMODOCARGA = (RS(24) & "")

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


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;USR.CARGA:","si")

      sInput = "SIN ASIGNAR"


      strsql = "Select * From USUARIOS where ID = '" & rs(25) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then
            sniv = rsniv(1)
            sInput = sniv
        END IF 
        call ofv.cerrarconsulta(rsniv)



      celdas = celdas & ofv.GenCelda("","ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")


    celdas = celdas & ofv.GenCelda("","Ut style='background-color:#336699;color:#ffffff;' ","","","","","","&nbsp;CONTACTO:","si")



    if bModificar then
      strsql = "Select * From GPMODOCONTACTO Order By DESCRI"
      sInput = ofv.GenerarCombo("UsuariosX",rs(26).name,rs(26),strsql,cn,0,1,"")
    else
      sInstrumento = "---"

        Strsql = "Select descri From GPMODOCONTACTO Where OBJid = 0" & rs(26)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento
    end if


    celdas = celdas & ofv.GenCelda("","Ut style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")








 END IF



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQUERIMIENTO"
      formularioX = ofv.GenForm("UsuariosX","",sAccion,"","post","",filas,"si")
    else
      formularioX = ofv.GenForm("","","","","post","",filas,"si")
    end if



    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#FFFFFF;BACKGROUND-color:#F0F0F0;' ","","100%","","0","","0","0","","",formularioX,"si")
    sHTML = sHTML & "</center>"







''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    filas = ""

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

    celdas = ofv.GenCelda("","VtI style='background-color:#888888;color:#ffffff;'  colspan=2","","","","","","SOLICITUD DEL REQUERIMIENTO Nro:&nbsp;" & rs(0) ,"si")
    filas  = filas & ofv.GenRow("","","","","25","","",celdas,"si")



    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;REQUERIMIENTO:","si")

     if wkfestado = "1" then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","Usuarios","70","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","utI bgcolor='#f7eedd' style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  IF XTC = "XX" THEN


    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;TIPO:","si")

   if wkfestado = "1" then
      strsql = "Select * From TIPOS WHERE TIPOREQ = " & XTREQ & " Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(8).name,rs(8),strsql,cn,0,1,"")
    else
      strsql = "Select * From TIPOS where ID = " & rs(8)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","utI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;DESCRIPCION ORIGEN:","si")
    if wkfestado = "1" then
      sInput = ofv.GenerarInput(rs(12).name,rs(12),"text","Usuarios","70","dt","")
    else
      sInput = ofv.GenerarInput("",rs(12),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","utI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


 END IF


    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;SISTEMA:","si")

    if bModificar then
      strsql = "Select * From GPSISTEMAS Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(21).name,rs(21),strsql,cn,1,2,"")
    else
      strsql = "Select * From GPSISTEMAS where CODSIST = '" & rs(21) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(2)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","UtI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;MODULO:","si")

    if bModificar then
      strsql = "Select * From GPMODULOS WHERE CODSIST = '" & rs(21) & "' Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(22).name,rs(22),strsql,cn,2,3,"")
    else
      strsql = "Select * From GPMODULOS WHERE CODMOD = '" & rs(22) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(3)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","UtI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;CASOS DE USO:","si")

   ' if bModificar then
   '   sInput = ofv.GenerarInput(rs(23).name,rs(23),"text","Usuarios","70","dt","")
   ' else
   '   sInput = ofv.GenerarInput("",rs(23),"readonly","","70","dt","")
   ' end if

    if bModificar then
      strsql = "Select * From GPCASOSUSO WHERE CODMOD = '" & rs(22) & "' Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(23).name,rs(23),strsql,cn,3,4,"")
    else
      strsql = "Select * From GPCASOSUSO WHERE CODUSO = '" & rs(23) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(4)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if



    celdas = celdas & ofv.GenCelda("","utI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","Ut style='background-color:#AAAAAA;color:#ffffff;' ","","","","","","&nbsp;PRIORIDAD:","si")

    if bModificar then
      strsql = "Select * From PRIORIDAD Order By ID"
      sInput = ofv.GenerarCombo("Usuarios",rs(18).name,rs(18),strsql,cn,0,1,"")
    else
      strsql = "Select * From PRIORIDAD where ID = '" & rs(18) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = celdas & ofv.GenCelda("","UtI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;'  ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





  celdas = ofv.GenCelda("","AA COLSPAN=2 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#bbbbbb;font-size:6pt;' ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    Celda = ofv.GenCelda("","VtI style='background-color:#888888;color:#ffffff;' colspan=2 ","","100%","","","","DESCRIPCION DEL PEDIDO","si")
    FilaS =  filaS & ofv.GenRow("","","","","25","","",Celda,"si")


  if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

  if not  isnull(rs(10)) then 
     wdescri = rs(10)
  end if 




    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;FONT-SIZE:7PT;' "
    t3b = t3b & "name=" & rs(10).name & " cols=140 rows=8 "
 if wkfestado = "1" then
      t3b = t3b & " onblur=Usuarios.submit() "
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
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQUERIMIENTO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if



    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#FFFFFF;' ","","100%","","0","","0","0","","",formulario0,"si")
    sHTML = sHTML & "</center>"




''EDWKFSTART.ASP?OPCION=AUTO


    sHTML = sHTML & "<form style='display:none;' name='reqadj' action='edarchivosnv.asp'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "<form style='display:none;' name='reqobs' action='edrequernv.asp?opcion=SI'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "<form style='display:none;' name='reqobst' action='edobservatodonv.asp?modal=TODOS'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

'''    sHTML = sHTML & "<form style='display:none;' name='reqwkf' action='EDWKFSTARTnv.ASP?OPCION=AUTO'  "
    sHTML = sHTML & "<form style='display:none;' name='reqwkf' action='COWKFREQUERnv.ASP?OPCION=INIT" & XMODAL & "'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "<form style='display:none;' name='reqhist' action='EDHISTWKFNV.asp'  "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"


  end if

 END IF



    celdas = ""
    filas  = ""
   reqmsg = "&nbsp;"

   GPTOTREQSA = 0
   GPTOTREQCA = 0

  StrSql = "Select count(*) From REQINFO where codigo = '" & session("NCliente") & "' AND NOT (ARCHIVO IS NULL) "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     if rs(0) > 0 then GPTOTREQCA = RS(0)

  end if


  StrSql = "Select count(*) From REQINFO where CODIGO = '" & session("NCliente") & "' AND (ARCHIVO IS NULL) "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     if rs(0) > 0 then GPTOTREQSA = RS(0)

  end if

  reqmsg = "El pedido tiene " & (GPTOTREQSA + GPTOTREQCA) & " observaciones, " & GPTOTREQCA & " archivos adjuntos."

  celdas = celdas & ofv.GenCelda("","vtI style='background-color:#F0F0F0;color:#993333;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","80%","","","",reqmsg,"si")


 if wkfestado = "1" then

   celdas = celdas & ofv.GenCelda("","Ut onclick='disp_confirm();' style='background-color:#F7AA22;color:#ffffff;CURSOR:HAND;border-style:OUTSET;border-width:2;' TITLE='Enviar pedido para su tratamiento.' ","CENTER","10%","","","","&nbsp;ENVIAR&nbsp;","si")

 ELSE

     StrSql = "Select * From SEGCIRCUITO WHERE  OBJCOD = '" & session("NCliente") & "'  "

     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then

       celdas = celdas & ofv.GenCelda("","Ut onclick='document.reqhist.submit();' style='background-color:#336699;color:#ffffff;CURSOR:HAND;' TITLE='Ver historial de circuitos.' ","CENTER","10%","","","","&nbsp;HISTORIAL&nbsp;","si")


     END IF


 end if

   filas = ofv.GenRow("","","","","20","","",celdas,"si")





    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#FFFFFF;' ","","100%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"


'''''''''''''''''''''''''''''''''''''''''''''''''''

  sPath = Request.ServerVariables("PATH_INFO")
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1))


  Xfer   =  Xruta
  rutdir = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "legajos/" & session("NCliente")
  nomdir = "Documentos-archivos-imagenes"



  rutdir = "legajos\" & session("NCliente")
  rutdir2 = "legajos/" & session("NCliente") & "/"


  sHTML = sHTML & "<center>"

  celdas = ofv.GenCelda("","VtI style='background-color:#888888;color:#ffffff;' colspan=2","","","","","","OBSERVACIONES","si")
  celdas = celdas & ofv.GenCelda("","Ut onclick='document.reqobs.submit();' style='border-style:outset;border-width:2;background-color:#993333;color:#ffffff;CURSOR:HAND;' TITLE='Agregar observaciones.' ","CENTER","10%","","","","&nbsp;AGREGAR&nbsp;","si")
  filas  = ofv.GenRow("","","","","25","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
'  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





  formulario0 = filas

  filas = ""



  StrSql = "Select * From REQINFO where CODIGO = '" & session("NCliente") & "' " & SWHERE & " ORDER BY USUARIO, ID DESC  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 

 

  else



    set cn2 = ofv.conectar(ofv.strconn0)

    DO WHILE NOT RS.EOF


    wkfecact = rs(5) & "" 


  IF (wkfecact & "") = "0" THEN

       vNuevoValor = clng(DATE()) - clng(cdate("01-01-1900"))

      sInput = ofv.GenerarInput(rs(5).name,vNuevoValor,"HIDDEN","Usuariosobs","1","dt","")

      nCol = 140

      if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

      if not  isnull(rs(3)) then 
         wdescri = rs(3)
      end if 






      t3b = "<font class=at><textarea class=dt STYLE='text-align:justify;background-color:transparent;padding-top:2px;FONT-SIZE:7pt;'' "
      t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=5 "

         t3b = t3b & " onkeypress=" & chr(34)
         t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
         t3b = t3b & " extenso'); return false;}" & chr(34) & " >"

      xdatos =  wdescri
      t3b = t3b & xdatos & "</textarea></font>" & sInput

  ELSE

      wdescri = replace(rs(3) & "",chr(13),"<br>")

      t3b = wdescri & ""

  END IF


  celdas = ofv.GenCelda("","UTi  colspan=3 STYLE='color:#000000;text-align:justify;border-style:solid;border-width:1;border-color:#CCCCCC;' ","","","","","",t3b,"si")
  filas  = ofv.GenRow("","","","","35","","",celdas,"si")


'''''''''''''''''''''''''''


  celdas = ""


    if clng(rs(5)) = 0 then
      dFecha = "---"
    else
    ''  dFecha = cdate(rs(5) + clng(cdate("01-01-1900")))
      dFecha = GTFechadorsh(rs(5)) & "."
    end if

  sInputx = "------"
       strsql = "Select * From USUARIOS where ID = '" & rs(6) & "' "
       set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rsniv.eof then sInputx = rsniv(1)
       call ofv.cerrarconsulta(rsniv)

  SINPUT = dFecha & "&nbsp;/&nbsp;" & sInputx


  SINPUTnv = " POR " & sInputx & " - " & dFecha & "&nbsp;"

 ' celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","","10%","","","",dFecha,"si")
 ' celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","","30%","","","",sInputx,"si")


  celdas = celdas & ofv.GenCelda("","UT colspan=2 style='font-size:6pt;' ","right","","","","",sInputnv,"si")






  IF (RS(7) & "") <> "" THEN

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  TARGET='_blank' href=" & chr(34) & rutdir2 & rs(7) & chr(34) &  " "
     sinput = sinput  & "  TITLE='CONSULTAR'  ><b>CONSULTAR</b></a>" 

     sinputk1 = sinput

     sinputk2 = "N"

  ELSE

     sinputk1 = "N"

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href='GPuploadn.asp?direct=" &  rutdir & "&nomdir=" & nomdir & "&GPITEM=" & rs(0) & "' "
     sinput = sinput  & "  TITLE='ASOCIAR'  ><b><FONT COLOR='#990000' >ASOCIAR</FONT></b></a>" 

     sinputk2 = sinput

  END IF




  if sinputk1 = "N" then
     celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","CENTER","","","","",sinputk2,"si")
  else
     celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","CENTER","","","","",sinputk1,"si")
  end if 

  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  IF (wkfecact & "") = "0" THEN

     celdas = ofv.GenCelda("","AA COLSPAN=3 bgcolor='#f0f0f0'  ","RIGHT","","","","","<input type='submit' value='ACTUALIZAR' class=bt >","si")
     filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  end if

     celdas = ofv.GenCelda("","AA COLSPAN=3 bgcolor='#DDDDDD'   ","CENTER","","","","","&nbsp;","si")
     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQINFO"
     formulario0 = formulario0 &  ofv.GenForm("Usuariosobs","",sAccion,"","post","",filas,"si")





  RS.MOVENEXT
  LOOP


  end if

    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;' ","","100%","","0","","2","2","","",formulario0,"si")
    sHTML = sHTML & "</center>"


'''''''''''''''''''''''''''''''''''''''''''''''''''''






    sHTML = sHTML & "<form style='display:none;' name=init action='requerindexnv.asp' target='WK3' "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "</center>"



    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm()"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm('Confirma el envio del pedido?');"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    document.reqwkf.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Cancelo el envio!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"



 '' RESPONSE.WRITE CARGARF & " " & SACCIONF


' IF CARGARF = "S" AND SACCIONF = "S" THEN

'    shtml = shtml & "<script  language=javascript>" & chr(13)
'    shtml = shtml & " document.init.submit(); " & chr(13)
'    shtml = shtml & " </script>" & chr(13)

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

Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



%>