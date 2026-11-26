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

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

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
    sImagen4sA = "<B><FONT FACE=WEBDINGS COLOR=#BBBBBB SIZE=2>&#0128;</FONT></B>"
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
  
    celdas = ofv.GenCelda("","TTT colspan=3","","","","","",sImagen & "&nbsp;REQUERIMIENTO","si")
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

     sinput = "<a  class=tt href=EDREQUER.asp?ACCION=S&OPCION=" & SOPCION
     sinput = sinput  & " ><b>&nbsp;CONFIRMA&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

     sinput = "<a  class=tt href=EDREQUER.asp"
     sinput = sinput  & " ><b>&nbsp;CANCELA&nbsp;</b></a>" 
     celdas = CELDAS & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

    celdas = CELDAS & ofv.GenCelda("","TTT","right","","","","","COR&nbsp;ON-TIME&nbsp;&nbsp;&nbsp;","si")
    filas = filas & ofv.GenRow("","aaa bgcolor='#f8f8f8' ","","","30","","",celdas,"si")
  sHTML = sHTML  & "<CENTER>" & ofv.GenTabla("","AAA STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;border-color:#888888;' ","","70%","","0","","3","0","","",filas,"si")  & "</CENTER>"

  END IF


  IF CARGAR = "S" THEN

  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 








    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","TTT colspan=3","","","","","",sImagen & "&nbsp;SOLICITUD DEL REQUERIMIENTO","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sImagen = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"

    celdas = ofv.GenCelda("","vt","","","","","","TITULO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","Usuarios","70","dt","")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","vt","","","","","","CODIGO:","si")
  celdas = celdas & ofv.GenCelda("","","","","","","","","si")


  if rs(19) <> "" then

        XTIPOREQ  = rs(19)

  else

        XTIPOREQ  = rs(1)

  end if

  celdas = celdas & ofv.GenCelda("","vt ","","","","","",XTIPOREQ,"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  celdas = ofv.GenCelda("","vt","","","","","","CLASE:","si")
  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

   XTREQ = 0

     StrSql = "Select * From TIPOREQ where TIPOREQ = '" & RS(15) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XTIPOREQ  = RS(15) & " - " & rsX2(2)
        XTREQ     = rsX2(0)

     ELSE
        XTIPOREQ  = "SIN DETERMINAR"

     END IF

  celdas = celdas & ofv.GenCelda("","ut","","","","","","<u>" & REPLACE(TRIM(XTIPOREQ)," ","&nbsp;") & "</u>","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","vt","","","","","","TIPO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      strsql = "Select * From TIPOS WHERE TIPOREQ = " & XTREQ & " Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(8).name,rs(8),strsql,cn,0,1,"")
    else
      strsql = "Select * From TIPOS where ID = " & rs(8)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if



  IF RS(11) = 1 THEN
        strsql = "Select * From TIPOS where ID = " & rs(8)
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then 
           sniv = rsniv(1)
           snivD = rsniv(2) & ""
           IF NOT ISNULL(snivD) THEN

       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' href='JAVASCRIPT:VOID(0);' ONCLICK=" & chr(34) & "document.all.icx1.style.display='none';" & chr(34)
       sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenin & "</b></a>" 


              sHTML = sHTML & "<DIV id='icx1' STYLE='display:block;z-index:0;margin-top:10;margin-left:10;background-color:#ffffff;position:absolute;left:110px;top:170px;width:400px;height:280px;border-style:ridge;border-color:#f7aa22;border-width:2;'>"
              sHTML = sHTML & "<SPAN><FONT FACE=VERDANA SIZE=1><b><P STYLE='border-style:outset;background-color:#eeeeee;' >" & sinput2 & "&nbsp;INFORMACION A COMPLETAR</P>"
              sHTML = sHTML & "&nbsp;&nbsp;" & REPLACE(snivD,CHR(13),"<BR>&nbsp;&nbsp;")


              sHTML = sHTML & "</b></FONT>"
              sHTML = sHTML & "</SPAN>"
              sHTML = sHTML & "</DIV>"

       sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' href='JAVASCRIPT:VOID(0);' ONCLICK=" & chr(34) & "document.all.icx1.style.display='block';" & chr(34)
       sinput2 = sinput2  & " TITLE='INFORMACION A COMPLETAR' ><b>" & sImagenin & "</b></a>" 


       sInput = sInput & "&nbsp;&nbsp;" & sinput2

           END IF


        END IF
        call ofv.cerrarconsulta(rsniv)



  END IF



    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","vt","","","","","","ORIGEN:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      strsql = "Select * From ORIGEN Order By DESCRI"
      sInput = ofv.GenerarCombo("Usuarios",rs(7).name,rs(7),strsql,cn3,0,1,"")
    else
      strsql = "Select * From TIPOS where ID = " & rs(7)
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","DESCRIPCION ORIGEN:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(12).name,rs(12),"text","Usuarios","70","dt","")
    else
      sInput = ofv.GenerarInput("",rs(12),"readonly","","70","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","AA COLSPAN=3 style='border-bottom-style:ridge;border-bottom-width:2;border-bottom-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","PRIORIDAD:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

     StrSql = "Select * From PRIORIDAD where ID = '" & RS(18) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XTIPOREQ  = rsX2(1)

     ELSE
        XTIPOREQ  = "SIN DETERMINAR"

     END IF
    celdas = celdas & ofv.GenCelda("","Vt","","","","","",sImagenPR & "&nbsp;" & XTIPOREQ,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","FECHA CREACION:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    if clng(rs(3)) = 0 then
      dFecha = "---"
    else
   ''   dFecha = cdate(rs(3) + clng(cdate("01-01-1900")))
      dFecha = Fechador(rs(3))
    end if

 ''     sInput = dFecha
      sInput = sImagenC & "&nbsp;" & ucase(dFecha)

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","vt","","","","","","FECHA SOLICITUD:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    if clng(rs(4)) = 0 then
      dFecha = "---"
    else
    ''  dFecha = cdate(rs(4) + clng(cdate("01-01-1900")))
      dFecha = Fechador(rs(4))
    end if

  ''    sInput = dFecha

      sInput = sImagenS & "&nbsp;" & ucase(dFecha)

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","FECHA ESPERADA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    if clng(rs(16)) = 0 then
      dFecha = "---"
    else
    ''  dFecha = cdate(rs(16) + clng(cdate("01-01-1900")))
      dFecha = "<B>" & Fechador(rs(16)) & "</B>"
    end if

  ''    sInput = dFecha

      sInput = sImagenES & "&nbsp;" & ucase(dFecha)

    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","vt","","","","","","FECHA PLANIFICADA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    if clng(rs(5)) = 0 then
      dFecha = "---"
    else
  ''    dFecha = cdate(rs(5) + clng(cdate("01-01-1900")))
      dFecha = Fechador(rs(5))
    end if

  ''    sInput = dFecha

      sInput = sImagenP & "&nbsp;" & ucase(dFecha)


    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","vt","","","","","","FECHA ENTREGA:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")

    if clng(rs(6)) = 0 then
      dFecha = "---"
    else
 ''     dFecha = cdate(rs(6) + clng(cdate("01-01-1900")))
      dFecha = Fechador(rs(6))
    end if

      sInput = sImagenE & "&nbsp;" & ucase(dFecha)


    celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","AA COLSPAN=3 style='border-bottom-style:ridge;border-bottom-width:2;border-bottom-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","vt","","","","","","ESTADO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

      sInstrumento = "---"

        Strsql = "Select descri From reqestados Where id = " & rs(11)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sImagen6ss & "&nbsp;" & sInstrumento


    celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","vt","","","","","","ORIGINANTE:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if RS(13) = "NO" then
      sInput = "SIN ORIGINANTE"
      celdas = celdas & ofv.GenCelda("","uT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagen4so & "&nbsp;<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","vt","","","","","","ANALISTA ASIGNADO:","si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if RS(14) = "NO" then
      sInput = sImagen4sA & "&nbsp;SIN ASIGNAR"
      celdas = celdas & ofv.GenCelda("","ut","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(14) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagen4si & "&nbsp;<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQUERIMIENTO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if



    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario0,"si")
    sHTML = sHTML & "</center>"

 
  IF session("ModMain2") = "CL" THEN
     sinput = "<a  class=tt href=Clientes.asp"
     sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ELSE
     sinput = "&nbsp;"
     celdas =  ofv.GenCelda("","","","1","","","",sInput,"si")
  END IF


  celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
  filas = ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")

  end if

 END IF




    sHTML = sHTML & "<form style='display:none;' name=init action='requerindex.asp' target='WK3' "
    sHTML = sHTML & "   method=post>"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

    sHTML = sHTML & "</center>"

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

  Agregar = Celt1
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