<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

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
  sopc    = request.querystring("opc")
  sTIPO   = request.querystring("TIPO")
  sRESP   = request.querystring("RESP")

 XMOD = request.querystring("XMOD")
 XDST = request.querystring("XDST")
 XDET = request.querystring("XDET")
 XSET = request.querystring("XSET")
 XOPT = request.querystring("XOPT")

 DIM ATABLA(50,3) 

 ITAB=0
 TTAB=0 
 ITABB=0



 LISTA = ""
 DESTINO = ""

 IF XMOD    = "PRY" THEN
    STABLA  = "PROYECTO"
    STABLA2 = "PROYRESP"

 ELSE
    STABLA  = "REQUERIMIENTO"
    STABLA2 = "REQSECTORES"


 END IF

             LISTA = " LISTA DE  " & STABLA & "S "


SELECT CASE XOPT


 CASE "XGLB"

       XDEST="GLOBAL"

  Strsql = "Select * from clienteplan where id = 1"
  set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
  if not rs.eof then
     NOMEMP = rs(1)
     LOGEMP = rs(5)
  end if
  call ofv.cerrarconsulta(rs)

       XDESTN=NOMEMP

             REQWHERE = ""
             DESTINO = " GLOBAL " 


  CASE "XGRN"

       XDEST="GERENCIA"

     StrSql = "Select * From ECO where ECO = " & SRESP
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XDESTN=sniv

       REQWHERE = " INNER JOIN " & STABLA2 & " ON " & STABLA & ".CODIGO = " & STABLA2 & ".CODIGO AND "

       IF XDST = "ORG" THEN 

             REQWHERE = REQWHERE & STABLA2 & ".TIPO = 'OR' "
             DESTINO = " ORIGINADOS POR " & XDESTN

       ELSE

             REQWHERE = REQWHERE & STABLA2 & ".TIPO = 'AN' "
             DESTINO = " ASIGNADOS A " & XDESTN

       END IF

             REQWHERE = REQWHERE & " AND " & STABLA2 & ".GERENCIA = " & SRESP & " "


  CASE "XSCT"

       XDEST="SECTOR"

     StrSql = "Select * From ECO where ECO = " & SRESP
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XDESTN=sniv

       REQWHERE = " INNER JOIN " & STABLA2 & " ON " & STABLA & ".CODIGO = " & STABLA2 & ".CODIGO AND "

       IF XDST = "ORG" THEN 

             REQWHERE = REQWHERE & STABLA2 & ".TIPO = 'OR' "
             DESTINO = " ORIGINADOS POR " & XDESTN

       ELSE

             REQWHERE = REQWHERE & STABLA2 & ".TIPO = 'AN' "
             DESTINO = " ASIGNADOS A " & XDESTN

       END IF

             REQWHERE = REQWHERE & " AND " & STABLA2 & ".SECTOR = " & SRESP & " "

  CASE "XRESP"

       XDEST="RESPONSABLE"

      strsql = "Select * From USUARIOS where ID = '" & SRESP & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsniv)


       XDESTN=sniv

       REQWHERE = " INNER JOIN " & STABLA2 & " ON " & STABLA & ".CODIGO = " & STABLA2 & ".CODIGO AND "

       IF XDST = "ORG" THEN 

             REQWHERE = REQWHERE & STABLA2 & ".TIPO = 'OR' "
             DESTINO = " ORIGINADOS POR " & XDESTN

       ELSE

             REQWHERE = REQWHERE & STABLA2 & ".TIPO = 'AN' "
             DESTINO = " ASIGNADOS A " & XDESTN

       END IF

             REQWHERE = REQWHERE & " AND " & STABLA2 & ".USRAFC = '" & SRESP & "' "

END SELECT


     StrSql = "Select " & STABLA & ".* From " & STABLA

     if reqwhere = "" then
        reqwhere = " WHERE "

     ELSE
        reqwhere = reqwhere & " AND "
     END IF

     REQWHERE2 =  STABLA & ".ESTADO > 3 ORDER BY " & STABLA & ".ESTADO"
     StrSql = STRSQL & REQWHERE & REQWHERE2




    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenX  = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>4</FONT></B>"
    sImagenIX = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=3>2</FONT></B>"









    celdas = ""



  'Encabezados
  tit = ofv.GenCelda("","ttr colspan=6","center","","","","",LISTA & DESTINO,"si")
  formulario = formulario & ofv.GenRow("","aaa height=30  ","","","","","",tit,"si")





   ''  StrSql = "Select * From " & STABLA & " WHERE ESTADO = " & ATABLA(ITABB,1) & REQWHERE
   ''  response.write strsql & "<br>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        rs.move(CantRegMoverx)
        formulario = formulario & GENLISTA(RS)

      celdas = ""

      sInput = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","vt colspan=2","","","","","",sInput,"si")

      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")


      sInput = "<b>&nbsp;Fecha Creacion</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      if xmod = "PRY" then
         sInput = "<b>&nbsp;Fecha Inicio</b>" 
      else
         sInput = "<b>&nbsp;Fecha Solicitud</b>"
      end if
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      formulario = formulario & ofv.GenRow("","aaa height=30  ","","","","","",celdas,"si")

      celdas = ""

      sInput = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","vt colspan=2","","","","","",sInput,"si")

      sInput = "<b>" & sImagenO & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      sInput = "<b>&nbsp;Usuario Originante</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      sInput = "<b>" & sImagenI & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")

      sInput = "<b>&nbsp;Usuario Interviniente</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","","","","",sInput,"si")

      formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     ELSE

        celdas = ofv.GenCelda("","VT colspan=6","center","","","","","NO HAY " & STABLA & "S PARA LA SELECCION ","si")
        formulario = formulario & ofv.GenRow("","aaa  ","","","","","",celdas,"si")

     END IF





  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""


  Buscar = ""
  

  
 ' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML







FUNCTION GENLISTA(RSX)


 XESTANT = 0

      ATABLA(0,0) = "TOTAL GENERAL"
      ATABLA(0,1) = 0 
      ATABLA(0,2) = 0 
      ATABLA(0,3) = 0 


  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

  XVA = "S"

  IF SESSION("SELOPC") = "1" THEN


              XTIPO = RSX(8)


           StrSql = "Select * From TIPOS where ID = 0" & XTIPO
           set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not  rsX2.eof then

              IF RSX2(4) = 0 THEN
                 XVA = "N"

              END IF

           END IF




  END IF 
  


   IF XVA = "S" THEN

  IF SESSION("FECOPC") = "2" THEN
     XFECDES = RSX(3)
     IF RSX(4) > 0 THEN XFECDES = RSX(4)

     XFECHAS = RSX(5)

     IF RSX(6) > 0 THEN XFECHAS = RSX(6)

     IF XFECHAS = 0 THEN XFECHAS = XFECDES

      
     IF CLNG(SESSION("FECDES")) > CLNG(XFECHAS) OR CLNG(SESSION("FECHAS")) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


  END IF 

 ' RESPONSE.WRITE RSX(1) & " " & RSX(2) & " " & XVA & " " & SESSION("FECDES") & " " & XFECDES  & " " & SESSION("FECHAS")  & " " & XFECHAS  & "<BR>"

  IF XVA = "S" THEN

  IF XESTANT <> RSX(11) THEN



      sniv = "---"

                strsql = "Select * From REQESTADOS where ID = '" & RSX(11) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sinput = "<b>ESTADO: " & sniv & "</b>" 
      celdas = ofv.GenCelda("xx1","vt COLSPAN=6 bgcolor='#f4f4f4' style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#6699cc;' VALIGN='BOTTOM' ","","1%","30","","",sInput,"si")
      IF XDET = "1" THEN formu = formu & ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")

      XESTANT = RSX(11)
      ITAB=ITAB + 1
      TTAB=ITAB 
      ATABLA(ITAB,0) = sniv 
      ATABLA(ITAB,1) = 0 
      ATABLA(ITAB,2) = 0 
      ATABLA(ITAB,3) = 0 

  END IF


    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "vt"

      ATABLA(ITAB,1) = ATABLA(ITAB,1) + 1
      ATABLA(0,1)    = ATABLA(0,1) + 1

    IF XMOD = "PRY" THEN

       xlink = "<A  href='&nbsp;'  "
       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(1) & "&tipo=PRY'); return false;" & chr(34)
       xlink = xlink & " ><IMG  src='../CORESS20/Imagenes/ICorONTIME.jpg'  border=0 ></A>"

       xlink2 = "<b>" & rsX(2) & "</b>" 

    ELSE

      sniv = "NO"

                strsql = "Select * From PROYECTO where CODIGO = '" & RSX(1) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then sniv = "SI"
                call ofv.cerrarconsulta(rsniv)

      IF sniv = "NO" THEN

                strsql = "Select * From PROYTAREAS where PROYTSKOBJ = '" & RSX(1) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsniv.eof then sniv = "SI"
                call ofv.cerrarconsulta(rsniv)

      END IF

      IF sniv = "SI" THEN
         ATABLA(ITAB,2) = ATABLA(ITAB,2) + 1 
         ATABLA(0,2)    = ATABLA(0,2) + 1 
      ELSE
         ATABLA(ITAB,3) = ATABLA(ITAB,3) + 1 
         ATABLA(0,3)    = ATABLA(0,3) + 1 
      END IF 
 

       xlink = "<A  href='&nbsp;'  "
       xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(1) & "&tipo=REQ'); return false;" & chr(34)
       xlink = xlink & " ><IMG  src='Imagenes/ICorAbierto.jpg'  border=0 ></A>"

       xlink2 = "<b>" & rsX(2) & "</b>" 

    END IF

      sInput = xlink

      celdas = ofv.GenCelda("","vt","","1%","","","",sInput,"si")



      sInput = xlink2

      celdas = celdas & ofv.GenCelda("","vt","","30%","","","",sInput,"si")


      sInput = "<b>" & sImagenC & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")




    if clng(rsx(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha
      celdas = celdas & ofv.GenCelda("",sc & " ","","37%","","","",sInput,"si")

      sInput = "<b>" & sImagenE & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")

    if clng(rsx(4)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rsx(4) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha



      celdas = celdas & ofv.GenCelda("",sc & " ","","37%","","","",sInput,"si")



    IF XDET = "1" THEN filas = ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")


      celdas = ofv.GenCelda("","vt ","","1%","","","","","si")

    IF XMOD = "PRY" THEN

       sinput = "&nbsp;"

    ELSE

       sinput = "<FONT COLOR=#336699 SIZE=1><U>" &  rsx(19) & "</U></FONT>"

    END IF



      celdas = celdas & ofv.GenCelda("","vt ","","3%","","","",sInput,"si")

      sInput = "<b>" & sImagenO & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")


      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(13) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
      celdas = celdas & ofv.GenCelda("",sc & " " ,"","37%","","","",sInput,"si")

      sInput = "<b>" & sImagenI & "</b>"
      celdas = celdas & ofv.GenCelda("","vt","","1%","","","",sInput,"si")

      sniv = "---"

                strsql = "Select * From USUARIOS where ID = '" & RSX(14) & "' "
                set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
                if not rsniv.eof then sniv = rsniv(1)
                call ofv.cerrarconsulta(rsniv)

      sInput = sniv
      celdas = celdas & ofv.GenCelda("",sc & " " ,"","37%","","","",sInput,"si")



    IF XDET = "1" THEN  filas = filas & ofv.GenRow("","aaa  " & sr,"","","","","",celdas,"si")


    IF XDET = "1" THEN  formu = formu & filas

      celdas = ofv.GenCelda("","AAA COLSPAN=6 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#CCCCCC;FONT-SIZE:5pt' ","","1%","","","","&nbsp;","si")

    IF XDET = "1" THEN  formu = formu & ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")



    posi = posi + 1
   
    END IF 

    end if
    rsX.Movenext
  loop


      sinput = "<b>TOTAL DE " & STABLA & "S: </b>" 
      celdas = ofv.GenCelda("xx1","vt COLSPAN=6 style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#6699cc;' VALIGN='BOTTOM' ","","1%","40","","",sInput,"si")
      formu = formu & ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")

      FOR ITAB=1 TO TTAB


          celdas = ofv.GenCelda("","VT COLSPAN=2 " ,"","","","","",ATABLA(ITAB,0),"si")

    IF XMOD = "PRY" THEN

       sInput = "&nbsp;" & ATABLA(ITAB,1)

    ELSE

       sInput = "&nbsp;Planificados:&nbsp;&nbsp;" & ATABLA(ITAB,2)
       sInput2 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No planificados:&nbsp;&nbsp;" & ATABLA(ITAB,3)
       sInput3 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total:&nbsp;&nbsp;" & ATABLA(ITAB,1)

       SINPUT = SINPUT & SINPUT2 & SINPUT3

    END IF




      celdas = celdas & ofv.GenCelda("","VT COLSPAN=4 " ,"","","","","",sInput,"si")
      formu = formu & ofv.GenRow("","aaa " & sr,"","","","","",celdas,"si")

      NEXT

          celdas = ofv.GenCelda("","VT bgcolor=#ffffff COLSPAN=2 STYLE='COLOR:#993333;' style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#6699cc;' " ,"","","","","",ATABLA(0,0),"si")

    IF XMOD = "PRY" THEN

       sInput = "&nbsp;" & ATABLA(0,1)

    ELSE

       sInput = "&nbsp;Planificados:&nbsp;&nbsp;" & ATABLA(0,2)
       sInput2 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;No planificados:&nbsp;&nbsp;" & ATABLA(0,3)
       sInput3 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Total:&nbsp;&nbsp;" & ATABLA(0,1)

       SINPUT = SINPUT & SINPUT2 & SINPUT3

    END IF




      celdas = celdas & ofv.GenCelda("","VT bgcolor=#ffffff COLSPAN=4  STYLE='COLOR:#993333;' style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#6699cc;' " ,"","","","","",sInput,"si")
      formu = formu & ofv.GenRow("","aaa  VALIGN='BOTTOM'  ","","","30","","",celdas,"si")


  GENLISTA = FORMU

END FUNCTION

%>