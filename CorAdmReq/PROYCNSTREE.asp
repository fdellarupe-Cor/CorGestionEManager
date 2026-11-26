<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.MenuHeader("",funcion)


if session("FECOPC") = "" THEN
   SESSION("FECOPC") = "1"
   SESSION("FECDES") = 0
   SESSION("FECHAS") = 99999 
   SESSION("SELOPC") = "1"
END IF



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

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

 XMOD = request.querystring("XMOD")
 XDST = request.querystring("XDST")
 XSET = request.querystring("XSET")
 XMNN = request.querystring("MNN")
 sopcion = request.querystring("opcion")
 sRESP = request.querystring("RESP")

 IF XMOD = "" THEN XMOD = "TAR"
 IF XDST = "" THEN XDST = "ORG"

SELECT CASE XMOD

       CASE "TAR"

          XTABLA = "PROYRESP"
          XSELEC = " USRAFC <> 'NO' "

       CASE "PRY"

          XTABLA = "PROYRESP"
          IF XDST = "ORG" THEN
             XSELEC = " USRAFC <> 'NO' AND TIPO = 'OR' "
          ELSE
             XSELEC = " USRAFC <> 'NO' AND TIPO = 'AN' "
          END IF

       CASE "REQ"

          XTABLA = "REQSECTORES"
          IF XDST = "ORG" THEN
             XSELEC = " USRAFC <> 'NO' AND TIPO = 'OR' "
          ELSE
             XSELEC = " USRAFC <> 'NO' AND TIPO = 'AN' "
          END IF

END SELECT


session("NCliente") = "NO"
SESSION("VPLAN")    = 5555


    DIM ATAREA(200,30)
    DIM ATAREA2(200,30)
    DIM ATAREA3(200,30)

   ITART = 0
   ITATI = 0
   TTAR  = 0

   IETPT = 0
   IETPI = 0 
   TETP  = 0

   CDESP = 0
   CHASP = 0
   CDESR = 0
   CHASR = 0
   CMAX  = 0  




     sopcion = "VP"


    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>&#0128;</FONT>"
    sImagen4n = "<FONT FACE=WEBDINGS COLOR=#FFFFFF SIZE=2>x</FONT>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>&#0233;</FONT></B>"
    sImagen6A = "<B><FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=3>&#0190;</FONT></B>"
    sImagen6W = "<B><FONT FACE=WINGDINGS COLOR=#CCCCCC SIZE=2>s</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"
    sImagenE4 = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>&#0232;</FONT></B>"


    sImageG = "<B><FONT FACE=WEBDINGS COLOR=#007700 SIZE=3>=</FONT></B>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#6699CC SIZE=3>4</FONT></B>"
    sImageR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0128;</FONT></B>"


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCHECK.asp"
     sinput = sinput  & "  TITLE='CHECK'  ><b>" & sImageR & "</b></a>" 

'   RESPONSE.WRITE "MENU: " & sinput & "<BR>"

'   RESPONSE.WRITE "MENU: " & SESSION("FORMN") & "<BR>"


   SELECT CASE XSET

          CASE "XRESP" 

  ITART   = 1
  ITART2  = 1
  ITART3  = 0
  TTAR    = 1

  StrSql = "Select DISTINCT USRAFC From " & XTABLA & " where " & XSELEC & "  "
 '' response.write strsql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then

        ATAREA (ITART,2)    = "H"
        ATAREA (ITART,3)    = 1
        ATAREA (ITART,4)    = 1

        ATAREA2 (ITART2,2)  = "H"
        ATAREA2 (ITART2,3)  = 1
        ATAREA2 (ITART2,4)  = 0

     do while not rs1.eof






    if RS1(0) = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(0) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF



    ITART3              = ITART3 + 1
    ATAREA3 (ITART3,0)  = RS1(0)
    ATAREA3 (ITART3,1)  = sniv
    ATAREA3 (ITART3,2)  = "N"
    ATAREA3 (ITART3,3)  = 0
    ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST & "&XOPT=XRESP&RESP=" & RS1(0)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageR & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

    rs1.movenext
    loop
    ATAREA2 (ITART2,4)  = ITART3

    call ofv.cerrarconsulta(rs1)
  end if




          CASE "XGS" 

  XGERANT = 0
  XSECANT = 0

  ITART   = 1
  ITART2  = 0
  ITART3  = 0
  TTAR    = 1

  Strsql = "Select * from clienteplan where id = 1"
  set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
  if not rs.eof then
     NOMEMP = rs(1)
     LOGEMP = rs(5)
  end if
  call ofv.cerrarconsulta(rs)


  sImage = ofv.GenImage("","","","","../COREss20/Empresa/" & LOGEMP ,"0","")

        ATAREA (ITART,0)   = 1
        ATAREA (ITART,1)   = NOMEMP
        ATAREA (ITART,2)   = "N"
        ATAREA (ITART,3)   = 0
        ATAREA (ITART,4)   = 0


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST & "&XOPT=XGLB&RESP=1"
     sinput = sinput  & "  TITLE='REPORTES DE " & NOMEMP & "'  ><b>" & sImage & "</b></a>" 
        ATAREA (ITART,5)  = sinput

  StrSql = "Select DISTINCT GERENCIA, SECTOR, USRAFC From " & XTABLA & " where  " & XSELEC & " ORDER BY GERENCIA, SECTOR, USRAFC "
  '''response.write strsql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof


     IF XGERANT <> RS1(0)  then

        ATAREA2 (ITART2,4) = ITART3 
        ATAREA (ITART,4)   = ITART2 


        ITART              = ITART + 1
        XGERANT            = RS1(0)

     StrSql = "Select * From ECO where ECO = " & RS1(0)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

        ATAREA (ITART,0)   = RS1(0)
        ATAREA (ITART,1)   = sniv
        ATAREA (ITART,2)   = "S"
        ITART2             = ITART2 + 1
        ATAREA (ITART,3)   = ITART2
        ATAREA (ITART,4)   = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=XGRN&RESP=" & RS1(0)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageG & "</b></a>" 
        ATAREA (ITART,5)  = sinput


        TTAR               = ITART 

        XSECANT            = RS1(1)

     StrSql = "Select * From ECO where ECO = " & RS1(1)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

        ATAREA2 (ITART2,0)  = RS1(1)
        ATAREA2 (ITART2,1)  = sniv
        ATAREA2 (ITART2,2)  = "S"
        ITART3              = ITART3 + 1
        ATAREA2 (ITART2,3)  = ITART3
        ATAREA2 (ITART2,4)  = 0


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=XSCT&RESP=" & RS1(1)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA2 (ITART2,5)  = sinput

    if RS1(2) = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(2) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF

        ATAREA3 (ITART3,0)  = RS1(2)
        ATAREA3 (ITART3,1)  = sniv
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=XRESP&RESP=" & RS1(2)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageR & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

     ELSEIF XSECANT <> RS1(1)  then

        ATAREA2 (ITART2,4) = ITART3 

        XSECANT            = RS1(1)

     StrSql = "Select * From ECO where ECO = " & RS1(1)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

        ITART2              = ITART2 + 1

        ATAREA2 (ITART2,0)  = RS1(1)
        ATAREA2 (ITART2,1)  = sniv
        ATAREA2 (ITART2,2)  = "S"
        ITART3              = ITART3 + 1
        ATAREA2 (ITART2,3)  = ITART3
        ATAREA2 (ITART2,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=XSCT&RESP=" & RS1(1)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageS & "</b></a>" 
        ATAREA2 (ITART2,5)  = sinput

    if RS1(2) = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(2) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF

        ATAREA3 (ITART3,0)  = RS1(2)
        ATAREA3 (ITART3,1)  = sniv
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=XRESP&RESP=" & RS1(2)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageR & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

     ELSE


        ITART3             = ITART3 + 1

    if RS1(2) = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(2) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF



        ATAREA3 (ITART3,0)  = RS1(2)
        ATAREA3 (ITART3,1)  = sniv
        ATAREA3 (ITART3,2)  = "N"
        ATAREA3 (ITART3,3)  = 0
        ATAREA3 (ITART3,4)  = 0

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYCNSLST.asp?XSET=" & XSET & "&XMOD=" & XMOD & "&XDST=" & XDST  & "&XOPT=XRESP&RESP=" & RS1(2)
     sinput = sinput  & "  TITLE='REPORTES DE " & sniv & "'  ><b>" & sImageR & "</b></a>" 
        ATAREA3 (ITART3,5)  = sinput

     END IF





    rs1.movenext
    loop

        ATAREA2 (ITART2,4) = ITART3 
        ATAREA (ITART,4)   = ITART2 


    call ofv.cerrarconsulta(rs1)
  end if


   END SELECT 



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 


      celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>REPORTES " & SESSION("FORMN") & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")





      celdas = ofv.GenCelda("","VT COLSPAN=4 STYLE='border-bottom-style:RIDGE;border-bottom-width:2;border-bottom-color:#f7aa22;' ","CENTER","90%","","","","RESPONSABLES","si")

      filas = ofv.GenRow("","","","","10","","",celdas,"si")


''''''''''''''''


  for ITART=1 to ttar

      IF ATAREA (ITART,2) <> "H"  THEN

      celdas = ofv.GenCelda("","VT ","right","","","","",ATAREA (ITART,5),"si")
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=3 ","","","","","",ATAREA (ITART,1),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
    
      END IF

      IF ATAREA (ITART,2) = "S" OR ATAREA (ITART,2) = "H" THEN

         for ITART2=ATAREA (ITART,3) to ATAREA (ITART,4)

             IF ATAREA2 (ITART2,2) <> "H"  THEN

             celdas = ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
             celdas = celdas & ofv.GenCelda("","","right","","","","",ATAREA2 (ITART2,5),"si")
             celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",ATAREA2 (ITART2,1),"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

             END IF

             IF ATAREA2 (ITART2,2) = "S" OR ATAREA2 (ITART2,2) = "H"  THEN

                for ITART3=ATAREA2 (ITART2,3) to ATAREA2 (ITART2,4)

                    celdas = ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
                    celdas = celdas & ofv.GenCelda("","","CENTER","","","","","&nbsp;","si")
                    celdas = celdas & ofv.GenCelda("","","right","","","","",ATAREA3 (ITART3,5),"si")
                    celdas = celdas & ofv.GenCelda("","ut ","","","","","",ATAREA3 (ITART3,1),"si")
                    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

               next

             END IF


          next

      END IF


  next



      celdas = ofv.GenCelda("","","CENTER","","2%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","10%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","10%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","80%","","","&nbsp;","si")

    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    FilaS2 = ofv.GenTabla("","aa style='border-style:ridge;border-width:2;border-color:#f7aa22;' ","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","AAA COLSPAN=2","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")



    sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





'''''''''''''''












 







  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>