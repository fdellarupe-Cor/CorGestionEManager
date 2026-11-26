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



 FUNCION = ""

  sHTML = ofv.FormHeader(replace(session("FormN"),"_",FUNCION))


  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 






    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;RESUMEN DEL REQUERIMIENTO","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sImagen = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"


  celdas = ofv.GenCelda("","vt colspan=3 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","","","","","",sImagen & "&nbsp;<B>SOLICITUD</B>","si")
  filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    celdas = ofv.GenCelda("","ut","","","","","","TITULO:","si")
'    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
     sInput = rs(2)
    celdas = celdas & ofv.GenCelda("","Vt COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","ut","","","","","","CODIGO:","si")
'  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

  celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",rs(1),"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  celdas = ofv.GenCelda("","ut","","","","","","CLASE:","si")
 ' celdas = celdas & ofv.GenCelda("","","","","","","","","si")

     StrSql = "Select * From TIPOREQ where TIPOREQ = '" & RS(15) & "' "
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        XTIPOREQ  = RS(15) & " - " & rsX2(2)

     ELSE
        XTIPOREQ  = "SIN DETERMINAR"

     END IF

  celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","","" & REPLACE(TRIM(XTIPOREQ)," ","&nbsp;") & "","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    sImagen = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"


    celdas = ofv.GenCelda("","ut","","","","","","TIPO:","si")
'    celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")

      sniv = ""
      strsql = "Select * From TIPOS where ID = " & rs(8)
      set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = sniv
    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","ORIGEN:","si")
   ' celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
      sniv = ""
      strsql = "Select * From ORIGEN where OBJID = " & rs(7)
      set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = sniv
    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","DESCRIPCION ORIGEN:","si")
  '  celdas = celdas & ofv.GenCelda("","","","","","","",sImagen,"si")
     sInput = rs(12)
    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","FECHA INICIO:","si")
 '   celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(3)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(3) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","FECHA SOLICITUD:","si")
  '  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(4)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(4) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","FECHA ENTREGA:","si")
  '  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(5)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(5) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","FECHA FINALIZACION:","si")
  '  celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(6)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(6) + clng(cdate("01-01-1900")))
    end if

      sInput = dFecha


    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    celdas = ofv.GenCelda("","ut","","","","","","ESTADO:","si")
 '   celdas = celdas & ofv.GenCelda("","","","","","","","","si")

      sInstrumento = "---"

        Strsql = "Select descri From reqestados Where id = " & rs(11)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento


    celdas = celdas & ofv.GenCelda("","uT COLSPAN=2 ","","20","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    celdas = ofv.GenCelda("","ut","","","","","","ORIGINANTE:","si")
   ' celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if RS(13) = "NO" then
      sInput = "SIN ORIGINANTE"
      celdas = celdas & ofv.GenCelda("","uT COLSPAN=2 ","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = "" & sniv & ""
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas = ofv.GenCelda("","ut","","","","","","ANALISTA ASIGNADO:","si")
'    celdas = celdas & ofv.GenCelda("","","","","","","","","si")
    if RS(14) = "NO" then
      sInput = "SIN ASIGNAR"
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & rs(14) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = "" & sniv & ""
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","","","","",sInput,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vt colspan=3 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","","","","","",sImagen & "&nbsp;<B>ALCANCE</B>","si")
  filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if RS(9) = 0 then
      celdas = ofv.GenCelda("","ut","","","","","","TEMA/SUBTEMA:","si")
      sInput = "SIN DETERMINAR"
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    else
      XTST = SPLIT(GETTEMASUBTEMA(RS(9)),CHR(10))
      SINPUT1 = "" & XTST(0) & ""
      SINPUT2 = "" & XTST(1) & ""
      
      celdas = ofv.GenCelda("","ut ","","","","","","TEMA:","si")
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=2","","","","","",sInput1,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","ut ","","","","","","SUBTEMA:","si")
      celdas = celdas & ofv.GenCelda("","ut COLSPAN=2","","","","","",sInput2,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    end if





    'Encabezado DESCRIPCION
    Celda = ofv.GenCelda("","aaa colspan=3","","100%","","","","&nbsp;","si")
    FilaS =  FilaS & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","Ut colspan=3","","100%","","","","DESCRIPCION DEL REQUERIMIENTO","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

  if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

  if not  isnull(rs(10)) then 
     wdescri = rs(10)
  end if 

    t3b =  wdescri
    Celda = ofv.GenCelda("","ut colspan=3  style='border-style:solid;border-width:1;border-color:#6699cc;' ","","100%","","","",t3b,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vt colspan=3 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","","","","","",sImagen & "&nbsp;<B>GERENCIAS Y SECTORES AFECTADOS</B>","si")
  filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select * From REQSECTORES where codigo = '" & session("NCliente") & "' ORDER BY ORDEN "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  celdas = ofv.GenCelda("","ut  style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","center","40%","","","","GERENCIA","si")
  celdas = celdas & ofv.GenCelda("","ut  style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","center","40%","","","","SECTOR","si")
  celdas = celdas & ofv.GenCelda("","ut  style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","center","20%","","","","TIPO","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  DO UNTIL RS1.EOF

    if RS1(4) = 0 then
      celdas = ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS1(4)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
        sInput = replace(sInput," ","&nbsp;")
     celdas = ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","",sInput,"si")


    end if

    if RS1(3) = 0 then
      celdas = CELDAS & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS1(3)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
        sInput = replace(sInput," ","&nbsp;")
     celdas = CELDAS & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","",sInput,"si")


    end if

    SELECT CASE RS1(5)
           CASE "OR"
                 sInput = "ORIGINANTE"
           CASE "AN"
                 sInput = "INTERVINIENTE"
           CASE "AF"
                 sInput = "AFECTADO"
    END SELECT
        sInput = replace(sInput," ","&nbsp;")
     celdas = CELDAS & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","","<b>" & sInput & "</b>","si")





      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    RS1.MOVENEXT
    LOOP

dim atab(15,1)
dim atabN(15,1)


    for xx = 1 to 15
           atab(xx,0) = 99
           atabN(xx,0) = 99
    next

    for xx = 1 to 9

           atab(XX,0) = xx + 2



    next



     StrSql = "Select * From TIPOREQ WHERE TIPOREQ = '100' "
     set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsX.eof then




        for yy = 1 to 15
 
            if clng(atab(yy,0)) = 99 then exit for
               xx = atab(yy,0)

               IF RSX(XX) THEN
                  atabN(yy,0) = XX
                  atabN(yy,1) = rsX(xx + 15)
               ELSE
                  atabN(yy,0) = 99
               END IF 
        next

     end if




  StrSql = "Select * From REQINFO where CODIGO = '" & session("NCliente") & "' ORDER BY TIPO "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  DO UNTIL rs1.eof 

     wdescri = rs1(3)

  if len(wdescri) > 0 then 

     celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     celdas = ofv.GenCelda("","vt colspan=3 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","","","","","",sImagen & "&nbsp;<B>INFORMACION " & replace(Ucase(atabN(RS1(2)-2,1))," ","&nbsp;") & "</B>","si")
     filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     wdescri = rs1(3)

  if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

  if not  isnull(rs1(3)) then 
     wdescri = rs1(3)
  end if 


    t3b =  wdescri
    Celda = ofv.GenCelda("","ut colspan=3  style='border-style:solid;border-width:1;border-color:#6699cc;' ","","100%","","","",t3b,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    celdas = ofv.GenCelda("","ut","","","","","","FECHA&nbsp;ACTUALIZACION:","si")


    if clng(rs1(5)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs1(5) + clng(cdate("01-01-1900")))
    end if

      sInput = "&nbsp;" & dFecha

    celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 ","","80%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

   END IF

  RS1.MOVENEXT
  LOOP





  StrSql = "Select * From REQCONTACTOS where codigo = '" & session("NCliente") & "' ORDER BY ID "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  IF NOT RS1.EOF THEN





  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vt colspan=3 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","","","","","",sImagen & "&nbsp;<B>CONTACTOS</B>","si")
  filas  = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  celdas = ofv.GenCelda("","ut style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","center","","","","","CONTACTO","si")
  celdas = celdas & ofv.GenCelda("","ut COLSPAN=2 style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","center","","","","","SECTOR","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  DO UNTIL RS1.EOF


         sniv = "INEXISTENTE"
         snivS = 0

      strsql = "Select * From USUARIOS where ID = '" & rs1(2) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(1)
         snivS = rsniv(21)
      END IF
      call ofv.cerrarconsulta(rsniv)
      sInput = "" & sniv & ""
     celdas = ofv.GenCelda("","ut  style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","",sInput,"si")


     IF SNIVS > 0 THEN
        strsql = "Select * From ECO where ECO = " & SNIVS
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then
           sniv = rsniv(1)
        END IF
        call ofv.cerrarconsulta(rsniv)
        sInput = sniv
        celdas = CELDAS & ofv.GenCelda("","ut  COLSPAN=2  style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","",sInput,"si")  
     ELSE
        celdas = CELDAS & ofv.GenCelda("","ut  COLSPAN=2  style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","","SIN AFECTAR","si")
     END IF 




      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





    RS1.MOVENEXT
    LOOP


  END IF


  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
  sHTML = sHTML & "</center><br><br>"

 

  end if



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

FUNCTION GETTEMASUBTEMA(SSUBTEMA)

  xTEMA = 0
  xSUBTEMA = "SIN ASIGNAR"

  SQL2 = "SELECT * "
  SQL2 = SQL2 & " FROM csbuscasubtema where "
  SQL2 = SQL2 & " codPRODSERV = " & SSUBTEMA
  set rsX = ofv.crearconsultaEx(SQL2,cn3,1,parametros)
  if NOT rsX.eof then
     xTEMA = RSX(3)
     xSUBTEMA = RSX(1)
     SQL1 = "SELECT * From Tema WHERE CODTEMA = " & xTEMA
     set rsX = ofv.crearconsultaEx(SQL1,cn3,1,parametros)
     if NOT rsX.eof then
        xTEMA = RSX(1)
     END IF

  END IF

  GETTEMASUBTEMA = XTEMA & CHR(10) & XSUBTEMA

END FUNCTION


%>