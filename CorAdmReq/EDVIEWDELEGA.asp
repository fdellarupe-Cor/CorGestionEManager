<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FORMHeader("")



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





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","TTT colspan=3","","","","","",sImagen & "&nbsp;<B>DELEGACIONES DEL REQUERIMIENTO</B>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    sinput = "RESPONSABLE&nbsp;DELEGANTE"
    celdas = ofv.GenCelda("","TTR","CENTER","40%","","","",sInput,"si")

    sinput = "&nbsp;"
    celdas = celdas & ofv.GenCelda("","UT","CENTER","10%","","","",sInput,"si")

    sinput = "NUEVO&nbsp;RESPONSABLE"
    celdas = celdas & ofv.GenCelda("","TTR","CENTER","40%","","","",sInput,"si")

    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","vt colspan=3","","100%","","","",sLink,"si")
  '  filas = filas & ofv.GenRow("","","","","","","",Celda,"si")



    sImagenI3 = "<B><FONT FACE=WINGDINGS COLOR=#336699 SIZE=4>&#0232;</FONT></B>"
    sImagen4so = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=2>&#0128;</FONT></B>"
    sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#aa0000 SIZE=2>&#0128;</FONT></B>"



    StrSql = "Select * From REQDELEGA WHERE CODIGO = '" & session("NCliente") & "' ORDER BY ORDEN DESC "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof THEN

       DO WHILE NOT RS.EOF

celdas = ""

    if RS(7) = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & RS(7) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagen4so & "&nbsp;<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","vt","CENTER","","","","",sInput,"si")
    end if


      celdas = celdas & ofv.GenCelda("","vt ROWSPAN=2 ","","","","","",sImagenI3,"si")


    if RS(12) = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & RS(12) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagen4sI & "&nbsp;<B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","vt","CENTER","","","","",sInput,"si")
    end if




      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")









celdas = ""

    if RS(3) = 0 then
      celdas = celdas & ofv.GenCelda("","vt","CENTER","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(3)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = "<FONT COLOR=#6699CC>" & rsX2(1) & "</FONT>"

     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","vt","CENTER","","","","",sInput,"si")


    end if

  '    celdas = celdas & ofv.GenCelda("","vt","","","","","","&nbsp;","si")


    if RS(10) = 0 then
      celdas = celdas & ofv.GenCelda("","vt","CENTER","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(10)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = "<FONT COLOR=#6699CC>" & rsX2(1) & "</FONT>"
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","vt","CENTER","","","","",sInput,"si")


    end if


      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    IF RS(13) = 0 THEN
       sLink = "Efectivo"
    else
       sLink = "<font color=#993333>Cancelado</font>"
    end if 

    sLink = RS(15) & " - " & sLink
    Celda = ofv.GenCelda("","vt colspan=3 bgcolor=#ffffff style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","right","100%","","","",sLink,"si")
    filas = filas & ofv.GenRow("","","","","","","",Celda,"si")





    RS.MOVENEXT
    LOOP

   ELSE

    sLink = "NO SE REGISTRAN DELEGACIONES PARA ESTE REQUERIMIENTO"
    Celda = ofv.GenCelda("","vt colspan=3 bgcolor=#ffffff style='border-bottom-style:ridge;border-bottom-width:1;border-bottom-color:#f7aa22;' ","right","100%","","","",sLink,"si")
    filas = filas & ofv.GenRow("","","","","","","",Celda,"si")

   END IF
   call ofv.cerrarconsulta(RS)



   fila =  ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"


    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"




  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

''''''''''''''''''''''

  sHTML   = sHTML & CorPnlCOT()

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clpanel();' ", "S", "A", "N", "indFrm", "")

  Agregar = Celt1

  Buscar = ""
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

''''''''''''''''''''''


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML





%>