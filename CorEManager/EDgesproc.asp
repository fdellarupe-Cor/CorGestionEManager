<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  procnivel = request.querystring("procnivel")
  proccodg  = request.querystring("proccodg")
  sopcion   = request.querystring("opcion")
  nCodigo   = request.querystring("Codobjeto")
  sTipo     = request.querystring("Tipobjeto")
  sACTCOD   = request.querystring("OLDCOD")
  sCODEVEN  = request.querystring("CODEVEN")








  IF PROCNIVEL = "" THEN
     PROCNIVEL = 1
  ELSE
     PROCNIVEL = CINT("0" & PROCNIVEL)
  END IF



  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_procnivel=" &  procnivel & "_proccodg=" & proccodg


  if session("TABPROC") = "" THEN
     TABPROC = "NO" & CHR(9) &  "NO" & CHR(9) & "NO" & CHR(9) & "NO" & CHR(9) & "NO" & CHR(9) & "NO"

     session("TABPROC") = TABPROC

  END IF 

  TABPROC = session("TABPROC")

  ATABPROC = SPLIT(TABPROC,CHR(9))

  IF PROCNIVEL = 1 THEN 
     ATABPROC(1) = "NO"
     TABPROC = join(ATABPROC,CHR(9))
  else
     IF PROCNIVEL < 5 THEN 
        ATABPROC(PROCNIVEL - 1) = proccodg
     end if
     TABPROC = join(ATABPROC,CHR(9))
  end if

     session("TABPROC") = TABPROC


  set cn = ofv.conectar(ofv.strconn2)


  if sopcion = "SI" then

           strsql2 = ""

           newcodg = GTkey(procnivel)

 
           strsql = "INSERT INTO TABPROCESOS01 (codproceso,DESCRIPCION,SECUEN,nivel"

    IF PROCNIVEL = 5 THEN
           strsql = strsql & ",CODPROCESO05"
           strsql = strsql & ",CODPROCESO04"
           strsql = strsql & ",CODPROCESO03"
           strsql = strsql & ",CODPROCESO02"
           strsql = strsql & ",CODPROCESO01"

           strsql2 = strsql2 & ",'" & newcodg & "'"

           strsql2 = strsql2 & ",'" & ATABPROC(4) & "'"
           strsql2 = strsql2 & ",'" & ATABPROC(3) & "'"
           strsql2 = strsql2 & ",'" & ATABPROC(2) & "'"
           strsql2 = strsql2 & ",'" & ATABPROC(1) & "'"


    END IF

    IF PROCNIVEL = 4 THEN
           strsql = strsql & ",CODPROCESO04"
           strsql = strsql & ",CODPROCESO03"
           strsql = strsql & ",CODPROCESO02"
           strsql = strsql & ",CODPROCESO01"

           strsql2 = strsql2 & ",'" & newcodg & "'"


           strsql2 = strsql2 & ",'" & ATABPROC(3) & "'"
           strsql2 = strsql2 & ",'" & ATABPROC(2) & "'"
           strsql2 = strsql2 & ",'" & ATABPROC(1) & "'"


    END IF

    IF PROCNIVEL = 3 THEN
           strsql = strsql & ",CODPROCESO03"
           strsql = strsql & ",CODPROCESO02"
           strsql = strsql & ",CODPROCESO01"

           strsql2 = strsql2 & ",'" & newcodg & "'"


           strsql2 = strsql2 & ",'" & ATABPROC(2) & "'"
           strsql2 = strsql2 & ",'" & ATABPROC(1) & "'"

    END IF

    IF PROCNIVEL = 2 THEN
           strsql = strsql & ",CODPROCESO02"
           strsql = strsql & ",CODPROCESO01"

           strsql2 = strsql2 & ",'" & newcodg & "'"

           strsql2 = strsql2 & ",'" & ATABPROC(1) & "'"


    END IF

    IF PROCNIVEL = 1 THEN
           strsql = strsql & ",CODPROCESO01"

           strsql2 = strsql2 & ",'" & newcodg & "'"
    END IF





           strsql = strsql & " ) "
           strsql = strsql & " VALUES ('" & newcodg & "','NUEVO PROCESO','0','" & procnivel & "'"

           strsql = strsql & strsql2 


           strsql = strsql & ") "
           call ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if





  sHTML = ofv.FormHeader(session("FormN"))

  '  sHTML = sHTML & "<br><br><br>"

 '   for nn=0 to ubound(atabproc)
  '  sHTML = sHTML & "<p>" & atabproc(nn) & "</p>"

 '   next

    Strsql = "Select * From csgenobjetos "
    strsql = strsql & " where codobjeto = 0" & nCodigo
 
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if Not rs.EOF then

       Celda = ""
       sInput = rs(1)
       Celda = Celda & ofv.GenCelda("","ttt","","","","","",sInput,"si")

       Fila = ofv.genrow("","","","","25","","",Celda,"si")

       sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si") & "<br>"

     end if
 



    IF PROCNIVEL > 1 THEN

    Strsql = "Select dscproc From csprocesos "
    strsql = strsql & " where codproceso = '" & proccodg & "' "
 
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if Not rs.EOF then

    Celda = ""
    sInput = "Subprocesos de:"
    Celda = Celda & ofv.GenCelda("","utg","","20%","","","",sInput,"si")

    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","utg","","70%","","","",sInput,"si")

    SHREF = "edgesproc.asp?PROCNIVEL=" & (PROCNIVEL -1 ) & "&PROCCODG=" & ATABPROC(PROCNIVEL - 2)
    SHREF = SHREF & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&CODEVEN=" & sCODEVEN

    sImagenP = "&#0233;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='" & SHREF & "' "
    sinput = sinput  & " title='Subir nivel' >" & sImagenP & "</a>" 
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","25","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","80%","","0","0","1","0","","",Fila,"si") & "<br>"

   end if

  end if



  'Encabezados
 ' Fila = ofv.Encabezados("secuencia","proceso","descripcion","subprocesos")


    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","Asignar","si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","Proceso","si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","Descripcion","si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","subprocesos","si")
    fila = ofv.GenRow("","","","","","","",Celda,"si")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From TABPROCESOS01 "
  strsql = strsql & " where "

  if procnivel > 1 then

    IF PROCNIVEL = 2 THEN
       strsql = strsql & " codproceso01 = '" & proccodg & "' "
    ELSEIF PROCNIVEL = 3 THEN
       strsql = strsql & " codproceso02 = '" & proccodg & "' "
    ELSEIF PROCNIVEL = 4 THEN
       strsql = strsql & " codproceso03 = '" & proccodg & "' "
    ELSE
       strsql = strsql & " codproceso04 = '" & proccodg & "' "
    END IF

       strsql = strsql & " and "

  end if

       strsql = strsql & " nivel = '" & procnivel & "' "


  strsql = strsql & " Order By SECUEN"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  Do While (Not rs.EOF)
    Formulario = ""


    'Seleccionar

    SHREF = "edprocesos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ACTCOD=" & rs(1)
    SHREF = SHREF & "&opcion=SI&ACTID=" & sCODEVEN

    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='" & shref & "' "
    sinput = sinput  & " title='Asignar proceso " & rs(3) & "' >" & sImagenP & "</a>" 
    Celda  = ofv.gencelda("","ft","center","","1%","","",sInput,"si")


    'orden
 '   sInput = rs(5)
 '   Celda  = Celda & ofv.GenCelda("","UT","","1%","","","",sInput,"si")


    'PROCESO
    sInput = rs(4)
    Celda = Celda & ofv.GenCelda("","UT","","20%","","","",sInput,"si")

    'Descripcion
    sInput = rs(3)
    Celda = Celda & ofv.GenCelda("","UT","","50%","","","",sInput,"si")

    PROCCODGw = rs(1)

    SHREF = "edgesproc.asp?PROCNIVEL=" & (PROCNIVEL + 1) & "&PROCCODG=" & PROCCODGw 
    SHREF = SHREF & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&CODEVEN=" & sCODEVEN

    sImagenP = "&#049;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='" & shref & "' "
    sinput = sinput  & " title='Subprocesos de: " & rs(3) & "' >" & sImagenP & "</a>" 
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si")

  Agregar = ""
  Buscar = ""

  sAccion = "edprocesos.asp?codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>