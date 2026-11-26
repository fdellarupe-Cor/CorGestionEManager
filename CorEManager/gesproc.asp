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



  procnivel = request.querystring("procnivel")
  proccodg  = request.querystring("proccodg")
  sopcion   = request.querystring("opcion")
  codexterno=request.querystring("Ssecuen")

 ' RESPONSE.WRITE request.querystring


 if procnivel <> "" then

   session("procnivel")=procnivel
 else
   procnivel=session("procnivel")
 end if


 IF PROCNIVEL = "" THEN
    PROCNIVEL = 1
  ELSE
    PROCNIVEL = CINT("0" & PROCNIVEL)
 END IF



 if proccodg<>"" then
     session("proccodg")=proccodg
  else
    proccodg=session("proccodg")
 end if



 if codexterno<>"" then
     session("codexterno")=codexterno
  else
    codexterno=session("codexterno")
 end if




   CantRegMoverx = cdbl(request.querystring("CantRegMover"))
   'CantRegMoverx2 = cdbl(request.querystring("CantRegMover2"))
   if PROCNIVEL=1 then
    session("cantregmover2")=CantRegMoverx
   end if

'   response.write session("cantregmover2")
  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()




 ' vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
   vuelta = ofv.xVueltaAsp("../",Cantregmoverx,Cantregmoverx2,CantRegMostrarx) & "_procnivel=" &  procnivel & "_proccodg=" & proccodg

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



           IF PROCNIVEL > 1 AND (session("APLMDLOR") & "") = "BE" THEN strsql = strsql & ",CODEXTERNO"

           strsql = strsql & " ) "
           strsql = strsql & " VALUES ('" & newcodg & "','NUEVO PROCESO','0','" & procnivel & "'"

           strsql = strsql & strsql2

           IF PROCNIVEL > 1 AND (session("APLMDLOR") & "") = "BE" THEN strsql = strsql & ",'" & CODEXTERNO & "'"
           strsql = strsql & ") "
       '    response.write strsql & "   " & CODEXTERNO
           call ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if





  sHTML = ofv.FormHeader(session("FormN"))

  '  sHTML = sHTML & "<br><br><br>"

 '   for nn=0 to ubound(atabproc)
  '  sHTML = sHTML & "<p>" & atabproc(nn) & "</p>"

 '   next




    IF PROCNIVEL > 1 THEN

    Strsql = "Select dscproc,secuen,CODEXTERNO From csprocesos "
    strsql = strsql & " where codproceso = '" & proccodg & "' "
    'response.write strsql
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if Not rs.EOF then

    Celda = ""
    sInput = "Proceso :"
    Celda = Celda & ofv.GenCelda("","ttt","","10%","","","",sInput,"si")
    if (session("APLMDLOR") & "") = "BE" THEN
      'sInput = rs(1)&"-" &rs(0)
      'sSECUEN=rs(1)&"-"& RS(2)

	  if session("APLMDCLI") = "CFA" THEN
		sInput = rs(1)&"-" &rs(0)
		sSECUEN=rs(1)&"-"& RS(2)
	  ELSE
		sInput = rs(0)
		sSECUEN=""
	 END IF

    ELSE
      sInput = rs(0)
      sSECUEN=""
    END IF
    Celda = Celda & ofv.GenCelda("","vt","","80%","","","",sInput,"si")


    'sImagenP = "&#0232;"
   ' sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
   ' sinput = "<a class='ft1n' " & sFuncAnc & " href='gesproc.asp?PROCNIVEL=" & (PROCNIVEL -1 ) & "&PROCCODG=" & ATABPROC(PROCNIVEL - 2)
   ' sinput = sinput  & "' title='Volver' >" & sImagenP & "</a>"
    sAccion="gesproc.asp?PROCNIVEL=" & (PROCNIVEL -1 ) & "&PROCCODG=" & ATABPROC(PROCNIVEL - 2) &"&cantregmover=" & session("cantregmover2")
    Celdax = ofv.BotonVolver(sAccion)
    'Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","25","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80%","","0","0","0","0","","",Fila,"si") & ""

   end if

  end if



  'Encabezados
  Fila = ofv.Encabezados("codigo","secuencia","proceso","descripcion","subprocesos")


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
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi< CantRegMostrarx)
    Formulario = ""
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
    '  sHref = "asp/Eliminar.asp?ID=" & ofv.convertircar(rs(0)," ","_")
    '  sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC"
    '  Celda = ofv.BotonEliminar(rs(1),sHref)

      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=TABPROCESOS01&volver=" & vuelta  & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(3),sHref)


    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'orden
    sInput = (rs(0) & "") & "&nbsp;"
    Celda = Celda & ofv.GenCelda("","cc","right","","","","",sInput,"si")

    'orden
    sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"5","dtn","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


    'PROCESO

    sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"20","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")



    'Descripcion
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"80","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


       PROCCODGw = rs(1)

  if (session("APLMDLOR") & "") = "BE" and (PROCNIVEL > 1 ) then

    sinput = "&nbsp;---"

  else

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "


    sinput = "<a href='gesproc.asp?PROCNIVEL=" & (PROCNIVEL + 1) & "&PROCCODG=" & PROCCODGw & "&Ssecuen=" &Ssecuen
    sinput = sinput  & "' title='Subprocesos de " & rs(3) & "' ><i class='las la-arrow-right'></i></a>"

  end if

    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=TABPROCESOS01&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    posi=posi+1
    rs.Movenext
  Loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px'","","80%","","0","","0","0","","",filas,"si")

  Agregar = ""
  if bAgregar then
    sAccion = "GESPROC.asp?OPCION=SI&PROCNIVEL=" & PROCNIVEL & "&PROCCODG=" & PROCCODG &"&Ssecuen=" & Ssecuen
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRIPCION From TABPROCESOS01 Order By SECUEN "    ' orden pagina
  str2 = "Select DESCRIPCION From TABPROCESOS01 Order By DESCRIPCION "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

   'consulta
    sAccion = "GESPROCCONS.asp?" & request.querystring
    Celda = ofv.BotonConsultar(sAccion)
    
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,celda,celdax )
  'sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda,volverxx)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>" '&PROCNIVEL& request.querystring '& "PROCNIVEL=" & (PROCNIVEL -1 ) & "&PROCCODG=" '& ATABPROC(PROCNIVEL - 2)
end if

Response.Write sHTML

%>