<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar

  'if session("Nomodificar") = "S" then bModificar = False
   IF Session("Form2")="DESARROLLO" OR  Session("Form2")="PARA APROBAR" OR  Session("Form2")="ELIMINADAS" THEN
    if session("Nomodificar") = "S" then bModificar = False

   END IF

  vuelta = ""


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=ddd>"

'-------------------------------------------------

  xopc = request.querystring("Opcion")
  if  xopc = "HE" then
      bModificar = True
      vuelta     = "?Opcion=" & xopc
      sformh     = "Habilitacion de Esquemas"
      shtml      = ofv.FormHeader(sFormh)
  end if

  smodalidad  = request.querystring("modalidad")


  if smodalidad <> "" then

     session("modalidad") = smodalidad

  else

     smodalidad  = session("modalidad")

  end if

  nCodigo  = request.querystring("Codobjeto")
  sTipo    = request.querystring("Tipobjeto")
  vuelta   = ofv.vueltaasp2("../") & vuelta
  vuelta2  = ofv.vueltaasp2("")


  set cn = ofv.conectar(ofv.strconn2)


   WKSPDESC = "Permisos del Esquema"


   IF SMODALIDAD = "PRG" OR SMODALIDAD = "PRGCOR" THEN

       Strsql = "Select * From OBJETOS  where CODOBJETO = " & nCodigo
       Strsql = Strsql & " and tipobjeto = '" & sTipo & "'  "

 '' response.write strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rs1.eof then
          WKTIPODESC2 = RS1(3)
       ELSE
          WKTIPODESC2 = "--"
       END IF

       Strsql = "Select * From tipobjeto  where tipobjeto = '" & sTipo & "' "
 '' response.write strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rs1.eof then
          WKTIPODESC1 = RS1(1)
       ELSE
          WKTIPODESC1 = "--"
       END IF

       WKSPDESC = "PERMISOS DE " & WKTIPODESC1 & ": " & WKTIPODESC2

       IF SMODALIDAD = "PRGCOR" THEN

          WKSPDESC = WKSPDESC & "&nbsp;-&nbsp;CIRCUITO DE AUTOEVALUACION"

       end if



   END IF

    Celda = ofv.gencelda("","TT colspan=4","center","2%","","","",WKSPDESC,"si")
    Fila = ofv.GenRow("","","","","20","","",Celda,"si")


   IF SMODALIDAD = "PRGCOR" THEN

      Strsql = "SELECT OBJID,CODOBJETO,TIPOBJETO,SEGCT,SEGMO,SEGCO,SEGCE,OWNER FROM ORRSKCIRCUITO  "
      Strsql = Strsql & " Where codobjeto=" & nCodigo
      Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
      Strsql = Strsql & " And PERIODO = 0" & session("ORPERIODO")
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      WKSPTABLA = "ORRSKCIRCUITO"


   ELSE



      Strsql = "SELECT * FROM OBJSDK  "
      Strsql = Strsql & " Where codobjeto=" & nCodigo
      Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
    ''response.write Strsql & " Strsql " & "<br>"
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if rs.eof then

         strsql = "Insert into OBJSDK (codobjeto, tipobjeto, segct,segmo,segco,segce,segown "
         strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',0,0,0,0,'"
         strsql = strsql & Session("Usuario") & "')"

         call ofv.crearconsultaEx(StrSql,cn,1,volver)

         Strsql = "SELECT * FROM OBJSDK  "
         Strsql = Strsql & " Where codobjeto=" & nCodigo
         Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      end if

      WKSPTABLA = "OBJSDK"

  END IF


  if not rs.eof then

     sOwner = "S/D"
     set cx = ofv.conectar(ofv.strconn0)

     if xopc = "HE" then

       ' Strsql = "SELECT ID,descripcion FROM usuarios"
       ' set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       ' do until rx.eof
       '    sxsel  = ""
       '    if rx(0) = rs(7) then sxsel = " selected "
       '    sxopc1 = sxopc1 & "<option value=" & rx(0) & sxsel & " >" & rx(0) & " - " & rx(1)
       '    sxopc1 = sxopc1 & "</option> "
       ' rx.movenext
       ' loop
       xxxx     = getInfoUser(rs(7))
	     sOwner = session("SGU_nombre")

        sOwner = "&nbsp;&nbsp;" & rs(7) & " - " & sOwner

       ' sInput  = "<select class=at name=" & rs(7).name & " id=" & rs(7).name
       ' sInput  = sInput & " onchange=document.edHE.submit(); >" & sxopc1 & "</select>"
       ' sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver=" & vuelta
       ' sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       ' sOwner  =  ofv.GenForm("edHE","",sAccion,"","post","",sInput,"si")
       ' call ofv.cerrarconsulta(rx)
       ' call ofv.cerrarconn(cx)

         if bModificar  and xopc = "HE"  then
	   	        'sImagenP = "&#0232;"
	   		       'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
	   		       sinput = "<a " & sFuncAnc & " href='WKFREASIGNADAVUSR.asp?HE=S&Cod=" & nCodigo &"'"
	   		       sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
	             sinput = sinput  & "' title='Propiedades de " & rs(1) & "' <i class='las la-arrow-right'></i></a>"
	   	    else
	   	        sInput = "&NBSP;" 'ofv.GenerarInput("",rs(26),"readonly","","10","dt","")
	   	    end if
	  '  Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

        sOwner=sOwner & "&nbsp;&nbsp;&nbsp;" & sinput
     else


        xxxx     = getInfoUser(rs(7))
        sOwner = session("SGU_nombre")

        sOwner = "&nbsp;&nbsp;" & rs(7) & " - " & sOwner


     end if


    Celda = ofv.gencelda("","tt","","","","","","Propietario:","si")
    Celda = Celda & ofv.gencelda("","vti colspan=3","","","","","",sOwner,"si")
    Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda = ofv.gencelda("","ttr colspan=2","","","","","","Restricci&oacute;n","si")
    Celda = Celda & ofv.gencelda("","ttr","","","","","","Activo","si")
    Celda = Celda & ofv.gencelda("","ttr","","","","","","Permisos","si")
    Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")

    sForm  = "Sdk"
    sinput = "Eliminar&nbsp;-&nbsp;Copiar&nbsp;-&nbsp;Reemplazar&nbsp;-&nbsp;"
    sinput = sinput & "Restringir&nbsp;-&nbsp"
    sinput = sinput & "Autorizar:"
    Celda  = ofv.gencelda("","tt colspan=2","","","","","",sinput,"si")

    if bModificar then
       sInput = ofv.combosino(sForm & "1",rs(3),rs(3).name)

       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "1","",sAccion,"","post","",sInput,"si")
       sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
    else
      ACTIVO="NO"
      if rs(3) then activo="SI"
       sinput = ofv.gencelda("","ut","","","","","",ACTIVO,"si")
    end if

    Celda = Celda & sinput

    if xopc = "HE" then
       sInput = ofv.gencelda("","ut","","","","","","---","si")
       Celda = Celda & sinput
    else

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a " & sFuncAnc & " href='edobjsdkdet.asp?cod=1&volver=" & vuelta2
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&circuito=" & rs(0)
    sinput = sinput  & "' title='Lista de autorizados.' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","","","",sInput,"si")

    end if


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt colspan=2","","","","","","Modificacion:","si")

    if bModificar then
       sInput = ofv.combosino(sForm & "2",rs(4),rs(4).name)
   ''''    sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "2","",sAccion,"","post","",sInput,"si")
       sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
    else
       ACTIVO="NO"
      if rs(4) then activo="SI"
      sinput = ofv.gencelda("","ut","","","","","",ACTIVO,"si")
    end if

    Celda = Celda & sinput

    if xopc = "HE" then
       sInput = ofv.gencelda("","ut","","","","","","---","si")
       Celda = Celda & sinput
    else

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a " & sFuncAnc & " href='edobjsdkdet.asp?cod=2&volver=" & vuelta2
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&circuito=" & rs(0)
    sinput = sinput  & "' title='Lista de autorizados.' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","","","",sInput,"si")

    end if


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt colspan=2","","","","","","Cambio de Estado:","si")

    if bModificar then
       sInput = ofv.combosino(sForm & "4",rs(6),rs(6).name)
   ''''    sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "4","",sAccion,"","post","",sInput,"si")
       sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
    else
        ACTIVO="NO"
        if rs(6) then activo="SI"
       sinput = ofv.gencelda("","ut","","","","","",ACTIVO,"si")
    end if

    Celda = Celda & sinput

    if xopc = "HE" then
       sInput = ofv.gencelda("","ut","","","","","","---","si")
       Celda = Celda & sinput
    else


   '''sImagenP = "&#57421;"
    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a " & sFuncAnc & " href='edobjsdkdet.asp?cod=4&volver=" & vuelta2
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&circuito=" & rs(0)
    sinput = sinput  & "' title='Lista de autorizados.' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","","","",sInput,"si")

    end if


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt colspan=2","","","","","","Acceso Consulta:","si")

    if bModificar then
       sInput = ofv.combosino(sForm & "3",rs(5),rs(5).name)
   ''''    sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sInput =  ofv.GenForm(sForm & "3","",sAccion,"","post","",sInput,"si")
       sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
    else
        ACTIVO="NO"
        if rs(5) then activo="SI"
       sinput = ofv.gencelda("","ut","","","","","",ACTIVO,"si")
    end if

    Celda = Celda & sinput

    if xopc = "HE" then
       sInput = ofv.gencelda("","ut","","","","","","---","si")
       Celda = Celda & sinput
    else

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a " & sFuncAnc & " href='edobjsdkdet.asp?cod=3&volver=" & vuelta2
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&circuito=" & rs(0)
    sinput = sinput  & "' title='Lista de autorizados.' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","","","",sInput,"si")

    end if


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    if xopc = "HE" then

       Celda = ofv.gencelda("","tt","","2% colspan=4","","","","&nbsp;","si")
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

       sInput = ofv.GenerarInput("botonx","Volver","submit","vuelvex","5","bt","")
       sInput = ofv.gencelda("","ut","center","2% colspan=4","","","",sInput,"si")
       sAccion = "LisObjHabil.asp"
       sInput =  ofv.GenForm("vuelvex","",sAccion,"","post","",sInput,"si")

       Celda = sinput
       Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    end if


 end if

    Celda = ofv.gencelda("","ut","","10%","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","ut","80%","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","ut","5%","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","ut","5%","","","","","&nbsp;","si")
    Fila  = fila & ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","aaa","","90%' ","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"

 volver = ""

 if (smodalidad & "") = "PRG" OR (smodalidad & "") = "PRGCOR" then

  sAccion = Session("ASPSTRD")
  volver = ofv.BotonVolver(sAccion)

 end if





  sHTML = sHTML & "</center>"
if  xopc <> "HE" then
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,volver)
end if

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>