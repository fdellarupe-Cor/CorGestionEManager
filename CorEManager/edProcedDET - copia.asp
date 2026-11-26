  <!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  response.buffer=true
  APLCHECK()
  dim wstipo(6)
  ofv.ObtenerAtributos "acciones",""
  bagregarbtn=ofv.aagregar

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  if request.querystring("XEVEN") <> "" then
     Session("XEVEN") = request.querystring("XEVEN")
     XEVEN = request.querystring("XEVEN")
  else
     XEVEN = Session("XEVEN")
  end if

  if request.querystring("Opcion") <> "" then
     OpcAgregar = request.querystring("Opcion")
     Fltipo = request.form("Fltipo")
  else
     OpcAgregar = "no"
  end if
  if request.querystring("Opcion2") <> "" then
     OpcCopiar = request.querystring("Opcion2")
  else
     OpcCopiar = "no"
  end if
  if request.querystring("Opcion3") <> "" then
     OpcCopiar2 = request.querystring("Opcion3")
  else
     OpcCopiar2 = "no"
  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  sTest = request.querystring("Test")

  CantRegAMover = 0

  CantRegAMover2 = "0" & request.querystring("CantRegAMover2")

  if CantRegAMover2 = "0" then
     CantRegAMover2 = "0" & request.querystring("CantRegMover")
 ' else
 '    CantRegAMover2 = "0"
  end if

  CantRegAMover2 = cdbl(CantRegAMover2)

  CantRegAMostrar = 1

  If BrowserName() = "IE" OR BrowserName() = "CHROME" then
    CantRegAMostrar2= 8
  else
    CantRegAMostrar2= 8
  end if

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vueltax = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=ddd>"

'-------------------------------------------------


  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
 'response.write strsql & " strsql " & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  Strsql = "SELECT * from csprocmaster "
  Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codproced = '" & XEVEN & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    nID = rs(0)
    nID2 = rs(9)
    sEven = rs(2)

    if not isnull(seven) then sEven = ofv.convertircar(sEven," ","_")



      if wkaplic <> "PL" then

  'Encabezados
  Formulario1 = ofv.Encabezados("Sec","Evento","Sector","Perfil","Frecuencia","Referencia")
  formulario = ""




      'Secuencia evento
      Celda =  ofv.GenCelda("","ttt","","1%","","","",rs(8) & ".","si")

      'Secuencia procedimiento
      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text","edProcedimientos","2","dtn","SI")
      else
 '       sInput = ofv.GenerarInput("",rs(7),"readonly","","2","","SI")
        sInput = rs(7)
      end if
    Celda = Celda & ofv.GenCelda("","vti","","1%","","","",sInput,"si")

    'Evento

      sEvento = ""
      if rs(4) <> "" then
        Strsql = "Select descri From Objeven Where objid = " & rs(4)
        set rsEvento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsEvento.eof then sEvento = rsEvento(0)
        call ofv.cerrarconsulta(rsEvento)
      end if
 '     sInput = ofv.GenerarInput("",sEvento,"readonly","","40","","")
      sInput = sEvento
 '   end if
    Celda = Celda & ofv.GenCelda("","ttt","","15%","","","",sInput,"si")

    'Sector
    if bModificar then

      IF session("EMPMUN") = "S" then

	    'strsql = "SELECT a.*, c.descri,b.cliente FROM  OBJSECTOR a "
	    'strsql = strsql & " inner join ECO c on (a.codsector=c.eco) "
	    'strsql = strsql & " inner join clienteplan b on (b.ID=c.ecoUN) "
	    'strsql = strsql & " where a.codsector = c.eco "
	    'strsql = strsql & " and c.estado=3 and a.codobjeto=18035 "
	    'strsql = strsql & " And (calificador = 'O' or calificador = 'C') "

		Strsql = "Select objsector.*,Eco.descri,a.cliente From Objsector, Eco  "
		strsql = strsql & " inner join clienteplan a on (eco.ecoUN=a.id) "
		strsql = strsql & " where eco.estado=3 and "
        Strsql = Strsql & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo

	  else

	     Strsql = "SELECT OBJSECTOR.*, ECO.descri FROM OBJSECTOR, ECO "
	      Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and codobjeto=" & nCodigo
	    Strsql = Strsql & " And (calificador = '" & sCalificador & "' or calificador = 'C' or calificador = 'O') "

	end if
'response.write strsql
      Set rsSectorn = ofv.crearconsultaEx(strsql, cn, 1, parametros)
        'sInput = ofv.GenerarCombo("edProcedimientos",rs(6).name,rs(6).value,Strsql,cn,0,1,"sb")

            scomK = ""'"null" & Chr(9)
        scomD = "" '"" & Chr(9)
        SCOMSEP =""
      do while not rsSectorn.eof
         scomK=scomK & SCOMSEP & rsSectorn(0)
         IF session("EMPMUN") = "S" then
          scomd=scomd & SCOMSEP & rsSectorn(14) & "&nbsp;-&nbsp;" & rsSectorn(13)
         else
          scomd=scomd & SCOMSEP & rsSectorn(13)
         end if

         SCOMSEP = chr(9)
         rsSectorn.movenext
        loop
            sInput = ofv.gencomboFX("edProcedimientos",RS(5).Name,RS(5),scomK,scomD,"")
    else
      sSector = ""
      if rs(5) <> "" then

        Strsql = "Select descri From Eco, Objsector where "
        Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.objid = " & rs(5)
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then sSector = rsSector(0)
        call ofv.cerrarconsulta(rsSector)
      end if
'      sInput = ofv.GenerarInput("",sSector,"readonly","","40","","")
      sInput = sSector
    end if
    Celda = Celda & ofv.GenCelda("","tu","","20%","","","",sInput,"si")

    'Perfil
    if bModificar then

      Strsql = "SELECT TIPOPERFIL.tipoperfil,TIPOPERFIL.descri "
      Strsql = Strsql & "FROM Objsector, ECO, ECOPERFIL, TIPOPERFIL "
      Strsql = Strsql & " where Eco.eco = Objsector.codsector and "
      Strsql = Strsql & " ECO.eco = ECOPERFIL.codsector and TIPOPERFIL.estado=3 and"
      Strsql = Strsql & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
      Strsql = Strsql & " and objsector.objid = " & rs(5).value
      Set rsPerfil = ofv.crearconsultaEx(strsql, cn, 1, parametros)
      'sInput = ofv.GenerarCombo("edProcedimientos",rs(6).name,rs(6).value,Strsql,cn,0,1,"sb")
       scomK = ""'"null" & Chr(9)

	    scomD = "" '"" & Chr(9)
	  	SCOMSEP =""
		  do while not rsperfil.eof
		   scomK=scomK & SCOMSEP & rsperfil(0)
		   scomd=scomd & SCOMSEP & rsperfil(1)
		   SCOMSEP = chr(9)
		   rsPerfil.movenext
	      loop
          sInput = ofv.gencomboFX("edProcedimientos", RS(6).Name, RS(6), scomK, scomD, "")
    else
      sPerfil = ""
      if rs(6) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then sPerfil = rsPerfil(0)
        call ofv.cerrarconsulta(rsPerfil)
      end if
 '     sInput = ofv.GenerarInput("",sPerfil,"readonly","","40","","")
      sInput = sPerfil
    end if
    Celda = Celda & ofv.GenCelda("","tu","","20%","","","",sInput,"si")

    'FRECUENCIA
    if bModificar then
      Strsql = "Select * From frecuencia where estado=3  Order By descri"
    ''sInput = ofv.GenerarCombo("edProcedimientos",rs(11).name,rs(11).value,Strsql,cn,0,1,"sb")

       Set rsfrecx = ofv.crearconsultaEx(strsql,cn,1,parametros)

         scomK = ""'"null" & Chr(9)
         scomD = "" '"" & Chr(9)
         SCOMSEP =""
       do while not rsfrecx.eof
          scomK=scomK & SCOMSEP & rsfrecx(0)
          scomd=scomd & SCOMSEP & rsfrecx(1)
          SCOMSEP = chr(9)
          rsfrecx.movenext
       loop

        sInput = ofv.gencomboFX("edProcedimientos",rs(11).name,rs(11).value,scomK,scomD, "")
    else
      sfrec = ""
      if not isnull(rs(11))  then
        Strsql = "Select descri From frecuencia Where codfrec = '" & rs(11) & "' "
        set rsfrec = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsfrec.eof then sfrec = rsfrec(0)
        call ofv.cerrarconsulta(rsfrec)
      end if
  '    sInput = ofv.GenerarInput("",sfrec,"readonly","","40","","")
      sInput = sfrec
    end if

    Celda = Celda & ofv.GenCelda("","tu","","20%","","","",sInput,"si")


      'Texto de Referencia
      sID = "edProcedRef.asp?"
      sID = sID & "ID=" & rs(0) & "&Tabla=Objproced&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHref = sID

      if rs(12) <> "" then
         xclas  = "btbs"
       '  if BrowserName()="IE" OR BrowserName()="CHROME" THEN
       '  sInput = "<font class='btic' >4&#0157;</font>"
       '  ELSE
         sInput = "<font class='btic' >4&#0157;</font>"
      '   END IF
      else
         xclas  = "btbn"
   '      if BrowserName()="IE" OR BrowserName()="CHROME" THEN
      '   sInput = "<font class='btic' >4&#0157;</font>"
      '    ELSE
		          sInput = "<font class='btic' >4&#0157;</font>"
       '  END IF
      end if

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      stitle = "Actualizar referencia del circuito."
      XXX= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
       Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","center","5%","","","",XXX,"si")

      fila = ofv.GenRow("","","","","","","",celda,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID2 & "&Tabla=Objproced&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      tabla= ofv.GenTabla("","","","100%","","0","","-1","2","","",formulario1 & fila,"si")
      Formulario =   ofv.GenForm("edProcedimientos","",sAccion,"","post","",tabla,"si")
      'Formulario = Formulario & ofv.GenForm("edProcedimientos","",sAccion,"","post","",Celda,"si")
    else
      tabla= ofv.GenTabla("","","","100%","","0","","-1","2","","",formulario1 & fila,"si")
      Formulario =  ofv.GenForm("edProcedimientos","",sAccion,"","post","",tabla,"si")
      'Formulario = Formulario & ofv.GenForm("edProcedimientos","","","","post","",Celda,"si")
    end if

    'fila = ofv.GenRow("","","","","","","",celda,"si")
    'Formulario = ofv.GenRow("","","","","","","",Formulario,"si")

   else

      wksecuencia = rs(8) & "." & rs(7)


    'Evento

      sEvento = ""
      if rs(4) <> "" then
        Strsql = "Select descri From Objeven Where objid = " & rs(4)
        set rsEvento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsEvento.eof then sEvento = rsEvento(0)
        call ofv.cerrarconsulta(rsEvento)
      end if

    'Sector

      sSector = ""
      if rs(5) <> "" then

        Strsql = "Select descri From Eco, Objsector where "
        Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.objid = " & rs(5)
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then sSector = rsSector(0)
        call ofv.cerrarconsulta(rsSector)
      end if


    'Perfil

      sPerfil = ""
      if rs(6) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then sPerfil = rsPerfil(0)
        call ofv.cerrarconsulta(rsPerfil)
      end if

    wkprocdesc = wksecuencia & "." & sEvento & "-"  & sSector & "-"  & sPerfil

    Formulario1 = ""
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Proceso: ","si")
    Celda = Celda & ofv.GenCelda("","ttt","","85%","","","",wkprocdesc,"si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")



   end if



    rs.movenext

  else
    nID = "no"
  end if

  ''sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","-1","2","","",formulario1 & Formulario,"si")
  sHTML = sHTML & Formulario


  response.write sHTML
  sHTML = ""

' Acciones del procedimiento

 if OpcAgregar = "si" or OpcAgregar = "fl"  then

    Strsql = "Select * From Modo Order By descri"

    if Fltipo &"" = "" then

       set rsFiltro = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsFiltro.eof then Fltipo = rsfiltro(0)

    end if

    response.write "<br>"
    Celda = ofv.GenCelda("","ttt","","","","","","Agregar Paso","si")
''''Fila  = ofv.GenRow("","","","","","","",Celda,"si")

''''Nuevo Agregado
    'Celda  = Celda & ofv.gencelda("","ttt","","10%","","","","Tipo de Accion","si")



IF session("APLMDCLI") <> "BNX" AND session("APLMDCLI") <> "GP"  and false THEN
   scomKK = ""
    scomDD = ""
  '  response.write strsql & "<br>" & chr(13)
    set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    DO UNTIL rsniv.eof
       scomKK = scomKK & rsniv(0) & chr(9)
       scomDD = scomDD & rsniv(1) & chr(9)
       RSNIV.MOVENEXT
    LOOP
    call ofv.cerrarconsulta(rsniv)

  sInput = ofv.gencomboFX("Flform","Fltipo",Fltipo,scomKK,scomDD,"")

' sInput = ofv.GenerarCombo("Flform","Fltipo",Fltipo,Strsql,cn,0,1,"si")

    Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    'Celda  = Celda & ofv.gencelda("","vti","","","","","","&nbsp;","si")
    Fila   = ofv.genrow("","","","","","","",Celda,"si")
''''Nuevo Agregado Fin

    sTabla = ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  sAccion = "edProcedDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	sAccion = sAccion & "&Opcion=si&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
	sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

	sHTML = ofv.GenForm("Flform","",sAccion,"","post",sFuncion,sTabla,"si")
 ELSE
       Fila  = ofv.GenRow("","","","","","","",Celda,"si")
      sHTML = ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")
      response.write sHTML
      sHTML = ""
  END IF
    response.write sHTML
    sHTML = ""

 else

  'Encabezados
      if wkaplic <> "PL" then

 'Agregar Accion

'' if bModificar then
''    sInput  = ofv.GenerarInput("","Agregar Accion","submit","","20","bt","")
''    Celda   = "" 'ofv.GenCelda("","vt title='Agregar nueva accion a la lista.' ","","1%","","","",sInput,"si")
''    sAccion = "AgregarAcciones.asp?tabla=Tipoaccalt&Opcion=si&volver=" & vueltaC
''    sAccion = sAccion & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
''    test    = ofv.GenForm("agracc","",sAccion,"","post","",sInput ,"si")
''    test= ofv.GenCelda("","vt title='Agregar nueva accion a la lista.' ","","1%","","","",test,"si")
''   ' test= test & "</form>"
'' end if

  sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

 ' sql = " select atributo from atrmUsuarios where usuario='" & session("Usuario") & "' and atributo ='19-COREM20-M-ACCIONES'"
 ' set rf = ofv.crearconsultaex(sql, cn, 1, parametros)

 ' permiso=""
 ' if rf(0) = "19-COREM20-M-ACCIONES" then
 ' permiso = false
 ' else
 ' permiso = true
 ' end if

''INC 1355
  if bModificar  then
      sInput  = ofv.GenerarInput("","Agregar Accion","submit","","20","bt","")
      sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Opcion=si&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

      test    = ofv.GenForm("agracc","",sAccion,"","post","",sInput ,"si")

 '''''test= ofv.GenCelda("","vt title='Agregar nueva accion a la lista.' ","","1%","","","",test,"si")
      test= ofv.GenCelda("","","","1%","","","",test,"si")

  end if


 'Cursograma
    sInput = ofv.GenerarInput("","Cursograma","submit","","20","bt","")
    Celda = ""'ofv.GenCelda("","vt title='Abrir grafico del cursograma.' ","","1%","","","",sInput,"si")
    ' if session("APLMDCLI") <> "BC" then
    '    sAccion = "Cursograma.asp?volver=" & vuelta & "&Codproced=" & nID & "&Tipobjeto=" & sTipo
    ' else
        sAccion = "CursogramaAdmin.asp?Codobjeto=" & nCodigo & "&Codproced=" & nID & "&Tipobjeto=" & sTipo
    ' end if
    Formularioc = ofv.GenForm("Cursograma","",sAccion,"Grama","post","",sInput,"si")
    Formularioc=ofv.GenCelda("","vt title='Abrir grafico del cursograma.' ","","1%","","","",Formularioc,"si")
''''
 'Importar procedimiento
    Formularioi = "&nbsp;"
    if bmodificar then
      sInput = ofv.GenerarInput("","Importar","submit","","10","bt","")
      Celda = "" 'ofv.GenCelda("","vt title='Incorporar pasos desde otro Circuito.' ","","1%","","","",sInput,"si")
      sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Opcion=no&Opcion2=si&Opcion3=no&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2
      Formularioi = ofv.GenForm("Importar","",sAccion,"","post","",sInput,"si")
      Formularioi=ofv.GenCelda("","vt title='Incorporar pasos desde otro Circuito.' ","","1%","","","",Formularioi,"si")
    end if
''''



         Celda = ofv.GenCelda("","vtig","","90%","","","","Detalle del Circuito ","si")
         Celda = celda & indice
         Celda = celda & test
         Celda = celda & Formularioc
         Celda = celda & Formularioi
      else
         Celda = ofv.GenCelda("","vtig","","90%","","","","Actividades del proceso para contingencia ","si")

      end if



  Fila = ofv.GenRow("","","","","25","","",Celda,"si")
  sHTML = "<br>" & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;background-color:#F9F9F9;' ","","100%","","0","","0","0","","",Fila,"si") & "<br>"
  response.write sHTML
  sHTML = ""
 end if

  sHTML = ""
  response.flush

 if OpcAgregar = "si" then

 else



  bColor = "False"
  if nID <> "no" then



    strsql = "select OBJACCION.*, b.codinx2 OBJOBSCMT  from OBJACCION "

    'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (OBJACCION.objid = b.codinx2) "
    'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (OBJACCION.objid = b.codinx2) "
    strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast ( OBJACCION.objid as nvarchar) = b.codinx2) "
    strsql = strsql & " WHERE codproced = '" &  nID & "'"
    if bVersion = "NO" then
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    end if
    strsql = strsql & " Order By secuen"

   'response.write strsql & "<br>" & chr(13)

    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then rs2.move(CantRegAMover2)
    posicion = 0


    Filam = ""
    xTrColor = "#F9F9F9"
'---
     sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"
     sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#4;</font>"
      wktitdelete = "Confirma la eliminacion???."
    do while not rs2.eof and posicion < CantRegAMostrar2
      nAccion   = rs2(0)
      nAccionID = rs2(18)
      nselec    = "Expr2"
      sForm     = "edProcedimientos" & rs2(0)

      'Eliminar
      sHref = ""
      if rs2(14) And rs2(15) = "E"  and (rs2(20) & "") <> "PL" then
 ''     sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        sImage = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
        Celda = ofv.GenCelda("","vt","","2","","","",sImage,"si")
      else
        if bModificar  and (rs2(20) & "") <> "PL" then
          if rs2(14) And rs2(15) = "A" then
            sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
           ' sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          else
            sHref = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI" '& chr(34) & " onclick="& chr(34) & "alert(1);"
          end if
        end if

        if sHref <> ""  and (rs2(20) & "") <> "PL" then
         ' CeldaZ = ofv.BotonEliminar("Accion,"qqq")
 ''       sImageEL   =ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
          sImage = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
		   sformobj   = "FRMOBJ" & rs2(0)
		  ' sAccion    = "eliminarproceso.asp?codobjeto=" & rs(0) & "&tipobjeto=" & rs(20) & "&OPCION=ACC"
		   formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sHref,"","post","","","si")
		   sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "
        '  celda= ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageEL ,"si") & formobj
          Celda = ofv.BotonEliminar("Accion ",sHref)
         ' response.write celda & "2222222222222222222"
           'celda=replace(celda,"onclick=" & chr(34) & "document", "onclick=""var r=confirm('"&wktitdelete&"');if (r==true) {document")
           celda=replace(celda,"onclick='document", "onclick=""var r=confirm('"&wktitdelete&"');if (r==true) {document")
           celda=replace(celda,".submit();'", ".submit();} else {alert('Se cancelo la accion!');}" &chr(34))
         ' response. write ofv.BotonEliminar("AcCIqqqqqqqON ",sHref)
        else
            Celda = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")
        end if
      end if


      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;" ,"si")


     'response.write nAccion & " nAccion " & "<br>"
     'response.write nCodigo & " nCodigo " & "<br>"
     'response.write sTipo   & " sTipo " & "<br>"


''''' agregado para color del cuadrito de Observaciones version 4.6
      wComentario = ""
      strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codinx2 = " & nAccion & " and "
      strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"      ' and "

     ' strsql = strsql & " AND codcircuito='"& IDC &"'"
      strsql = strsql & " order by codcircuito, codusr, objid"
'	response.write strsql
	  set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rss.eof then
         wComentario = rss(13)
      end if
      call ofv.cerrarconsulta(rss)
''''''

      if (Nestado=5  or Nestado=6) and (rs2(30) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion

  ''''''agregado para color del cuadrito de Observaciones version 4.6

               if wComentario = "" then
      	          sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
       	       else
       	          sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
       	       end if
  ''''''
              Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")
      	      Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","", Celdaobs,"si")

      end if


      Celda = Celda & Celdaobs


      'Secuencia
      sInput = "XX" & rs2(10)

      if len(sInput) > 3 then sinput = right(sinput,3)
      sinputS = replace(sinput,"X","&nbsp;")
   '   Celda = Celda & ofv.GenCelda("","vt","","2","","","",sInput,"si")

      'Acciones


        sInput1 = ""
        sAccion = " "
        sModo   = " "
        sArch   = "NO"
        if rs2(4) <> "" then
          Strsql = "Select descri,modo From Tipoaccalt Where tipoaccionalt = " & rs2(4)
          set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsAcciones.eof then
             sAccionM = rsAcciones(0)
             sModo   = rsAcciones(1)
          END IF
          call ofv.cerrarconsulta(rsAcciones)
        end if


      westado = "S"

      for it=1 to 6
          wstipo(it) = "N"
      next


      strsql12 = "SELECT * from modo "
      strsql12 = strsql12 & "Where tipomodo = '" & sModo & "' "
      scontrol=""
      set rs12 = ofv.crearconsultaEx(StrSql12,cn,1,volver)
      if not rs12.eof then
         sArch   = rs12(2)
         scontrol=rs12("control")
         for it=1 to 6
             if rs12(it+3) then
                wstipo(it) = "S"
             end if
         next
      end if
      call ofv.cerrarconsulta(rs12)


      if wkaplic <> "PL" then


      'Icono
      if sArch <> "NO" then
        sImage = ofv.GenImage("","","","","Iconos/" & sArch,"0","")
      else
        sImage = ofv.GenImage("","","","","Iconos/ICorAImagen.jpg","0","")
      end if
     ' sLink =  ofv.GenerarInput(rs2(10).name,rs2(10),"text",sForm,"10","dtn","SI")' ""'sImage
      sAccionO = "asp/ActualizarEN.asp?ID=" & RS2(0) & "&Tabla=OBJACCION&volver=" & vuelta
       sAccionO = sAccionO & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sLink = "<form method=post action='" & sAccionO & "' name=tobjcc2" & rs2(0) & " id=tobjcc2" & rs2(0) &  ">" & ofv.GenerarInput("secuen",rs2("secuen"),"text","tobjcc2" & rs2(0),"10","dtn","") & "</form>"
      Celda = Celda & ofv.GenCelda("","vt","","1%","","","",sLink,"si")

      end if



      sVerbo = sAccionM
      sAcciont = sAccionM
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccionM,i-1)
        sAcciont = mid(sAccionM,i+1)
      else
        sAcciont=""
      end if
'response.write  sAcciont & "--" & i

      if wkaplic <> "PL" then

        ' sInput = "-" & sInputS & ".&nbsp;<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
         sInput =  sVerbo  & "&nbsp;&nbsp;&nbsp;"
         sInput = sInput & replace(sAcciont," ","&nbsp;") & ""

         sID = "EdProcedDetProp.asp?"
         sID = sID & "AccId=" & nAccion & "&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&CantRegMover=" & CantRegAMover2

         sHref = sID
         if rs2(4)=a301 then
          sAccionO = "asp/ActualizarEN.asp?ID=" & RS2(0) & "&Tabla=OBJACCION&volver=" & vuelta
          sAccionO = sAccionO & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
          t3b = "<font class=at><textarea class=dt STYLE='FONT-SIZE:12PT;BACKGROUND-COLOR:#EEEEFF;' "
		  t3b = t3b & "name=" & "descriacc" & " cols=90 rows=2 "

		  t3b = t3b & " onblur = " & Chr(34) & " this.value = this.value.replace(/^\s+/, '');  if (this.value.length > 0) {document.getElementById('" & "tobjcc" & rs2(0) & "').submit() } else {alert('no puede estar en blanco.')}" & Chr(34)
		  t3b = t3b & " onkeypress = " & Chr(34)
		  t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
		  t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"

		'   else
		'      t3b = t3b & " readonly >"
		'   end if
		   xdatos =  rs2("descriacc")
		   t3b    = t3b & xdatos & "</textarea></font>"


         ' sInput =  "<form method=post action='" & sAccionO & "' name=tobjcc" & rs2(0) & " id=tobjcc" & rs2(0) &  ">" & ofv.GenerarInput("descriacc",rs2("descriacc"),"text","tobjcc" & rs2(0),"40","dt","") & "</form>"
         sInput =  "<form method=post action='" & sAccionO & "' name=tobjcc" & rs2(0) & " id=tobjcc" & rs2(0) &  ">" & t3b & "</form>"

         else
         sInput = ofv.GenLink("","nt title='Propiedades del Paso.' ",sHref,"","","",sInput,"si")
         end if

      else

         sInput = "-" & sInputS & ".&nbsp;" & sVerbo & "" & "&nbsp;&nbsp;&nbsp;"
         sInput = sInput & replace(sAcciont," ","&nbsp;") & ""

      end if


      Celda  = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")




      if wkaplic <> "PL" then

  'Relacion entre eventos (conectores)


         sID = "EdProcedrel.asp?"
         sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo



      if not isnull(rs2(12))  then
         if  clng(rs2(12)) > 0 then
             xclas  = "btbs"
             sInput = "<font class='btic' >&#060;</font>"
         else
             xclas  = "btbn"
             sInput = "<font class='btic' >&#060;</font>"
         end if
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#060;</font>"
      end if


      sHref = sID

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      sInput = "<div class=" & xclas & " align=middle>" & sinput & "</div>"

      stitle = "Relacionar evento con conectores."
       XXX=ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",XXX,"si")
     ' Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      end if

  'Texto de referencia

      sID = "TextoEsquema.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
      sID = sID & "_Tipobjeto=" & sTipo



      if rs2(11) <> "" or rs2(16) <> "" or rs2(17) <> "" then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0165;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0165;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " ALIGN=MIDDLE>" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar texto de referencia."
      xxx=ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","", xxx,"si")
   '   Celda = Celda &


      if wkaplic <> "PL" then

  'Link a paginas internas/externas


         sID = "EdGENLnk.asp?"
         sID = sID & "ID=" & nAccion & "&TABLA=OBJACCION&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo



      if wstipo(3) = "S" then
         if not isnull(rs2(19))  then
            if  clng(rs2(19)) > 0 then
                xclas  = "btbs"

                 sInput = "<font class='btic' >&#0126;</font>"

            else
                xclas  = "btbn"

				 sInput = "<font class='btic' >&#0126;</font>"

            end if
         else
            xclas  = "btbn"

				sInput = "<font class='btic' >&#0126;</font>"
            END IF


         sHref = sID

      else
         sHref  = "javascript:void(0);"
         xclas  = "ut"
         sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
      end if

      sInput = "<div class=" & xclas & " ALIGN=MIDDLE>" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Actualizar link a otro documento."
      xxx=ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")
     ' Celda = Celda &

   end if



  'Asociacion de Complementos

      sID = "EdProcedCMP.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
      sID = sID & "_Tipobjeto=" & sTipo


         sID = "EdGENLnkCOMP.asp?"
         sID = sID & "ID=" & nAccion & "&TABLA=OBJACCION&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo



      if cstr(rs2(6) & "") <> "0" then
         xclas  = "btbs"


         sInput = "<font class='btic' >&#0163;</font>"

      else
         xclas  = "btbn"

          sInput = "<font class='btic' >&#0163;</font>"

      end if

       sHref = sID
      sInput = "<div class=" & xclas & " ALIGN=MIDDLE >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Asociar componente."
      xxx= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")
     ' Celda = Celda &

      if wkaplic <> "PL" then

  'Copias de formularios

      sInput1  = ""
      xfuncion = ""

      if wstipo(2) = "S" and cstr(rs2(6) & "") <> "0" then
         nCopias = 0
         xclas  = "btbs"
         if not isnull(rs2(13)) then nCopias = rs2(13)
         if cint(nCopias) = 0 then
            xclas   = "btbn"
            westado = "N"
            wntexto = "ACCION RELACIONADA A EMISION/RECEPCION DE DOCUMENTOS: No tiene informado la cantidad de copias a administrar!."
         end if





         if bModificar then



            sForm = "FACC" & rs2(0)

            xfuncion = " onclick=""window.top.xopcom('" & rs2(13).name & "','" & rs2(13).name & "2','" & sForm & "','3'); return false;"" "

            SCOMSEP = ""
            scomK   = ""
            scomD   = ""

            For i=0 to 9
                scomK = scomK & SCOMSEP & i
                scomD = scomD & SCOMSEP & i
                SCOMSEP = chr(9)
            next

            sInput1 = ofv.gencomboFX(sForm,rs2(13).name,rs2(13),scomK,scomD,"")




            sAccion = "asp/ActualizarEN.asp?ID=" & rs2(0) & "&Referencia=SI&Tabla=Objaccion&volver=" & vueltaC
            sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
            sInput1  = ofv.GenForm(sForm,"aaa STYLE='display:none;' ",sAccion,"","post","",sInput1,"si")


         end if

         sInput = "<font  class='btic2' " & xfuncion & " title='Cantidad de copias: " & nCopias & ".' style='cursor:pointer;' >" & nCopias & "</font>"


      else
         xclas  = "ut"
         sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
      end if

      sHref  = "javascript:void(0);"

      sInput = "<div class=" & xclas & " style='width:14;' ><center>&nbsp;" & sinput & "&nbsp;</center></div>" & sInput1

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" " & xfuncion & " "
      stitle = "Cantidad de copias: " & nCopias & "."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")


      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")





  'Procedimiento alternativo (control)
     'Acciones alternativas
      session("volverproc") = ""

      sID = "EdProcedAlt.asp?"
      sID = sID & "Accion=" & nAccionID & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
      sID = sID & "_Tipobjeto=" & sTipo

      if wstipo(1) = "S" then
        strsql11 = "SELECT idaccion from accionmaster2 "
        strsql11 = strsql11 & "Where idaccion = '" & nAccionID & "'"
        if bVersion = "NO" then
           strsql11 = strsql11 & " and (modi = 0 or (modi <> 0 and "
           strsql11 = strsql11 & "tipmodi <> 'E')) "
        end if
        set rs11 = ofv.crearconsultaEx(StrSql11,cn,1,volver)
        if not rs11.eof then
           xclas  = "btbs"
           sInput = "<font class='btic'  >2</font>"
        else
           westado = "N"
           wntexto = "ACCION DE CONTROL: No tiene acciones alternativas informadas!."
            xclas  = "btbn"
           sInput = "<font class='btic'  >2</font>"
        end if
        call ofv.cerrarconsulta(rs11)

        sHref = sID

      else
         sHref  = "javascript:void(0);"
         xclas  = "ut"
         sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
      end if

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer procedimiento alternativo (por salida negativa del control)."
             'sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      'Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
                  xxx=ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")

  'Continuador


      'Sector Continuador

      sID = "EdProcedCON.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
      sID = sID & "_Tipobjeto=" & sTipo

      if  cstr(rs2(8)) <> "0"  then
          xclas  = "btbs"


          sInput = "<font class='btic' >G</font>"

      else
          xclas  = "btbn"

          sInput = "<font class='btic' >G</font>"


      end if
       sHref = sID

      sInput = "<div class=" & xclas & " ALIGN=MIDDLE>" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer Sector Continuador."
      xxx=ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")
      'Celda = Celda &

      'perfil Continuador

      if  cstr(rs2(8)) <> "0"  then
          sID = "EdProcedPER.asp?"
          sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
          sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
          sID = sID & "_Tipobjeto=" & sTipo

          sper = rs2(9)

          if isnull(sper) then
             xclas  = "btbn"
             sInput = "<font class='btic' >&#0128;</font>"
          elseif sper = " " then
             xclas  = "btbn"
             sInput = "<font class='btic' >&#0128;</font>"
          else
             xclas  = "btbs"
             sInput = "<font class='btic' >&#0128;</font>"
          end if

          sHref = sID

      else
          sHref  = "javascript:void(0);"
          xclas  = "ut"
          sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
      end if


      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer Perfil Continuador."
      xxx= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")
     ' Celda = Celda &

  'Observaciones


      if westado = "N" then
          xclas  = "btbn"

          sInput = "<font class='btic'  >=</font>"


      else

           sInput = "<font class='btic'  >=</font>"

          xclas  = "btbs"
          wntexto = "Paso correcto - sin observaciones."
      end if

      sHref  = "javascript:void(0);"
      sInput = "<div class=" & xclas & " ALIGN=MIDDLE>" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Ver observaciones."
      xxx= ofv.GenLink("",xclas & " title='" & stitle & "' onclick=" & chr(34) & "alert('" & wntexto & "');"  & chr(34),sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")
     ' Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' onclick=" & chr(34) & "alert('" & wntexto & "');"  & chr(34),sHref,"","","",sInput,"si")

   end if

  if wkaplic = "PL" then

     'Contingencias para la actividad



      sID = "EdProcedCONtingencias.asp?"
      sID = sID & "Accion=" & nAccionID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
      sID = sID & "_Tipobjeto=" & sTipo

      xclas  = "btbs"
      sInput = "<font class='btic' >G</font>"

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Contingencias para la actividad."
    '     '  sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
   ''   Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
                 xxx= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")

   end if



  'Modi (actualizaciones por versionado)

       'Version
      sSelect = ""
      if bVersion = "SI" then
         if rs2(14) then

              if rs2(15) = "A" then
                 stitle = "Paso agregado."
                 xclas  = "btba"
                 sInput = "<font class='btic'  >6</font>"
              elseif rs2(15) = "E" then
                 stitle = "Paso eliminado."
                 xclas  = "btbn"
                 sInput = "<font class='btic'  >5</font>"
              elseif rs2(15) = "M" then
                 stitle = "Paso modificado."
                 xclas  = "btbs"
                 sInput = "<font class='btic'  >4</font>"
              end if

         else
              stitle = "Paso sin modificaciones."
              xclas  = "ut"
              sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
         end if
      else
         stitle = "Paso original."
         xclas  = "ut"
         sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
      end if

      sHref  = "javascript:void(0);"
      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"
            'sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      'Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
                 xxx= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")



  'Libre
      sInput = "&nbsp;"
      Celda = Celda & ofv.GenCelda("","ut","","2%","","","",sInput,"si")



      FilaM  = FilaM & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","10","","",Celda,"si")

      if xTrColor = "#FFFFFF" then
         xTrColor = "#F9F9F9"
      ELSE
         xTrColor = "#FFFFFF"
      END IF


      Celda = ""



      rs2.movenext
      posicion = posicion + 1
    loop

      Tabla = ofv.GenTabla("","aaa style={border-bottom-style:solid;border-width:1;border-color:#cccccc;} ","","95%","","0","","-1","1","","",Filam,"si")
      response.write Tabla

  end if

 end if




  'Boton test
   sTest = "T"
'  sAccion = "CoProcedimientos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
'  sAccion = sAccion & "&Menu=PORESTADO&Sector=3&origen=ED&CantRegAMover=A&proced=" & XEVEN
  sAccion= " Consultar.asp?Codigo=" & nCodigo & "&Tipo=" & STIPO &"&niv1=5&niv2=" & nID & "&test=" & sTest
  TEST = ofv.BotonTest(sAccion)

  'Boton Volver
  sAccion = "edProcedimientos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Indice = ofv.BotonVolver(sAccion)



  if OpcCopiar = "si" then
    Celda = ofv.GenCelda("","tt","","5%","","","","Procedimiento","si")

    Strsql = ""
    Strsql = Strsql & "SELECT a.codobjeto,a.descri from Objetos a "
    Strsql = Strsql & "   inner join tipobjeto b on a.tipobjeto = b.tipobjeto "
    Strsql = Strsql & "   inner join objperfil c on b.perfil = c.tipoperfil "
    Strsql = Strsql & "   left join objperfildoc d on a.codobjeto = d.codobjeto "
    Strsql = Strsql & " Where c.objproced = 1 or d.objproced = 1 "
    Strsql = Strsql & " order by a.descri, a.codobjeto "



  '  Strsql = "SELECT * from csProcObjetos Where tipobjeto = '" & sTipo & "' "
   '  Strsql = Strsql & "Order By codiden"
    formulario=""'Agregado PDC'


         Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
       scomK = ""'"null" & Chr(9)

      scomD = "" '"" & Chr(9)
      SCOMSEP =""
      do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
       scomd=scomd & SCOMSEP & rsperfil(1)
       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop
 formulario=""
      sInput =  ofv.gencomboFX("Ok", "selproc", "", scomK, scomD, "")
    'sInput = ofv.GenerarCombo("Ok","selproc","",Strsql,cn,0,1,"no")
    'Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")

   ' sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
   ' Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = sAccion & "&Opcion=no&opcion2=no&opcion3=si&CantRegAMover=" & CantRegAMover
    sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

    fila22= ofv.GenForm("Ok","",sAccion,"","post","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",fila22,"si")
    tabla22 = ofv.GenTabla("","","","100%","","0","","0","0","","",Celda,"si")
    Formulario = Formulario & tabla22 ' ofv.GenForm("Ok","",sAccion,"","post","",tabla22,"si")



    sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = sAccion & "&CantRegAMover=" & CantRegAMover
    sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

 '   sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  '  sAccion = sAccion & "&Opcion=no&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
  '  sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

    sInput = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
    fila23= ofv.GenForm("Cancel","",sAccion,"","post","",sInput,"si")
    Celda = ofv.GenCelda("","ttr","","","","","",fila23,"si")
    tabla23 = ofv.GenTabla("","","","100%","","0","","0","0","","",Celda,"si")
    formulario= Formulario & tabla23 ' ofv.GenForm("xCancel","",sAccion,"","post","",tabla23,"si") 'ARREGLO A MEDIAS v1


    Fila = ofv.GenRow("","","","","","","",Formulario,"si")
   '' sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")
   sHTML = sHTML & formulario
  elseif OpcCopiar2 = "si" then
    Celda = ofv.GenCelda("","ttr","","5%","","","","Evento","si")


    Strsql = "SELECT codproced, nombre "
    Strsql = Strsql & " from CSCEventoSectorPerfil Where Obcodobjeto = " & request.form("selproc")
    Strsql = Strsql & " order by secuen, secuencia, codproced "
   'response.write "strsql " & strsql
  ''  sInput = ofv.GenerarCombo("Ok","selcom","",Strsql,cn,0,1,"no")
   '' sInput = ofv.GenerarCombo(sForm,rs(11).name,VALOR,Strsql,cn,0,10,"")
''   response.write StrSql

         Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
       scomK = ""'"null" & Chr(9)

      scomD = "" '"" & Chr(9)
      SCOMSEP =""
      do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
       scomd=scomd & SCOMSEP & rsperfil(1)
       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop
 formulario=""
      sInput =  ofv.gencomboFX("Ok", "selcom", "", scomK, scomD, "")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")

    sInput = ofv.GenerarInput("","Aceptar","submit","","10","bt","")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sAccion = "asp/Importar.asp?volver=" & vueltaC & "&Codobjeto=" & nCodigo & "&Tipobjeto="
    sAccion = sAccion & sTipo & "&Codproced=" & nID
    fila28= ofv.GenForm("ok","",sAccion,"","post","",Celda,"si")
    tabla28 = ofv.GenTabla("","","","100%","","0","","0","0","","",fila28,"si")
  formulario= Formulario & ofv.GenForm("Ok","",sAccion,"","post","",tabla28,"si")
  ''    Formulario = formulario & ofv.GenForm("Ok","",sAccion,"","post","",Celda,"si")

    sInput = ofv.GenerarInput("","Cancelar","submit","","10","bt","")
    Celda = ofv.GenCelda("","ttr","","","","","",sInput,"si")

    sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = sAccion & "&CantRegAMover=" & CantRegAMover
    sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

    'sAccion = "edProceddet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    'sAccion = sAccion & "&Opcion=no&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
    'sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2

    fila29= ofv.GenForm("Cancel","",sAccion,"","post","",Celda,"si")'
    tabla29 = ofv.GenTabla("","","","100%","","0","","0","0","","",fila29,"si")'
    formulario= Formulario & ofv.GenForm("Cancel","",sAccion,"","post","",tabla29,"si")'
    'Formulario = Formulario & ofv.GenForm("Cancel","",sAccion,"","post","",Celda,"si")

    Fila = ofv.GenRow("","","","","","","",Formulario,"si")
   ' sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")
    sHTML = sHTML & formulario
  end if

    if OpcAgregar = "si" then

      sql = "SELECT max(OBJACCION.secuen) FROM Objaccion, Objproced where "
      sql = sql & "Objaccion.codproced = Objproced.codproced and Objproced.codproced='" & nID & "'"
      set rsSecuencia = ofv.crearconsultaEx(Sql,cn,1,volver)
      nSecuencia = 10
      if not rsSecuencia.eof then
         if not isnull(rsSecuencia(0)) then nSecuencia = cint(rsSecuencia(0)) + 10
      end if
      call ofv.cerrarconsulta(rsSecuencia)

      Celda   = ""

      sInput  = ofv.GenerarInput("secuen",nSecuencia,"agregar","","5","dtn","SI")
      Celda   = Celda & ofv.GenCelda("","vti","","1%","","","",sInput,"si")
      IF session("APLMDCLI") = "BNX" OR session("APLMDCLI") = "GP"  or true THEN
         Strsql  = "SELECT * From TIPOACCALT where estado=3  Order By Tipoaccalt.descri"
       ELSE
        Strsql  = "SELECT * From TIPOACCALT where estado=3  and modo = '" & Fltipo & "'  Order By Tipoaccalt.descri"
      END IF
       ' response.write strsql
      'response.write strsql


        Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
        scomK = ""'"null" & Chr(9)

        scomD = "" '"" & Chr(9)
        SCOMSEP =""
        do while not rsperfil.eof
         scomK=scomK & SCOMSEP & rsperfil(0)
         scomd=scomd & SCOMSEP & rsperfil(1)
         SCOMSEP = chr(9)
         rsPerfil.movenext
          loop

     ' sInput =  ofv.gencomboFX("Agregarb", "tipoaccion", "", scomK, scomD, "")
       sInput =  ofv.GenerarInput("tipoaccion","301","hidden","","10","","") & ofv.GenerarInput("descriacc","NUEVA ACCION","HIDDEN","","10","","")

    ''  sInput  = ofv.GenerarCombox2(Strsql,cn,0,1,"tipoaccion")'' & "cambiar esto"

      sInput1 = ofv.GenerarInput("rellink","0","hidden","","10","","")
      sInput1 = sInput1 & ofv.GenerarInput("Modi","1","hidden","","10","","")

      sInput2 = ofv.GenerarInput("TipModi","A","hidden","","10","","")

      sukey   = GTkey(nCodigo)
      sInput3 = ofv.GenerarInput("codaccion",sukey,"hidden","","10","","")

      sInput4 = ofv.GenerarInput("TIPPROCED",WKAPLIC,"hidden","","10","","")

      Celda   = Celda & ofv.GenCelda("","vti","","90%","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")


      Fila    = ofv.genrow("","","","","20","","",Celda,"si")


      sAccion = "asp/AgregarEN.asp?Tabla=Objaccion&volver=" & vueltaC & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Codproced=" & nID & "&Codeven="
      sAccion = sAccion & sEven

      tabla= ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;'","","90%","","0","","0","0","","",fila,"si")
      Formulario = ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,tabla,"no")

      'Formulario = ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

      sHTML = sHTML & "<center><div  >"
      'sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;'","","90%","","0","","0","0","","",Formulario,"si") & "<br><br><br><br><br><br>"
      sHTML = sHTML & formulario

     'Aceptar y Cancelar
      sAccion = "edProcedDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Opcion=no&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2
      sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"") & "</div></center>"


      Agregar = ""
      Buscar  = ""
      Celda   = ""

    ''sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda)
      sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Indice & Test,Agregar,Buscar,celda)


    else
      'Botonera
      Agregar = ""
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    if Modificar then
      sAccion = "edProcedDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Opcion=si&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)
    end if

      if bModificar and bagregarbtn then

     '' sInput  = ofv.GenerarInput("","Accion","submit","","20","bt","")
     '' Celda   = "" 'ofv.GenCelda("","vt title='Agregar nueva accion a la lista.' ","","1%","","","",sInput,"si")

        sAccion = "AgregarAcciones.asp?tabla=Tipoaccalt&Opcion=si&volver=" & vueltaC
        sAccion = sAccion & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
      ''Agregar = BtnAgAccion(21,"Accion","xform",sAccion)
        Agregar = ofv.BotonAgregar(sAccion)
      end if

 ''end if

      if nID <> "no" then
         Buscar = ""
         if posicion > 0 then

            strsql = "select descri from OBJACCION LEFT JOIN TIPOACCALT ON OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
            strsql = strsql & " WHERE codproced = '" &  nID & "'"
            if bVersion = "NO" then
               strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
            end if
            strsql = strsql & " Order By secuen"
        ''''response.write strsql & " strsql " & chr(13)

            str1 = strsql   ' orden pagina
            str2 = strsql    ' orden alfabetico
  '         buscar = str1
            Buscar = ofv.Combobuscar(str1,str2,cn,CantRegAMostrar2,CantRegAMover2,replace(replace(vueltax,"_","&"),"?","?codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&"))

         end if
        'sHTML = sHTML & ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posi,"","",Indice & Test,Agregar,buscar,sRetornar)
    '''''xxxx=ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posi,"","","",Agregar,buscar,sRetornar)
         xxxx=ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posi,"","",Indice & Test,Agregar,buscar,sRetornar)
      end if
    end if

  if nID <> "no" then call ofv.cerrarconsulta(rs2)


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","") & xxxx

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)




  sHTML = sHTML & GTCOBS()

     sHTML = sHTML & "<script type='text/javascript'>"
     sHTML = sHTML & "alert(document.getElementById(disp_frm));"
      sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
      sHTML = sHTML & "  {"
      sHTML = sHTML & "  var r=confirm(disp_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
    '  sHTML = sHTML & "    alert(fr.name);"
      sHTML = sHTML & "    fr.submit();"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion!');"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"
      sHTML = sHTML & "function send_confirm(send_msg,send_href,send_aid)"
      sHTML = sHTML & "  {"
      sHTML = sHTML & "  var r=confirm(send_msg);"
      sHTML = sHTML & "  if (r==true)"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('ok ');"
      sHTML = sHTML & "    var fr= document.getElementById(send_aid);"
      sHTML = sHTML & "    fr.href = send_href;"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  else"
      sHTML = sHTML & "    {"
      sHTML = sHTML & "    alert('Se cancelo la accion! ');"
    '  sHTML = sHTML & "    return void(0);"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"
      sHTML = sHTML & "</script>"



  sHTML = sHTML & "<script>document.getElementById('Agregarb').submit()</script>" & "</body></html>"


  FUNCTION ZXinput2()
bc = Request.servervariables("HTTP_USER_AGENT")

bcn = instr(bc,"; MSIE")
if bcn > 0 then

  xexplor = mid(bc,bcn + 7,1)

else

  xexplor = "4"

end if

    if xexplor > "5" then
       xalign="center"
    else
       xalign="left"
    end if

''color='#aaCCEE' color2='#225588'

 xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:10;position:RELATIVE;width:45px;height:20px;' >"
'' xfsextr = "  <v:extrusion on='True' metal='True' type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
'' xfsshdw = "  <v:shadow on='False' />"
 xfsshdw = "  <v:shadow on='True' color='#eeeeee' weight='1px' /><v:stroke on='True' color='#eeeeee' weight='1px' />"
 xfsfill = "  <v:fill on='true' color='XXXCOL1XXX' color2='XXXCOL2XXX' type='gradient' angle='30' opacity='100%'  />"
'' xfsfill = " <v:fill  type='gradientradial' method='sigma' color='XXXCOL2XXX' color2='XXXCOL1XXX' colors='95% XXXCOL2XXX, 0% XXXCOL1XXX' opacity='100%' angle='45'  focusposition=""95%, 15%""  />"

        XSCOLOR2 = Session("CMFPNBOT")
        XSCOLOR1 = Session("CMFMENU")

        xfsfillD = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
        xfsfillD = replace(xfsfillD,"XXXCOL2XXX",XSCOLOR2)

        Xtypm  = "D"
        Xweig  = "600"

          Xinput = ""
          Xinput2 = "4"
          Xinput3 = "corflaticon"
          Xinput4 = "bte"
          Xinput5 = "Eliminar"
          Xinput6 = "_self"

          if ofv.modbot = "2" then
             Xinput7 = "4"
         else
            Xinput7 = "Eliminar"
          end if

         ''' Xinput = Xinput & xfstilo
          Xinput = Xinput & "  <v:roundrect arcsize='65%' id='AP'  style='z-index:4;position:RELATIVE;width:12px;height:12px;' >"
          Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

          XSCOLOR1 = "#ff0000"
          XSCOLOR2 = "#770000"

          IF  Xtyp   = "C" THEN
              xfsfillC = replace(xfsfill,"XXXCOL1XXX",XSCOLOR1)
              xfsfillC = replace(xfsfillC,"XXXCOL2XXX",XSCOLOR2)
              Xinput = Xinput & xfsfillC
          ELSE
              Xinput = Xinput & xfsfillD
          END IF

          Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
          Xinput = Xinput & "	<v:textbox inset='0px,-4px,0px,0px' style=""text-align:" & xalign & ";font-size:7pt;font-family:'corflaticon';font-weight:400;color:#FFFFFF;"" >4"
          Xinput = Xinput & "     </v:textbox>"
          Xinput = Xinput & xfsshdw
         ' Xinput = Xinput & xfsextr
          Xinput = Xinput & "  </v:roundrect>"




          ZXinput2= Xinput 'OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

   END FUNCTION

end if

response.write sHTML
response.flush

%>