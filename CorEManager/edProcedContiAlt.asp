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

  dim wstipo(6)
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  if request.querystring("Conti") <> "" then
     Session("Conti") = request.querystring("Conti")
     nConti = request.querystring("Conti")
  else
     nConti = Session("Conti")
  end if

  if request.querystring("Accion") <> "" then
     Session("Accion") = request.querystring("Accion")
     nAccion = request.querystring("Accion")
  else
     nAccion = Session("Accion")
  end if

  if request.querystring("Opcion") <> "" then
     OpcAgregar = request.querystring("Opcion")
  else
     OpcAgregar = "no"
  end if

  if request.querystring("OpcionD") <> "" then
     OpcDefinir = request.querystring("OpcionD")
  else
     OpcDefinir = "no"
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


  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

'''''''''''''''''''''''''''''''''''
'' encabezado procesos ''''''''''''
'''''''''''''''''''''''''''''''''''


  Strsql = "SELECT * From PLTIPOCONT where tipocontin = " & nConti
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     wkcontidesc = rs(1)
  end if

  Strsql = "SELECT * From objaccion where CODACCION = '" & nAccion & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     nProced = rs(3)
     nAccCod = rs(18)
     wkactdesc = rs(4)
     wkactsec = rs(10)

        if (rs(4) & "") <> "" then
          Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs(4)
          set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsAcciones.eof then
             wkactdesc = rsAcciones(0)
          END IF
          call ofv.cerrarconsulta(rsAcciones)

        end if

  end if



  Strsql = "SELECT * from csprocmaster "
  Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codproced = '" & nProced & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then

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

      Nsector = 0
      sSector = ""

      if rs(5) <> "" then

        Strsql = "Select descri,Eco.eco From Eco, Objsector where "
        Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.objid = " & rs(5)
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then
           sSector = rsSector(0)
           Nsector = rsSector(1)
        end if
        call ofv.cerrarconsulta(rsSector)
      end if


    'Perfil

      NPerfil = ""
      sPerfil = ""

      if rs(6) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then
           sPerfil = rsPerfil(0)
           NPerfil = rs(6)
        end if
        call ofv.cerrarconsulta(rsPerfil)
      end if

    wkprocdesc = wksecuencia & "." & sEvento & "-"  & sSector & "-"  & sPerfil

    Formulario1 = ""
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Proceso: ","si")
    Celda = Celda & ofv.GenCelda("","uti style='background-color:#f9f9f9;border-style:solid;border-width:1;border-color:#dddddd;' ","","85%","","","",wkprocdesc,"si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Actividad: ","si")
    Celda = Celda & ofv.GenCelda("","uti style='background-color:#f9f9f9;border-style:solid;border-width:1;border-color:#dddddd;' ","","85%","","","",wkactsec & "." & wkactdesc,"si")
    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Contingencia: ","si")
    Celda = Celda & ofv.GenCelda("","uti style='background-color:#f9f9f9;border-style:solid;border-width:1;border-color:#dddddd;' ","","85%","","","",wkcontidesc,"si")
    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","-1","2","","",Formulario,"si")

  END IF


  response.write sHTML
  response.flush
  sHTML = ""



'''''''''''''''''''''''''''''''''''''''''''''''
'' opcion definir circuito alternativo ''''''''
'''''''''''''''''''''''''''''''''''''''''''''''

 if  OpcDefinir = "SI" then

     nProcAlt = GTkey(nCodigo) & "ALT"

     strsql = "Insert Into PLCIRCALT (codobjeto,tipobjeto,codcircalt,eco,codperfil,"
     strsql = strsql & "codcontin,codproced,codaccion,modi,RELLINK,REFLINK) "

     strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','"
     strsql = strsql & nProcAlt & "','" & Nsector & "','" & Nperfil & "','"
     strsql = strsql & nConti & "','" & nProced & "','" & nAccion & "',"
     strsql = strsql & "0,'0','')"

 ''    response.write strsql & "<br>"
     Call OFV.crearconsultaEx(strsql, CN, 1, volver)



 end if


''''''''''''''''''''''''''''''''''''
''' encabezado circuito alternativo
''''''''''''''''''''''''''''''''''''''
  Formulario1 = ""
  formulario = ""


  Strsql = "SELECT * from PLCIRCALT "
  Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codaccion = '" & nAccion & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    nID = rs(3)
    nID2 = rs(0)


   'Sector
    if bModificar then

      Strsql = "Select Eco.eco,Eco.descri From Objsector, Eco where "
      Strsql = Strsql & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo
      sInputST = ofv.GenerarCombo("edProcedimientos",rs(4).name,rs(4).value,Strsql,cn,0,1,"sb")
    else
      sSector = ""
      if rs(4) <> "" then

        Strsql = "Select descri From Eco, Objsector where "
        Strsql = Strsql & "Eco.eco = Objsector.codsector and Eco.eco = " & rs(4)
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then sSector = rsSector(0)
        call ofv.cerrarconsulta(rsSector)
      end if

      sInputST = sSector
    end if


    'Perfil
    if bModificar then

      Strsql = "SELECT TIPOPERFIL.tipoperfil,TIPOPERFIL.descri "
      Strsql = Strsql & "FROM Objsector, ECO, ECOPERFIL, TIPOPERFIL "
      Strsql = Strsql & " where Eco.eco = Objsector.codsector and "
      Strsql = Strsql & " ECO.eco = ECOPERFIL.codsector and "
      Strsql = Strsql & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
      Strsql = Strsql & " and ECO.eco = " & rs(4).value

      sInputPF = ofv.GenerarCombo("edProcedimientos",rs(16).name,rs(16).value,Strsql,cn,0,1,"sb")
    else
      sPerfil = ""
      if rs(16) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(16) & "'"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then sPerfil = rsPerfil(0)
        call ofv.cerrarconsulta(rsPerfil)
      end if

      sInputPF = sPerfil
    end if

    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Sector: ","si")
    Celda = Celda & ofv.GenCelda("","uti style='background-color:#f9f9f9;border-style:solid;border-width:1;border-color:#dddddd;' ","","85%","","","",sInputST,"si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","perfil: ","si")
    Celda = Celda & ofv.GenCelda("","uti style='background-color:#f9f9f9;border-style:solid;border-width:1;border-color:#dddddd;' ","","85%","","","",sInputPF,"si")
    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

    Formulario =  ofv.GenTabla("","","","100%","","0","","-1","2","","",Formulario,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID2 & "&Tabla=PLCIRCALT&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario2 = ofv.GenForm("edProcedimientos","",sAccion,"","post","",Formulario,"si")
    else
      Formulario2 = ofv.GenForm("edProcedimientos","","","","post","",Formulario,"si")
    end if

    sHTML = sHTML & Formulario2


  else
    nID = "no"
  end if



  response.write sHTML
  sHTML = ""

' Acciones del procedimiento

 if nID = "no" then

    Celda = ofv.GenCelda("","vtig","","90%","","","","Circuito Alternativo para el proceso","si")
    Fila = ofv.GenRow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;background-color:#F9F9F9;' ","","100%","","0","","0","0","","",Fila,"si")
    response.write sHTML
    response.flush
    sHTML = ""

     sPath  = Request.ServerVariables("SCRIPT_NAME")
     spathi = len(sPath) - InStrRev(sPath, "/", -1)
     sPath2 = right(sPath, spathi)
     sPath2 = sPath2 & "?" & "OpcionD=SI" & "&" & request.querystring


    SINPUTT1 = "<FONT COLOR='#993333' ><B><U>CIRCUITO ALTERNATIVO:</U></B></FONT><BR><BR>"
    SINPUTT1 = SINPUTT1 & "No esta definido ningun circuito alternativo para la actividad:&nbsp;<B><U>" & wkactsec & "." & wkactdesc & "</U></B>, "
    SINPUTT1 = SINPUTT1 & " perteneciente al proceso:&nbsp;<B><U>" & wkprocdesc & "</U></B>, para el tipo de contingencia:&nbsp;<B><U>" & wkcontidesc & "</U></B>.<br>"
    SINPUTT1 = SINPUTT1 & "<br>Desea definir un circuito alternativo?."

    BOTONT1  = "<INPUT TYPE=BUTTON CLASS=BT VALUE='GENERAR' onclick='document.FVINC1.submit();' >"

    Celda = ofv.GenCelda("","ut valign=top STYLE='text-transform:uppercase;text-align:justify;background-color:#ffffff;' ","","","150","","",SINPUTT1,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    celda = ""
    Celda = Celda & ofv.GenCelda("","vt","center","","","","",BOTONT1,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Formulario = ""

    sHTML = sHTML & "<CENTER><br><br><br>" & ofv.GenTabla("","AAA border=1 BGCOLOR='#DDDDDD' ","","60%","","0","","5","4","","",Fila,"si") & "</CENTER>"

    sAccion    = spath2

    sinputh    = ofv.GenerarInput("CODV","XC","hidden","","1","dt","")
    Formulario = Formulario & ofv.GenForm("FVINC1","aaa style='display:none;' ",sAccion,"","post","",sinputh,"si")

    sHTML = sHTML & Formulario



 else

 if OpcAgregar = "si" then

    response.write "<br>"
    Celda = ofv.GenCelda("","ttt","","","","","","Agregar Paso","si")
    Fila  = ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")
    response.write sHTML
    sHTML = ""
 else



  'Encabezados


 'Agregar Accion

  if bModificar then
    sInput  = ofv.GenerarInput("","Accion","submit","","20","bt","")
    Celda   = ofv.GenCelda("","vt title='Agregar nueva accion a la lista.' ","","1%","","","",sInput,"si")
    sAccion = "AgregarAcciones.asp?tabla=Tipoaccalt&Opcion=si&volver=" & vueltaC
    sAccion = sAccion & "_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
    test    = ofv.GenForm("agracc","",sAccion,"","post","",Celda,"si")
  end if

 'Cursograma
    sInput = ofv.GenerarInput("","Cursograma","submit","","20","bt","")
    Celda = ofv.GenCelda("","vt title='Abrir grafico del cursograma.' ","","1%","","","",sInput,"si")
    sAccion = "Cursograma.asp?volver=" & vuelta & "&Codproced=" & nID & "&Tipobjeto=" & sTipo
    Formularioc = ofv.GenForm("Cursograma","",sAccion,"Grama","post","",Celda,"si")
''''
 'Importar procedimiento
    Formularioi = "&nbsp;"
    if bmodificar then
      sInput = ofv.GenerarInput("","Importar","submit","","10","bt","")
      Celda = ofv.GenCelda("","vt title='Incorporar pasos desde otro Circuito.' ","","1%","","","",sInput,"si")
      sAccion = "edProcedContiAlt.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Opcion=no&Opcion2=si&Opcion3=no&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2
      Formularioi = ofv.GenForm("Importar","",sAccion,"","post","",Celda,"si")
    end if
''''



         Celda = ofv.GenCelda("","vtig","","90%","","","","Circuito Alternativo para el proceso ","si")
         Celda = celda & indice
       '  Celda = celda & test
       '  Celda = celda & Formularioc
      '   Celda = celda & Formularioi




  Fila = ofv.GenRow("","","","","25","","",Celda,"si")
  sHTML = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;background-color:#F9F9F9;' ","","100%","","0","","0","0","","",Fila,"si") & "<br>"
  response.write sHTML
  sHTML = ""
 end if

  sHTML = ""
  response.flush

 end if

 if OpcAgregar = "si" then

 else



  bColor = "False"
  if nID <> "no" then


    strsql = "select * from OBJACCION "
    strsql = strsql & "WHERE codproced = '" &  nID & "'"
    if bVersion = "NO" then
       strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    end if
    strsql = strsql & " Order By secuen"
'    response.write strsql
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then rs2.move(CantRegAMover2)
    posicion = 0


    Filam = ""
    xTrColor = "#F9F9F9"
'---

    do while not rs2.eof and posicion < CantRegAMostrar2
      nAccion   = rs2(0)
      nAccionID = rs2(18)
      nselec    = "Expr2"
      sForm     = "edProcedimientos" & rs2(0)

      'Eliminar
      sHref = ""
      if rs2(14) And rs2(15) = "E"  and (rs2(20) & "") <> "PL" then
        sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        Celda = ofv.GenCelda("","vt","","2","","","",sImage,"si")
      else
        if bModificar  and (rs2(20) & "") <> "PL" then
          if rs2(14) And rs2(15) = "A" then
            sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
          else
            sHref = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objaccion&volver=" & vueltaC
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          end if
        end if

        if sHref <> ""  and (rs2(20) & "") <> "PL" then
          Celda = ofv.BotonEliminar("Accion",sHref)
        else
            Celda = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")
        end if
      end if

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
          Strsql = "Select descri From PLTipoaccalt Where tipoaccionalt = " & rs2(4)
          set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsAcciones.eof then
             sAccionM = rsAcciones(0)
          END IF
          call ofv.cerrarconsulta(rsAcciones)
        end if


      westado = "S"

      for it=1 to 6
          wstipo(it) = "N"
      next






      sVerbo = sAccionM
      sAcciont = sAccionM
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccionM,i-1)
        sAcciont = mid(sAccionM,i+1)
      end if


      if wkaplic <> "PL" then

         sInput = "-" & sInputS & ".&nbsp;<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
         sInput = sInput & replace(sAcciont," ","&nbsp;") & ""

         sID = "EdProcedDetProp.asp?"
         sID = sID & "AccId=" & nAccion & "&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&CantRegMover=" & CantRegAMover2

         sHref = sID
         sInput = ofv.GenLink("","nt title='Propiedades del Paso.' ",sHref,"","","",sInput,"si")

      else

         sInput = "-" & sInputS & ".&nbsp;" & sVerbo & "" & "&nbsp;&nbsp;&nbsp;"
         sInput = sInput & replace(sAcciont," ","&nbsp;") & ""

      end if


      Celda  = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")






  'Texto de referencia

      sID = "TextoEsquema.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
      sID = sID & "_Tipobjeto=" & sTipo

      if rs2(11) <> "" or rs2(16) <> "" or rs2(17) <> "" then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar texto de referencia."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


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
            sInput = "<font class='btic'  >&#0126;</font>"
         end if

         sHref = sID

      else
         sHref  = "javascript:void(0);"
         xclas  = "ut"
         sInput = "<font face='corflaticon' size=2 color='#eeeeee' >&#032;</font>"
      end if

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Actualizar link a otro documento."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

   end if



  'Asociacion de Complementos

      sID = "EdProcedCMP.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
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
      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Asociar componente."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

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

      sInput = "<div class=" & xclas & " style='width:3;' ><center>&nbsp;" & sinput & "&nbsp;</center></div>" & sInput1

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
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

   end if


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

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Establecer Sector Continuador."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

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
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")










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
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


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
  sAccion = "CoProcedimientos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sAccion = sAccion & "&Menu=PORESTADO&Sector=3&origen=ED&CantRegAMover=A&proced=" & XEVEN
  TEST = ofv.BotonTest(sAccion)

  'Boton Volver
  sAccion = "edProcedContingENCIAS.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Accion=" & Session("Accion")
  Indice = ofv.BotonVolver(sAccion)





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

      Strsql  = "SELECT * From PLTIPOACCALT Order By PLTipoaccalt.descri"
      sInput  = ofv.GenerarCombox2(Strsql,cn,0,1,"tipoaccion")

      sInput1 = ofv.GenerarInput("rellink","0","hidden","","10","","")
      sInput1 = sInput1 & ofv.GenerarInput("Modi","1","hidden","","10","","")

      sInput2 = ofv.GenerarInput("TipModi","A","hidden","","10","","")

      sukey   = GTkey(nCodigo)
      sInput3 = ofv.GenerarInput("codaccion",sukey,"hidden","","10","","")
      sInput4 = ofv.GenerarInput("TIPPROCED","PLALT","hidden","","10","","")
      Celda   = Celda & ofv.GenCelda("","vti","","90%","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")

      Fila    = ofv.genrow("","","","","20","","",Celda,"si")


      sAccion = "asp/AgregarEN.asp?Tabla=Objaccion&volver=" & vueltaC & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Codproced=" & nID & "&Codeven="
      sAccion = sAccion & sEven

      Formulario = ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"no")

      sHTML = sHTML & "<center><div  >"
      sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;'","","90%","","0","","0","0","","",Formulario,"si") & "<br><br><br><br><br><br>"

     'Aceptar y Cancelar
      sAccion = "edProcedContiAlt.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Opcion=no&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
      sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2
      sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"") & "</div></center>"

      Agregar = ""
      Buscar  = ""
      Celda   = ""

      sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Indice & Test,Agregar,Buscar,celda)


    else

    if nID <> "no" then

      'Botonera
      Agregar = ""
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      if bModificar then
        sAccion = "edProcedContiAlt.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sAccion = sAccion & "&Opcion=si&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
        sAccion = sAccion & "&CantRegAMover2=" & CantRegAMover2
        Agregar = ofv.BotonAgregar(sAccion)
      end if



      if nID <> "no" then
         Buscar = ""
         if posicion > 0 then

            strsql = "select descri from OBJACCION LEFT JOIN TIPOACCALT ON OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
            strsql = strsql & " WHERE codproced = '" &  nID & "'"
            if bVersion = "NO" then
               strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
            end if
            strsql = strsql & " Order By secuen"

            str1 = strsql   ' orden pagina
            str2 = strsql    ' orden alfabetico
            Buscar = ofv.Combobuscar(str1,str2,cn,CantRegAMostrar2,CantRegAMover2,replace(replace(vueltax,"_","&"),"?","?codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&"))

         end if
        sHTML = sHTML & ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posi,"","",Indice & Test,Agregar,buscar,sRetornar)
      end if
     else

      Agregar = ""
      Buscar  = ""
      Celda   = Indice

      sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda)

     end if


    end if

  if nID <> "no" then call ofv.cerrarconsulta(rs2)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML
response.flush

%>