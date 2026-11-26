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

  if ofv.explorador = "MSIE" then
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
     wkcodasoc   = rs(2)
  end if

  Strsql = "SELECT * From eco where eco = 0" & wkcodasoc
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     wkasocdesc = rs(1)
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
    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","15%","","","","Asociacion: ","si")
    Celda = Celda & ofv.GenCelda("","uti style='background-color:#f9f9f9;border-style:solid;border-width:1;border-color:#dddddd;' ","","85%","","","",wkasocdesc,"si")
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

     nProcAlt = GTkey(nCodigo) & "CORR"

     strsql = "Insert Into PLCIRCORR (codobjeto,tipobjeto,codcircorr,codnivel,plazomax,"
     strsql = strsql & "codcontin,codproced,codaccion,modi,RELLINK,REFLINK) "

     strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','"
     strsql = strsql & nProcAlt & "','0','0','"
     strsql = strsql & nConti & "','" & nProced & "','" & nAccion & "',"
     strsql = strsql & "0,'0','')"

 ''    response.write strsql & "<br>"
     Call OFV.crearconsultaEx(strsql, CN, 1, volver)



 end if


''''''''''''''''''''''''''''''''''''
''' encabezado circuito CORRECTIVO
''''''''''''''''''''''''''''''''''''''
  Formulario1 = ""
  formulario = ""


  Strsql = "SELECT * from PLCIRCORR "
  Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codaccion = '" & nAccion & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    nID = rs(3)
    nID2 = rs(0)





  else
    nID = "no"
  end if



  response.write sHTML
  sHTML = ""

' Definicion inicial

 if nID = "no" then

    Celda = ofv.GenCelda("","vtig","","90%","","","","Niveles alcanzados en el Circuito Correctivo","si")
    Fila = ofv.GenRow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;background-color:#F9F9F9;' ","","100%","","0","","0","0","","",Fila,"si")
    response.write sHTML
    response.flush
    sHTML = ""

     sPath  = Request.ServerVariables("SCRIPT_NAME")
     spathi = len(sPath) - InStrRev(sPath, "/", -1)
     sPath2 = right(sPath, spathi)
     sPath2 = sPath2 & "?" & "OpcionD=SI" & "&" & request.querystring


    SINPUTT1 = "<FONT COLOR='#993333' ><B><U>CIRCUITO CORRECTIVO:</U></B></FONT><BR><BR>"
    SINPUTT1 = SINPUTT1 & "No esta definido ningun circuito correctivo para la actividad:&nbsp;<B><U>" & wkactsec & "." & wkactdesc & "</U></B>, "
    SINPUTT1 = SINPUTT1 & " perteneciente al proceso:&nbsp;<B><U>" & wkprocdesc & "</U></B>, para el tipo de contingencia:&nbsp;<B><U>" & wkcontidesc & "</U></B>.<br>"
    SINPUTT1 = SINPUTT1 & "<br>Desea definir un circuito correctivo?."

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

    Celda = ofv.GenCelda("","vtig","","90%","","","","Niveles alcanzados en el Circuito Correctivo","si")
    Fila = ofv.GenRow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#990000;background-color:#F9F9F9;' ","","100%","","0","","0","0","","",Fila,"si")
    response.write sHTML
    response.flush
    sHTML = ""


 end if

  sHTML = ""
  response.flush

 end if



 if OpcAgregar = "si" then

 else



  bColor = "False"
  if nID <> "no" then


     Strsql = "SELECT * from PLCIRCORR "
     Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codaccion = '" & nAccion & "' "
'    response.write strsql
     set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs2.eof then rs2.move(CantRegAMover2)
        posicion = 0


       Filam = ""
       xTrColor = "#F9F9F9"

       WKSTYLE = "style='border-RIGHT-style:solid;border-RIGHT-width:1;border-RIGHT-color:#888888;' "
'---
        Celda  = ""
        Celda  = Celda & ofv.GenCelda("","vt colspan=2 " & WKSTYLE,"","","","","","NIVEL","si")
        Celda  = Celda & ofv.GenCelda("","vt colspan=2 " & WKSTYLE,"","","","","","PLAZO","si")
        Celda  = Celda & ofv.GenCelda("","vt colspan=5 " & WKSTYLE,"","","","","","ENTIDADES","si")
        Celda  = Celda & ofv.GenCelda("","vt " & WKSTYLE,"","","","","","CIRCUITO","si")
        Celda  = Celda & ofv.GenCelda("","vt colspan=2 " & WKSTYLE,"","","","","","&nbsp;","si")

    FilaM  = FilaM & ofv.GenRow("","aaa bgcolor='#dddddd' ","","","10","","",Celda,"si")


    do while not rs2.eof and posicion < CantRegAMostrar2
      nAccioncorr   = rs2(0)
      nAccionID = rs2(3)

      sForm     = "edProcedimientos" & rs2(0)

      'Eliminar
      sHref = ""
      if rs2(16) And rs2(17) = "E"   then
        sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        Celda = ofv.GenCelda("","vt","","2","","","",sImage,"si")
      else
        if bModificar   then
          if rs2(16) And rs2(17) = "A" then
            sHref = "asp/EliminarEN.asp?ID=" & nAccioncorr & "&Tabla=PLCIRCORR&volver=" & vueltaC
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
          else
            sHref = "asp/ActualizarEN.asp?ID=" & nAccioncorr & "&Tabla=PLCIRCORR&volver=" & vueltaC
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          end if
        end if

        if sHref <> ""   then
          Celda = ofv.BotonEliminar("Accion",sHref)
        else
            Celda = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")
        end if
      end if



      'Nivel de Contingencia

''''rs2(4)

        if bModificar then
          Strsql = "Select * From PLNiveles Where codasoc =" &  wkcodasoc & " Order By secuen"
          sID = sForm
          sInput = ofv.GenerarCombo(sForm,rs2(4).name,rs2(4),Strsql,cn,0,1,"")
        else
          sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"readonly",sForm,"80","dt","")
        end if

        Celda  = Celda & ofv.GenCelda("","ut","","80%","","","",sInput,"si")


      'Plazo maximo de resolucion

''''rs2(5)

        if bModificar then
          sInput = ofv.GenerarInput(rs2(5).name,rs2(5),"text",sForm,"3","dtn","")
        else
          sInput = ofv.GenerarInput(rs2(5).name,rs2(5),"readonly",sForm,"80","dt","")
        end if

        Celda  = Celda & ofv.GenCelda("","ut","","5%","","","",sInput,"si")

        Celda  = Celda & ofv.GenCelda("","ut","","2%","","","","&nbsp;","si")

  'Objetos Relacionados

    'Proveedor Primario

      sID = "edProcedCCorrRel.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&Codrel=PP"

      if rs2(6) <> ""  then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar Proveedor principal."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")



    'Proveedor alternativo

      sID = "edProcedCCorrRel.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&Codrel=PA"

      if rs2(7) <> ""  then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar Proveedor alternativo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


    'Equipo primario

      sID = "edProcedCCorrRel.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&Codrel=EP"

      if rs2(8) <> ""  then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar Equipo Primario."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


    'equipo alternativo

      sID = "edProcedCCorrRel.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&Codrel=EA"

      if rs2(9) <> ""  then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar Equipo alternativo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")


    'Centro procesamiento alternativo

      sID = "edProcedCCorrRel.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&Codrel=CA"

      if rs2(10) <> ""  then
         xclas  = "btbs"
         sInput = "<font class='btic' >&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >&#0157;</font>"
      end if

      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar Centro de procesamiento alternativo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")



    'Detalle del circuito correctivo

      nNivel = rs2(4)
      Nccorr = rs2(3)

      sID = "edProcedCCorrDet.asp?"
      sID = sID & "Accion=" & nAccion & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
      sID = sID & "&Tipobjeto=" & sTipo & "&Nivel=" & nNivel & "&Nccorr=" & Nccorr  & "&Conti=" & nConti


      xclas  = "btbs"
      sInput = "<font class='btic' >4</font>"


      sHref = sID

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      stitle = "Actualizar Procedimiento correctivo."
      sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")






  'Modi (actualizaciones por versionado)

       'Version
      sSelect = ""
      if bVersion = "SI" then
         if rs2(16) then 

              if rs2(17) = "A" then
                 stitle = "Paso agregado."
                 xclas  = "btba"
                 sInput = "<font class='btic'  >6</font>"
              elseif rs2(17) = "E" then
                 stitle = "Paso eliminado."
                 xclas  = "btbn"
                 sInput = "<font class='btic'  >5</font>"
              elseif rs2(17) = "M" then
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


      Tabla = ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","10","","",Celda,"si")

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & rs2(0) & "&Tabla=PLCIRCORR&volver=" & vuelta2
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Celda = ofv.GenForm(sForm,"",sAccion,"","post","",Tabla,"si")
      else
        Celda = ofv.GenForm("","","","","post","",Tabla,"si")
      end if


      FilaM  = FilaM & Celda


   '   FilaM  = FilaM & ofv.GenRow("","aaa bgcolor=" & xTrColor & " ","","","10","","",Celda,"si")







      if xTrColor = "#FFFFFF" then
         xTrColor = "#F9F9F9" 
      ELSE
         xTrColor = "#FFFFFF" 
      END IF


      Celda = ""



      rs2.movenext
      posicion = posicion + 1
    loop

      Tabla = ofv.GenTabla("","aaa style={border-bottom-style:solid;border-width:1;border-color:#cccccc;} ","","95%","","0","","-1","3","","",Filam,"si")
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
      sAccion = "edProcedContiCorr.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
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
        sAccion = "edProcedContiCorr.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sAccion = sAccion & "&OpcionD=SI&opcion2=no&opcion3=no&CantRegAMover=" & CantRegAMover
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
        '    Buscar = ofv.Combobuscar(str1,str2,cn,CantRegAMostrar2,CantRegAMover2,replace(replace(vueltax,"_","&"),"?","?codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&"))

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