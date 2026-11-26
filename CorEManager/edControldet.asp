<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  sTest = request.querystring("Test")

  if request.querystring("CantRegAMover") = "x" then
    CantRegAMover = request.querystring("CantRegAMover")
  else
    CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  end if

  if request.querystring("XEVEN") <> "" then
     Session("XEVEN") = request.querystring("XEVEN")
     XEVEN = request.querystring("XEVEN")
  else
     XEVEN = Session("XEVEN")
  end if

  CantRegAMover2 = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  If BrowserName() = "IE" OR BrowserName() = "CHROME" then
    CantRegAMostrar2 = 7
  else
    CantRegAMostrar2 = 5
  end if

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)
 'response.write vuelta & "<br>" & vuelta2 & "<br>" & vueltac


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

   shtml = shtml & "<script language=javascript>" & chr(13)

   shtml = shtml & "function AbrirVentana(datos) { " & Chr(13)
   shtml = shtml & " window.open(datos"
   shtml = shtml & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
   shtml = shtml & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
   shtml = shtml & " ); " & Chr(13)
   shtml = shtml & " } " & Chr(13) & Chr(13)
   shtml=shtml & "</script>" & chr(13)


  'Encabezados
  Formulario = ofv.Encabezados("Frecuencia","Sector","Perfil","referencia")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  Strsql = "SELECT Objcontrol.objid, Objcontrol.codsector, Objcontrol.codsector, "
  Strsql = Strsql & "OBJCONTROL.frecuencia, OBJCONTROL.codperfil, OBJCONTROL.referencia FROM OBJCONTROL "
  Strsql = Strsql & "Where Objcontrol.codobjeto = " & nCodigo & " and  Objcontrol.objid = " & XEVEN
  Strsql = Strsql & " order by Objcontrol.objid"
 'response.write Strsql & " Strsql -85- " & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 'muestro = rs(4)
 'response.write muestro & " muestro " & "<br>"

  if not rs.eof then
    nID = rs(0)

    'Eliminar

      Celda = ofv.GenCelda("","UT","","1%","","","","-","si")

    'Frecuencia
      sFrecuencia = "&nbsp;"
      if rs(3) <> "" then
        Strsql = "Select descri From Frecuencia  Where codfrec = '" & rs(3) & "'"
        set rsFrecuencia = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsFrecuencia.eof then sFrecuencia = rsFrecuencia(0)
        call ofv.cerrarconsulta(rsFrecuencia)
      end if
      sInput = "&nbsp;&nbsp;" & sFrecuencia

    Celda = Celda & ofv.GenCelda("","vt","","30%","","","",sInput,"si")


    'Sector
    if bModificar then

      Strsql = "Select Eco.eco,Eco.descri,a.cliente From Objsector, Eco "
      Strsql = Strsql & " inner join clienteplan a on (eco.ecoUN=a.id) "
      Strsql = Strsql & " where eco.estado=3 and "
      Strsql = Strsql & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo
     'response.write Strsql & " Strsql -113- " & "<br>"
     '' sInput = ofv.GenerarCombo("edControles1",rs(1).name,rs(1).value,Strsql,cn,0,1,"Sector")

       set rsSectorx = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      scomK = ""'"null" & Chr(9)
      scomD = "" '"" & Chr(9)
      SCOMSEP =""

      do while not rsSectorx.eof
       scomK=scomK & SCOMSEP & rsSectorx(0)
       if Session("EMPMUN")  = "S" then
        scomd=scomd & SCOMSEP & rsSectorx(2) & "&nbsp;-&nbsp;" & rsSectorx(1)
       else
        scomd=scomd & SCOMSEP &  rsSectorx(1)
       end if

       SCOMSEP = chr(9)
       rsSectorx.movenext
    loop

      sInput =  ofv.gencomboFX("edControles1",rs(1).name,rs(1).value, scomK, scomD, "")
    else
      if rs(1) <> "" then
        Strsql = "Select descri From Eco Where eco = " & rs(1)
       'response.write Strsql & " Strsql -118- " & "<br>"
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then sSector = rsSector(0)
        call ofv.cerrarconsulta(rsSector)
      end if
      sInput = ofv.GenerarInput("",sSector,"readonly","","80","dt","")
    end if


   if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID & "&Tabla=Objcontrol&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sInput = ofv.GenForm("edControles1","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenForm("","","","","post","",sInput,"si")
    end if


    Celda = Celda & ofv.GenCelda("","UT","","30%","","","",sInput,"si")




    'Perfil
    if bModificar then

      Strsql = "SELECT TIPOPERFIL.tipoperfil,TIPOPERFIL.descri "
      Strsql = Strsql & "FROM Objsector, ECO, ECOPERFIL, TIPOPERFIL "
      Strsql = Strsql & " where Eco.eco = Objsector.codsector and "
      Strsql = Strsql & " ECO.eco = ECOPERFIL.codsector and "
      Strsql = Strsql & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
      Strsql = Strsql & " and objsector.codsector = " & rs(1)
      Strsql = Strsql & " and codobjeto = " & nCodigo
     'response.write Strsql & " Strsql -147- " & "<br>"
      'sInput = ofv.GenerarCombo("edControles2",rs(4).name,rs(4).value,Strsql,cn,0,1,"sb")

       set rsPerfilx = ofv.crearconsultaEx(Strsql,cn,1,parametros)

    scomK = ""'"null" & Chr(9)
    scomD = "" '"" & Chr(9)
    SCOMSEP =""

    do while not rsPerfilx.eof
       scomK=scomK & SCOMSEP & rsPerfilx(0)
       scomd=scomd & SCOMSEP & rsPerfilx(1)
       SCOMSEP = chr(9)
       rsPerfilx.movenext
    loop

      sInput =  ofv.gencomboFX("edControles2",rs(4).name,rs(4).value,scomK,scomD,"sb")
    else
      sPerfil = ""
      if rs(4) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(4) & "'"
      ''response.write Strsql & " Strsql -153- " & "<br>"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then sPerfil = rsPerfil(0)
        call ofv.cerrarconsulta(rsPerfil)
      end if
      sInput = ofv.GenerarInput("",sPerfil,"readonly","","40","","")
    end if



   if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID & "&Tabla=Objcontrol&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sInput = ofv.GenForm("edControles2","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenForm("","","","","post","",sInput,"si")
    end if

    Celda = Celda & ofv.GenCelda("","UT","","30%","","","",sInput,"si")


      'Texto de Referencia
      sID = "edControlRef.asp?"
      sID = sID & "ID=" & rs(0) & "&Tabla=Objcontrol&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sImage = ofv.GenImage("","","","","Iconos/ICorSAccAlt.jpg","0","")
      sHref = sID

      if rs(5) <> "" then
         xclas  = "btbs"
         sInput = "<font class='btic' >4&#0157;</font>"
      else
         xclas  = "btbn"
         sInput = "<font class='btic' >4&#0157;</font>"
      end if

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "

      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      stitle = "Actualizar referencia del Control."
      'sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","center","5%","","","",sInput,"si")
      'Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      XXX= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
      Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","center","5%","","","",XXX,"si")


    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Formulario = Formulario & Fila



    rs.movenext
  else
    nID = "no"
  end if
  sHTML = sHTML & ofv.GenTabla("","","","92%","","0","","0","0","","",Formulario,"si")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Controles

  'Boton test
   sTest = "T"
 ' sAccion = "CoControles.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
 ' sAccion = sAccion & "&Menu=PORESTADO&Sector=3&origen=ED&CantRegAMover=A&control=" & nID
   sAccion= " Consultar.asp?Codigo=" & nCodigo & "&Tipo=" & STIPO &"&niv1=7&niv2=" & nID & "&test=" & sTest
  TEST = ofv.BotonTest(sAccion)

  'Boton Volver
  sAccion = "edcontroles.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Indice = ofv.BotonVolver(sAccion)



  'Encabezados
  Celda = ofv.GenCelda("","ttT","","100%","","","","Controles","si")
'  Celda = celda & indice
'  Celda = celda & test
  Fila = ofv.GenRow("","","","","30","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",Fila,"si")




  if nID <> "no" then
    'Encabezados

     Celda      = ofv.GenCelda("","ttR","","","","","",".","si")
     Celda      = celda & ofv.GenCelda("","ttR","","","","","","OBS","si")
     Celda      = celda & ofv.GenCelda("","ttR","","","","","","Sec","si")
     Celda      = celda & ofv.GenCelda("","ttR COLSPAN=2","","","","","","Descripcion","si")
     Formulario = ofv.GenRow("","","","","","","",Celda,"si")

  '  Formulario = ofv.Encabezados("Sec","Descripcion")

  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:8pt;background-color:#f7eedd;} "


    strsql = "SELECT Objcontrol2.*,b.codinx2  "

    strsql = strsql & " FROM Objcontrol2 "
    strsql = strsql & " inner join Objcontrol on (Objcontrol2.control = Objcontrol.objid) "

    'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (Objcontrol2.objid = b.codinx2) "
     strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast( Objcontrol2.objid as nvarchar)= b.codinx2) "


    strsql = strsql & " where "

    strsql = strsql & "  Objcontrol.objid = " & nID

    if bVersion = "NO" then
       strsql = strsql & " and (isnull(Objcontrol.modi,0) = 0 or (isnull(Objcontrol.modi,0) <> 0 and Objcontrol.tipmodi <> 'E')) "
       strsql = strsql & " and (isnull(Objcontrol2.modi,0) = 0 or (isnull(Objcontrol2.modi,0) <> 0 and Objcontrol2.tipmodi <> 'E')) "
    end if
    strsql = strsql & " ORDER BY Objcontrol2.secuencia, Objcontrol2.descri, Objcontrol2.OBJID "

   ''response.write strsql & "<br>" & chr(13)
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then rs2.move(CantRegAMover2)
    posicion = 1
    do until rs2.eof or posicion > CantRegAMostrar2
      nAccion = rs2(0)
      sForm = "edControles" & rs2(0)

      'Eliminar
      sHref = ""
      if rs2(6) And rs2(7) = "E" then
''      sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        sImage = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
        Celda = ofv.GenCelda("","","","","","","",sImage,"si")
      else
        if bModificar then
          if rs2(6) And rs2(7) = "A" then
             sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objcontrol2&volver=" & vueltaC
             sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
          else
             sHref = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objcontrol2&volver=" & vueltaC
             sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          end if
        end if
        if sHref <> "" then
          Celda = ofv.BotonEliminar("Control",sHref)
        else
          Celda = ofv.GenCelda("","","","","","","","","si")
        end if
      end if


xclas = "ut"


      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")

     'response.write nCodigo & " nCodigo " & "<br>"
	 'response.write sTipo   & " sTipo " & "<br>"

 ''''' agregado para color del cuadrito de Observaciones version 4.6
       wComentario = ""
       strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and "
       strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"      ' and "
       strsql = strsql & " order by codcircuito, codusr, objid"
     ''response.write strsql
       set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rss.eof then
          wComentario = rss(13)
       end if
       call ofv.cerrarconsulta(rss)

      if (Nestado=5  or Nestado=6) and (rs2(14) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs2(14)  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion


      ''''''agregado para color del cuadrito de Observaciones version 4.6

	           if wComentario = "" then
	              sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
	       '''''''sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

	           else
	 	          sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
	           end if
      ''''''

              '' Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",sInput,"si")
              '' Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",Celdaobs,"si")
               Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")
               Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")
      end if


      Celda = Celda & Celdaobs



      'Secuencia
      if bModificar then
        sInput = ofv.GenerarInput(rs2(5).name,rs2(5),"text",sForm & "X1","2","dtn","SI")
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objcontrol2&volver=" & vueltaC
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")

      else
        sInput = ofv.GenerarInput("",rs2(5),"readonly","","2","","SI")
      end if


      Celda = Celda & ofv.GenCelda("","aaa valign=top","","","","","",sInput,"si")

      wasotxt = "N"
      stylob  = ""


         if not isnull(rs2(8))  then
            if  clng(rs2(8)) > 0 then

                Strsql = "SELECT * From OBJLNK where objid = 0" & rs2(8)
                set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                if not rsx.eof then
                   wasotxt  = rsx(14) & ""
                   IF wasotxt = "" THEN wasotxt = "N"
                end if
                call ofv.cerrarconsulta(rsx)

            end if
          end if

       IF wasotxt = "S" THEN

          stylob = "aaa style='border-style:solid;border-width:1;border-color:#990000;' "

       end if



      'Texto Descripcion
      If BrowserName() = "IE" OR BrowserName() = "CHROME" then
        nCol = 90
      else
        nCol = 90
      end if
      sInput = "<font class=at><textarea class=dt name=descri cols=" & nCol & " rows=5 "
      if bModificar and wasotxt = "N" then
        sInput = sInput & " onblur=" & sForm & "X2" & ".submit() "
        sInput = sInput & " onkeypress=" & chr(34)
        sInput = sInput & " if (this.value.length > 2450) {alert('El texto es demasiado"
        sInput = sInput & " extenso'); return false;}" & chr(34) & " >"
      else
        sInput = sInput & " readonly>"
      end if
      sInput = sInput & rs2(4) & "</textarea></font>"

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objcontrol2&volver=" & vueltaC
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sInput = ofv.GenForm(sForm & "X2","",sAccion,"","post","",sInput,"si")


      end if



      Celda = Celda & ofv.GenCelda("",stylob,"","","","","",sInput,"si")


  'Link a paginas internas/externas

         sID = "EdGENLnk.asp?"
         sID = sID & "ID=" & rs2(0) & "&TABLA=OBJCONTROL2&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo

         if not isnull(rs2(8))  then
            if  clng(rs2(8)) > 0 then
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


      sInput = "<div class=" & xclas & " >" & sinput & "</div>"

      xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      stitle = "Actualizar link a otro documento."
      'sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      'Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      xxx= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

     Celda = Celda &  ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")



      'Version
      sSelect = ""
      if bVersion = "SI" then

      if rs2(6) then sSelect = "selected"
      if sSelect <> "" then
        if bModificar then
          if rs2(7) = "A" then
             sInput1 = "Agregada"
          elseif rs2(7) = "E" then
             sInput1 = "Eliminada"
          elseif rs2(7) = "M" then
             sInput1 = "Modificada"
          end if
          sMVal = "TRUE"
          if not rs2(6) then  sMVal = "FALSE"

          sInput = "<font class=at><select name=" & rs2(6).name & " value=" & sMVal
          sInput = sInput & " onchange=" & sForm & ".submit(); >"
          sInput = sInput & "<option value=FALSE>No</option>"
          sInput = sInput & "<option value=TRUE selected>" & sInput1
          sInput = sInput & "</option></select></font>"
        else
          if rs2(7) = "A" then
             sInput1 = "Agregada"
          elseif rs2(7) = "E" then
             sInput1 = "Eliminada"
          elseif rs2(7) = "M" then
             sInput1 = "Modificada"
          end if
          sInput = ofv.GenerarInput(rs2(6).name,sInput1,"readonly","","10","dt","")
        end if
        Celda = Celda & ofv.GenCelda("","uts","","2%","","","",sInput,"si")
      else
        Celda = Celda & ofv.GenCelda("","uts","","2%","","","","&nbsp;","si")

      end if
      else
        Celda = Celda & ofv.GenCelda("","uts","","2%","","","","&nbsp;","si")
      end if


      Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")



      rs2.movenext
      posicion = posicion + 1
    loop
    sHTML = sHTML & ofv.GenTabla("","","","95%","","0","","0","0","","",Formulario,"si")

      Agregar = ""

    if bModificar then
      'Boton Agregar
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = "AgregarControl2.asp?Tabla=Objcontrol2&volver=" & vueltaC & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & nID & "&CantRegAMover="
      sAccion = sAccion & CantRegAMover & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)



    end if

  ''wwwbarra =  ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","",INDICE,Agregar,sRetornar)
    wwwbarra =  ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","",INDICE,TEST,Agregar,sRetornar)

    call ofv.cerrarconsulta(rs2)
  end if


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------


    sHTML = sHTML & wwwbarra

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>