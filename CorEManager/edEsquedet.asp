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

  CantRegAMover=cdbl(request.querystring("CantRegAMover"))
  CantRegAMostrar= 1
  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

  if request.querystring("XEVEN") <> "" then
     Session("XEVEN") = request.querystring("XEVEN")
     XEVEN = request.querystring("XEVEN")
  else
     XEVEN = Session("XEVEN")
  end if

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------





  Fila = ofv.Encabezados("evento","sector&nbsp;origen","sector&nbsp;DESTINO","referencia")



  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = 0" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    nestado = rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  Strsql = "SELECT objesquema.*, OBJEVEN.descri, OBJEVEN.secuen, OBJEVEN.descri "
  Strsql = Strsql & "FROM OBJEVEN, OBJESQUEMA where "
  Strsql = Strsql & "OBJEVEN.objid = OBJESQUEMA.codevento and objesquema.objid = " & XEVEN
  Strsql = Strsql & " and objesquema.codobjeto=" & nCodigo & " order by objesquema.objid"
 ' response.write "<br>" & strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    nID = rs(0)

    'Eliminar
      sInput = rs(14) & "."
      Celda = ofv.GenCelda("","ttt","right","1%","","","",sInput,"si")


    'Evento
    sInput = rs(15) & ""
    Celda  = Celda & ofv.GenCelda("","ttt","","30%","","","",sInput,"si")

    'Sector origen
    if bModificar then

      Strsql = "Select Eco.eco,Eco.descri,a.cliente From Objsector, Eco "
      strsql = strsql & " inner join clienteplan a on (eco.ecoUN=a.id) "
      strsql = strsql & " where eco.estado=3 and "
      Strsql = Strsql & " Objsector.codsector = Eco.eco and codobjeto = " & nCodigo
    ''  sInput = ofv.GenerarCombo("edEsquemas",rs(4).name,rs(4).value,Strsql,cn,0,1,"Sector1")
     Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
       scomK = ""'"null" & Chr(9)

      scomD = "" '"" & Chr(9)
      SCOMSEP =""
      do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
      if Session("EMPMUN")  = "S" then
       scomd=scomd & SCOMSEP & rsperfil(2) & "&nbsp;-&nbsp;" & rsperfil(1)
	   else
	    scomd=scomd & SCOMSEP & rsperfil(1)
		end if
       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop

      sInput =  ofv.gencomboFX("edEsquemas", rs(4).name,rs(4).value , scomK, scomD, "")
    else
      sInput = ofv.GenerarInput("",rs(7),"readonly","","40","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","ut","","30%","","","",sInput,"si")

    'Sector Destino

    if bModificar then

      Strsql = "Select Eco.eco,Eco.descri,a.cliente From Objsector, Eco  "
      strsql = strsql & " inner join clienteplan a on (eco.ecoUN=a.id) "
       strsql = strsql & " where eco.estado=3 and "
      Strsql = Strsql & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo
     ''  sInput = ofv.GenerarCombo("edEsquemas",rs(5).name,rs(5).value,Strsql,cn,0,1,"Sector2")
       Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
       scomK = ""'"null" & Chr(9)

      scomD = "" '"" & Chr(9)
      SCOMSEP =""
      do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
	   if Session("EMPMUN")  = "S" then
       scomd=scomd & SCOMSEP & rsperfil(2) & "&nbsp;-&nbsp;" & rsperfil(1)
	   else
	    scomd=scomd & SCOMSEP & rsperfil(1)
		end if

       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop

      sInput =  ofv.gencomboFX("edEsquemas", rs(5).name,rs(5).value , scomK, scomD, "")
    else
      sInput = ofv.GenerarInput("",rs(14),"readonly","","40","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","ut","","30%","","","",sInput,"si")

      'Texto de Referencia
      sID = "edEsquemaRef.asp?"
      sID = sID & "ID=" & rs(0) & "&Tabla=Objesquema&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHref = sID


    sImagenP ="&#0232;"
    if rs(6) <> "" then
       sImagenP = "&#0108;" & sImagenP
    end if

    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='" & sID
    sinput = sinput  & "' title='Texto de Referencia del Esquema' >" & sImagenP & "</a>"
    Celda = Celda & ofv.gencelda("","ft","center","","10%","","",sInput,"si")

    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objesquema&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if


    rs.movenext
  else
    nID = "no"
  end if
  Fila = ofv.GenTabla("","","","92%","","0","","0","0","","",Fila,"si")
  sHTML = sHTML & ofv.GenForm("edEsquemas","",sAccion,"","post","",Fila,"si")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ESQUEMA CONTABLE



  'Boton test
  sTest = "T"
  sAccion = "CoEsquemas.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sAccion = sAccion & "&Menu=PORESTADO&Sector=3&origen=ED&CantRegAMover=A&selcom=" & nID & "&test=" & sTest
  TEST = ofv.BotonTest(sAccion)

  'Boton Volver
  sAccion = "edesquemas.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Indice = ofv.BotonVolver(sAccion)





  'Encabezados
  Celda = ofv.GenCelda("","ttt","","90%","","","","Esquema Contable","si")
 ' Celda = celda & indice
 ' Celda = celda & test
  Fila = ofv.GenRow("","","","","30","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","95%","","0","","0","0","","",Fila,"si")



  if nID <> "no" then
    'Encabezados

  Celda = ofv.GenCelda("","ttr","","","","","","&nbsp;","si")
  Celda = celda & ofv.GenCelda("","ttr","","","","","","&nbsp;","si")
  Celda = celda & ofv.GenCelda("","ttr","","","","","","Cuenta","si")
  Celda = celda & ofv.GenCelda("","ttr","","","","","","Sentido","si")
  Celda = celda & ofv.GenCelda("","ttr colspan=3 ","","","","","","Referencia","si")
  Fila = ofv.GenRow("","","","","25","","",Celda,"si")






    Strsql = "SELECT OBJESQUEMA2.*, b.codinx2 "

    strsql = strsql & " FROM OBJESQUEMA2 "
    strsql = strsql & " inner join OBJESQUEMA on (OBJESQUEMA2.codesquema = OBJESQUEMA.objid) "

   ' strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (OBJESQUEMA2.objid = b.codinx2) "
   '  strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (OBJESQUEMA2.objid = b.codinx2) "
     strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast (OBJESQUEMA2.objid as nvarchar) = b.codinx2) "
    Strsql = Strsql & " where "
    Strsql = Strsql & " Objesquema2.codesquema=" & nID
    if bVersion = "NO" then
       strsql = strsql & " and (isnull(OBJESQUEMA.modi,0) = 0 or (isnull(OBJESQUEMA.modi,0) <> 0 and OBJESQUEMA.tipmodi <> 'E')) "
       strsql = strsql & " and (isnull(OBJESQUEMA2.modi,0) = 0 or (isnull(OBJESQUEMA2.modi,0) <> 0 and OBJESQUEMA2.tipmodi <> 'E')) "
    end if

'''response.write "<br>" & strsql & "<br>"

    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then rs2.move(CantRegAMover2)
    posicion = 0
    do until rs2.eof or posicion < CantRegMostrar2
      nCuenta = rs2(0)
      sForm = "edEsquemas" & rs2(0)

      sAccion = ""
      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & rs2(0) & "&Tabla=Objesquema2&volver=" & vuelta
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      end if



      'Eliminar
      sHref = ""
      if rs2(7) And rs2(8) = "E" then
''      sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        sImage = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
        Celda = ofv.GenCelda("","","","","","","",sImage,"si")
      else
        if bModificar then
          if rs2(7) And rs2(8) = "A" then
             sHref = "asp/EliminarEN.asp?ID=" & rs2(0) & "&Tabla=Objesquema2&volver=" & vuelta
             sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
          else
             sHref = "asp/ActualizarEN.asp?ID=" & rs2(0) & "&Tabla=Objesquema2&volver=" & vuelta
             sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          end if
        end if
        if sHref <> "" then
          Celda = ofv.BotonEliminar("Imputacion",sHref)
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


      if (Nestado=5  or Nestado=6) and (rs2(15) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs2(15)  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"

         ''''''agregado para color del cuadrito de Observaciones version 4.6

		       if wComentario = "" then
		          sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
		   '''''''sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"

		       else
		          sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
		       end if
         ''''''
             Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")
      	       Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")

      end if


      Celda = Celda & Celdaobs



      'Cuentas Contables
      if bModificar then
        sqlCuenta = "SELECT * From csesqcont where estado=3  "
        sqlCuenta = sqlCuenta & "Order By descri"
        sInput = ofv.GenerarCombo(sForm & "X1","codcuenta",rs2(4),sqlCuenta,cn,0,1,"")
        sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")
      else
       ' sInput = ofv.GenerarInput("",rs2(16) & " - " & rs2(15),"readonly","","80","dt","")
      end if
      Celda = Celda & ofv.GenCelda("","","","80%","","","",sInput,"si")

      'Sentido
    scomK = "Debe" & chr(9) & "Haber"
    scomD = "DEBE" & chr(9) & "HABER"

    if bModificar then
        sInput = ofv.gencomboFX(sForm & "X2",rs2(5).name,rs2(5),scomK,scomD," ")
        sInput = ofv.GenForm(sForm & "X2","",sAccion,"","post","",sInput,"si")
    else
        sInput = ofv.gencomboFX(sForm,rs2(5).name,rs2(5),scomK,scomD,"no")
    end if

      Celda = Celda & ofv.GenCelda("","","","10%","","","",sInput,"si")

      'Texto de Referencia
      sID = "edEsquemaRef.asp?"
      sID = sID & "ID=" & rs2(0) & "&Tabla=Objesquema2&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHref = sID


    sImagenP = "&#0232;"
    if rs2(6) <> "" then
       sImagenP = "&#0108;" & sImagenP
    end if

    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='" & sID
    sinput = sinput  & "' title='Texto de Referencia' >" & sImagenP & "</a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


  'Link a paginas internas/externas

         sID = "EdGENLnk.asp?"
         sID = sID & "ID=" & rs2(0) & "&TABLA=OBJESQUEMA2&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo


         if not isnull(rs2(9))  then
            if  clng(rs2(9)) > 0 then
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

     'Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
     xxx = ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
''   XXX= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")

      'sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",sInput,"si")
      celda = celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","2%","","","",xxx,"si")

''   Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","center","5%","","","",XXX,"si")





      'Version
      sSelect = ""
      if bVersion = "SI" then

      if rs2(7) then sSelect = "selected"
      if sSelect <> "" then
        if bModificar then
          if rs2(8) = "A" then
             sInput1 = "Agregada"
          elseif rs2(8) = "E" then
             sInput1 = "Eliminada"
          elseif rs2(8) = "M" then
             sInput1 = "Modificada"
          end if
          sMVal = "TRUE"
          if not rs2(7) then  sMVal = "FALSE"

          sInput = "<font class=at><select name=" & rs2(7).name & " value=" & sMVal
          sInput = sInput & " onchange=" & sForm & ".submit(); >"
          sInput = sInput & "<option value=FALSE>No</option>"
          sInput = sInput & "<option value=TRUE selected>" & sInput1
          sInput = sInput & "</option></select></font>"
        else
          if rs2(8) = "A" then
             sInput1 = "Agregada"
          elseif rs2(8) = "E" then
             sInput1 = "Eliminada"
          elseif rs2(8) = "M" then
             sInput1 = "Modificada"
          end if
          sInput = ofv.GenerarInput(rs2(7).name,sInput1,"readonly","","10","dt","")
        end if
        Celda = Celda & ofv.GenCelda("","uts","","","","","",sInput,"si")
      else
        Celda = Celda & ofv.GenCelda("","uts","","","","","","&nbsp;","si")

      end if
      else
        Celda = Celda & ofv.GenCelda("","uts","","","","","","&nbsp;","si")
      end if


      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")


      rs2.movenext
      posicion = posicion + 1
    loop
    sHTML = sHTML & ofv.GenTabla("","","","95%","","0","","0","0","","",Fila,"si")

    Agregar = ""

    if bModificar then
      'Boton Agregar
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = "AgregarCuenta.asp?Tabla=Objesquema2&volver=" & vuelta & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Esquema=" & nID & "&CantRegAMover="
      sAccion = sAccion & CantRegAMover & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)


    end if

   ''wwwbarra = ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","",Indice,Agregar,sRetornar)
     wwwbarra = ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","",Indice,Test,Agregar,sRetornar)


    call ofv.cerrarconsulta(rs2)
  end if

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si") '''aa onresize

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