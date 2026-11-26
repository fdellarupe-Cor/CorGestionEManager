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
  sTipo   = request.querystring("Tipobjeto")
  sTabla  = request.querystring("Tabla")

 '' response.write " sTabla  " & sTabla
 '' response.write " sSubTit " & sSubTit


 histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

  if  stabla <> "" then
      session("ESQTABLA") = stabla
  else
      stabla = session("ESQTABLA")
  end if

 '' response.write sTabla & " sTabla1 "

  sOpcion = request.querystring("OPC")

  if sOpcion <> "" then sOpcion = Ucase(sOpcion)

  sconte  = request.querystring("OConte")

  if sconte = "" then sconte = "NO"

  sconteK = request.querystring("OConteK")
  sconteD = request.querystring("OConteD")

  CantRegAMover    = cdbl(request.querystring("CantRegAMover"))
  CantRegAMover2   = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar  = 1
  CantRegAMostrar2 = 1


  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover) & "_Oconte=" & sconte  & "_OconteK=" & sconteK
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2) & "_Oconte=" & sconte  & "_OconteK=" & sconteK
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2) & "_Oconte=" & sconte  & "_OconteK=" & sconteK

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  sImageobs = "&nbsp;<font face='corflaticon' color='#990000' size=3 >&#095;</font>"

  set cn  = ofv.conectar(ofv.strconn2)


  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    nestado = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

''  response.write sTabla & " sTabla2 "


  Strsql  = "Select * From ObjTabdef Where objid=" & sTabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

''  response.write strsql & " aaaa "

  if not rs.eof then
    sForm = "edTablas" & rs(0)
    sid = rs(0)


      Celda = ofv.GenCelda("","tt","","1%","","","","Tabla:&nbsp;","si")

      sInput = rs(5) & "."
      Celda = Celda & ofv.GenCelda("","ttt","","1%","","","",sInput,"si")

      sInput = rs(4)
      Celda = Celda & ofv.GenCelda("","ttt","","98%","","","",sInput,"si")

      Fila = ofv.GenRow("","","","","25","","",Celda,"si")

  else
     sid = "NO"
  end if

  sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80%","","0","","0","0","","",fila,"si") '''agrego tabla aaa, para cabecera de tablas


  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & ""

      Fila = ""

  if sid <> "NO" then
'''''''''''''''''''''''''''''''''''''''''''''''

  'Encabezados
 '' Formulario1 = ofv.Encabezados("#","Subtitulo")

    Celda = ofv.GenCelda("","ttr","","","","","","&nbsp;","si")
    sInput = "#"
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sInput = "Subtitulo"
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sInput = "Encabezados"
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sInput = "Contenidos"
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sInput = "Texto Adicional"
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    sInput = "&nbsp;"
    Celda = Celda & ofv.GenCelda("","ttr","","","","","",sInput,"si")
    Formulario1 = ofv.GenRow("","","","","","","",Celda,"si")

  set cn = ofv.conectar(ofv.strconn2)

    strsql = "SELECT Objtabblk.*, b.codinx1 From Objtabblk "
    'strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objtabblk.objid = b.codinx1) "
    'strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where   codsol = " & session("codSol") & ") b on (Objtabblk.objid = b.codinx1) "
    strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where  TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on (cast ( Objtabblk.objid as nvarchar)= b.codinx1) "
    strsql = strsql & " Where Tabdef = " & sID & " order by tabblksec "

  ''  response.write strsql & "<br>"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     Strsqlx = "insert into Objtabblk (codobjeto,tipobjeto,Tabdef,tabblksec) values ("
     Strsqlx = Strsqlx & nCodigo & ",'" & sTipo & "'," & sID & ",10)"

     call ofv.crearconsultaEx(StrSqlx,cn,1,vueltac)
  end if
  call ofv.cerrarconsulta(rs)


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  Iform=0
  if not rs.eof then
     do until rs.eof
    nID = rs(0)

      if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vueltac
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      end if

    'Eliminar
    sHref = ""
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vueltac
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
   '' Celda = ofv.BotonEliminar("Subtitulo",sHref)
      Celda = ofv.BotonEliminar("Eliminacion",sHref)
      Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Orden
    sInput = ofv.GenerarInput(rs(4).name,rs(4),"text","edBloques" & Iform,"3","dt","")
    sInput = ofv.GenForm("edBloques" & Iform,"",sAccion,"","post","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Iform= Iform +1
    'Titulo

    sInput = ofv.GenerarInput(rs(5).name,rs(5),"textsb","edBloques" & Iform,"90","dt","")
    sInput = ofv.GenForm("edBloques" & Iform,"",sAccion,"","post","",sInput,"si")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='EdTablasBLKTIT.asp?Tabla=" & sTabla
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SubTit=" & rs(0)
    sinput = sinput  & "' title='Encabezados de " & rs(5) & "' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='EdTablasBLKCON.asp?Tabla=" & sTabla
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SubTit=" & rs(0)
    sinput = sinput  & "' title='Contenidos de " & rs(5) & "' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='EdTablasBLKTXT.asp?Tabla=" & sTabla
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SubTit=" & rs(0)
    sinput = sinput  & "' title='Texto adicional de " & rs(5) & "' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")


      if (Nestado=5  or Nestado=6) and (rs(12) & "") <> "" and bModificar then

               stitle = "Hay observaciones para este grupo."
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & sImageobs & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut title='" & stitle & "'  valign=top ","","","1","","",sInput,"si")
      end if


      Celda = Celda & Celdaobs



  ''  if bModificar then
  ''    sAccion = "asp/ActualizarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vueltac
  ''    sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  ''    Formulario = ofv.GenForm("edBloques" & Iform,"",sAccion,"","post","",celda,"si")
  ''  else
  ''    Formulario = ofv.GenForm("","","","","post","",celda,"si")
  ''  end if
''


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
    Iform=Iform +1
    loop
  else
    nID = "no"
  end if
  'sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario1 & fila,"si")
  sHTML = sHTML & ""
  sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","80%","","0","","0","0","","",Formulario1&fila,"si") '''agrego clase aaa, para contenedor de tablas, encabezados, contenido y texto adicional 12-03-26

  sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Agregar   = ""

  'Botonera
  if bModificar then
     sAccion   = "AgregarBloque.asp?Tabla=Objtabblk&volver=" & vueltac & "&Codobjeto="
     sAccion   = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & sID
     Agregar   = ofv.BotonAgregar(sAccion)
  end if

  sAccion  = "EdTablas.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sVolver  = ofv.BotonVolver(sAccion)

  'Boton test
 'sAccion   = "CoTablasCons.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sAccion= "Consultar.asp?Codigo=" & nCodigo & "&Tipo=" & STIPO &"&niv1=9" & "&niv2=" & sTabla
  Test      = ofv.BotonTest(sAccion)


  xxxx = ofv.botoneraESSX("5",rs,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","",Test,Agregar,sVolver,sRetornar)
''xxxx = ofv.botoneraESSX("5",rs,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","",Agregar,sVolver,sRetornar)

  call ofv.cerrarconsulta(rs)

end if

  call ofv.cerrarconn(cn)


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")'' se quito aaa onresize

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & xxxx

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>