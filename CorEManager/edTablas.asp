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
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
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

  sImageobs = "&nbsp;<font face='corflaticon' color='#990000' size=3 >&#095;</font>"


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------


  'Encabezados
    if sconte = "SI" then
       sdtab = "Edicion de Tabla: " & sconteD & " asociada al contenido"
       Celda = ofv.GenCelda("","tt","center colspan=5","","","","",sdtab,"si")
       Fila = ofv.GenRow("","","","","","","",Celda,"si")
    end if

    Fila = fila & ofv.Encabezados("#","Titulo","Modelo","Formato"," ")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    nestado = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "delete From Objtabcmb Where codobjeto = " & nCodigo & " "
  strsql = strsql & " and cmbcol = 'N' and cmbrow = 'N' and cmbstyl = 'N' "
  call ofv.crearconsultaEx(StrSql,cn,1,parametros)



  if sconte = "SI" then
     Strsql = "Select ObjTabdef.*, b.codinx3 From ObjTabdef "
     'strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (ObjTabdef.objid = b.codinx3) "
    ' strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where codsol = " & session("codSol") & ") b on (ObjTabdef.objid = b.codinx3) "
     strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on (cast ( ObjTabdef.objid as nvarchar)= b.codinx3) "
     Strsql = Strsql & " Where codobjeto=" & nCodigo & " and tabdettablas = '"
     Strsql = Strsql & sconteK & "' "
  else
     Strsql = "Select ObjTabdef.*, b.codinx3 From ObjTabdef "
     'strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (ObjTabdef.objid = b.codinx3) "
     'strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where   codsol = " & session("codSol") & ") b on (ObjTabdef.objid = b.codinx3) "
     strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where  TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast ( ObjTabdef.objid as nvarchar) = b.codinx3) "
     Strsql = Strsql & " Where codobjeto=" & nCodigo & " and tabdettablas is null "
     Strsql = Strsql & " Order By tabsecuen"
  end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof and sconte = "SI" then
     strsql = "Insert into Objtabdef (codobjeto,tipobjeto,tabtitulo,tabsecuen,tabdettablas) "
     Strsql = Strsql & " values (" & nCodigo & ",'" & sTipo & "','"
     Strsql = Strsql & ofv.convertircar(sconteD,"_"," ") & "',10,'"
     Strsql = Strsql & sconteK & "')"
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     Strsql = "Select ObjTabdef.*, b.codinx3 From ObjTabdef "
     'strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (ObjTabdef.objid = b.codinx3) "
     strsql = strsql & " left join (select distinct codinx3 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (ObjTabdef.objid = b.codinx3) "
     Strsql = Strsql & " Where codobjeto=" & nCodigo & " and tabdettablas = '"
     Strsql = Strsql & sconteK & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if

  if not rs.eof then rs.move(CantRegAMover)

  do until  rs.eof
    sForm = "edTablas" & rs(0)
    sid = rs(0)

  

     if bModificar then
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objtabdef&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
   
   end if

    'Eliminar
    sHref = ""

    if bModificar and sconte <> "SI" then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objtabdef&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if

    if sHref <> "" then
''    Celda = ofv.BotonEliminar("Tabla",sHref)
      Celda = ofv.BotonEliminar("Eliminacion",sHref)
      Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Secuencia
    if bModificar and sconte <> "SI" then
      sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm & "f1","3","dtn","SI")
      sInput = ofv.GenForm(sForm & "f1","",sAccion,"","post","",sInput,"si")  
    else      
      'sInput = ofv.GenerarInput("",rs(5),"readonly","","3","dtn","SI")
      sInput = rs(5)
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Descripcion
    if bModificar then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm & "f2","70","dt","")
      sInput = ofv.GenForm(sForm & "f2","",sAccion,"","post","",sInput,"si")
    else
     '' sInput = ofv.GenerarInput("",rs(4),"readonly","","70","dt","")
      sInput = rs(4)
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Perfil
    Strsql = "Select * From Tabperfil Order By tabdescri"
    

       scomKK = ""
    scomDD = ""
  ''  response.write strsql & "<br>" & chr(13)
    set rsniv = ofv.crearconsultaEx(Strsql,cn,1,parametros)
    DO UNTIL rsniv.eof
       scomKK = scomKK & rsniv(0) & chr(9)
       scomDD = scomDD & rsniv(1) & chr(9)
       RSNIV.MOVENEXT
    LOOP
    call ofv.cerrarconsulta(rsniv)

      sInput = ofv.gencomboFX(sForm & "f3",rs(3).name,rs(3).value,scomKK,scomDD,"")
      sInput = ofv.GenForm(sForm & "f3","",sAccion,"","post","",sInput,"si")
      'sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,Strsql,cn,0,1,"")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

     if BrowserName()="IE" or BrowserName()="CHROME" THEN
      sImagenP =""
    else
    sImagenP = ""
    end if
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='EdTablasBLK.asp?Tabla=" & rs(0)
    sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sinput = sinput  & "' title='Formato de " & rs(4) & "' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")


      if (Nestado=5  or Nestado=6) and (rs(13) & "") <> "" and bModificar then

               stitle = "Hay observaciones para este grupo."
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & sImageobs & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut title='" & stitle & "'  valign=top ","","","1","","",sInput,"si")

      end if


      Celda = Celda & Celdaobs


''  if bModificar then
''     sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objtabdef&volver=" & vuelta
''    sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
''    Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")
''  else
''    Formulario = ofv.GenForm("","","","","post","",Celda,"si")
''   end if





   '' Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")
    Fila = fila & ofv.GenRow("","","","","","","",celda,"si")

    rs.movenext
    loop

   sHTML = sHTML & ""
    sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")

  'Botonera
    Agregar = ""

   sAccion  = "EdTablas.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
   sVolver  = ofv.BotonVolver(sAccion)

    sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = "AgregarTablas.asp?Tabla=ObjTabdef&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
  if bModificar and sconte <> "SI"  then
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  'Boton test
  'sAccion   = "CoTablasCons.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
   sAccion= "Consultar.asp?Codigo=" & nCodigo & "&Tipo=" & STIPO &"&niv1=9" & "&niv2=" & nID
  Formul    = ofv.BotonTest(sAccion)


  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

        ' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if




'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 ''sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","",Listar,Agregar,Formul,sRetornar)
  
   sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","",Listar,Agregar,Formul,Volver)



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>