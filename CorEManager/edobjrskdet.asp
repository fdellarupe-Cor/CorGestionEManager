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
  sTipo = request.querystring("Tipobjeto")
  codrel = request.querystring("codrel")

  if codrel <> "" then
     session("wkcodrel") = codrel
  else
     codrel = session("wkcodrel")
  end if

  set cn = ofv.conectar(ofv.strconn2)



  sHTML  = ofv.FormHeader(session("FormN"))
  vuelta = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  sOPCION = request.querystring("OPCION")

  IF SOPCION = "SI" THEN

     strsql = "Insert into OBJRSKRELDET (codobjeto, tipobjeto, codRSKREL,itemrsk,ponderacion"
     strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "','" & codrel & "',0,0)"
   '''    response.write strsql & "<br>"
     call ofv.crearconsultaEx(StrSql,cn,1,volver)


  END IF




  sImagenI3 = "<B><FONT FACE=CorFlaticon COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
  sImagenP = "<B><FONT FACE=CorFlaticon COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"

  'Encabezados
  Formulario = ""


  Strsql = "Select * From Objrskrel Where objid=" & codrel
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     documento = rs(6)
     evento    = rs(5)

       Strsql = "SELECT * From objlnk where objid = " & rs(6)
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs1.eof then
          wtipoL   = clng(rs1(6))
          wtipOBJ  = rs1(7)
          wUrl     = rs1(8)
          wUrlDesc = rs1(9)
          wtipo    = rs1(10)


          if trim(wtipo) = "I" or trim(wtipo) = "IOR" then

             sInstrumento = wUrl


          else

             Strsql = "Select xdescri From csGenObjetos where codobjeto = " & wtipoL
     ''      response.write strsql & "<br>"
             set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
             call ofv.cerrarconsulta(rsInstrumento)

          end if


       end if
       call ofv.cerrarconsulta(rs1)

       Strsql = "Select GENDESRSK From CSEVENTORSK Where evento = 0" & rs(5)
       set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rsInstrumento.eof then sInstrumento2 = rsInstrumento(0)
       call ofv.cerrarconsulta(rsInstrumento)

'''''''        sInstrumento2 = Strsql



  end if




  Celda2 = ofv.GenCelda("","TTt colspan=2 ","","100%","","","","Items de ponderacion","si")

  Fila2 = ofv.GenRow("","","","","","","",Celda2,"si")


  Celda2 = ofv.GenCelda("","TTr","","20%","","","","Documento","si")
  Celda2 = Celda2 & ofv.GenCelda("","ut","","80%","","","",sInstrumento,"si")
  Fila2 = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")

  Celda2 = ofv.GenCelda("","TTr","","20%","","","","Evento","si")
  Celda2 = Celda2 & ofv.GenCelda("","ut","","80%","","","",sInstrumento2,"si")
  Fila2 = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")
  sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Fila2,"si")


  Strsql = "Select * From Objrskreldet Where codobjeto=" & nCodigo & " and codrskrel = '" & codrel & "' Order By objid"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    sForm = "edEventos" & rs(0)

    Fila2 = ""

    'Eliminar
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objrskreldet&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if


  'item de riesgo

    if  bModificar then
        Strsql = "Select * From rskitem Order By ponderador,descri"
        sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,Strsql,cn,0,1,"")
    else
        if  cstr(rs(5) & "") <> "" then
            Strsql = "Select descri From rskitem Where ID = '" & rs(5) & "' "
            set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
            call ofv.cerrarconsulta(rsInstrumento)
        end if
        sInput = sInstrumento
    end if


  Celda2 = ofv.GenCelda("","TTr","","","","","","item de riesgo","si")
  Fila2  = ofv.GenRow("","","","","","","",Celda2,"si")


    Celda2 = ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#eeeeee;' ","","100%","","","",sinput,"si")

    Fila2 = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")

  Celda2 = ofv.GenCelda("","TTr","","","","","","observaciones","si")
  Fila2  = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")


    Celda2 = ""
    'observacion

    IF cstr(rs(5) & "") = "0" THEN

      celda3 = "Debe asociar un item de riesgo."

    ELSE

      celda3 = "<font class=at><textarea name=" & rs(3).name & " rows=2 cols=40 "
      if bModificar then
        celda3 = celda3 & " onblur=" & sform & ".submit() "
        celda3 = celda3 & " onkeypress=" & chr(34)
        celda3 = celda3 & " if (this.value.length > 450) {alert('El texto es demasiado"
        celda3 = celda3 & " extenso'); return false;}" & chr(34)
      else
        celda3 = celda3 & " readonly"
      end if

     celda3 = celda3 & " >" & rs(3) & "</textarea></font>"

   END IF


    Celda2 = ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#eeeeee;' ","center","100%","","","",celda3,"si")
    Fila2 = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")
    stabla =  ofv.GenTabla("","AAA ","","98%","","0","","0","0","","",fila2,"si")

    Celda = Celda & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#336699;'  ","","85%","","","",stabla,"si")


    sInput = ofv.GenerarInput(rs(6).name,rs(6),"text",sForm,"50","dtn","")

    Celda = Celda & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#336699;'  ","","10%","","","",sinput,"si")

    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objrskreldet&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
    else
      Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
    end if

    rs.movenext
  loop

  sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario,"si")

  'Botonera
  if bModificar then
    sAccion = "EDOBJRSKdet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = sAccion & "&OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)

  end if

  sAccion = "EDOBJRSK.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  volver = ofv.BotonVolver(sAccion)



    sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,volver,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>