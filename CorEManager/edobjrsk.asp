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

  set cn = ofv.conectar(ofv.strconn2)



  sHTML  = ofv.FormHeader(session("FormN"))
  vuelta = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  sOPCION = request.querystring("OPCION")

  IF SOPCION = "SI" THEN

     strsql = "Insert into OBJRSKREL (codobjeto, tipobjeto, codobjRSK,EVENRSK,rellink,reflink"
     strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "',0,0,0,'0')"
   '''    response.write strsql & "<br>"
     call ofv.crearconsultaEx(StrSql,cn,1,volver)


  END IF




  sImagenI3 = "<B><FONT FACE=CorFlaticon COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
  sImagenP = "<B><FONT FACE=CorFlaticon COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"

  'Encabezados
  Formulario = ""

  Celda2 = ofv.GenCelda("","TTt","","100%","","","","Documentos relacionados por eventos de perdidas","si")

  Fila2 = ofv.GenRow("","","","","","","",Celda2,"si")

  sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Fila2,"si")


  Strsql = "Select * From Objrskrel Where codobjeto=" & nCodigo & " Order By objid"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    sForm = "edEventos" & rs(0)

    Fila2 = ""

    'Eliminar
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objrskrel&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if





    'Documento Emanager

     sInstrumento = "Sin documento asociado ........"

    if  cstr(rs(6) & "") <> "" then

       Strsql = "SELECT * From objlnk where objid = " & rs(6)
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rs1.eof then
          wtipoL   = clng(rs1(6))
          wtipOBJ  = rs1(7)
          wUrl     = rs1(8)
          wUrlDesc = rs1(9)
          wtipo    = rs1(10)


          if trim(wtipo) = "I" or trim(wtipo) = "IOR" then

             sInstrumento = ofv.GenLink("","nt title='Ver documento asociado.' target='_blank' ", wUrl,"","","","<b>" & wUrl & "</B>","si")


          else

           ''  sInstrumento = wUrl & " ." & wtipo & "."


             Strsql = "Select xdescri From csGenObjetos where codobjeto = " & wtipoL
     ''      response.write strsql & "<br>"
             set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
             call ofv.cerrarconsulta(rsInstrumento)

          end if


       end if
       call ofv.cerrarconsulta(rs1)




    end if
    sInput = sInstrumento

    if bModificar then


         sID = "EdGENRsk.asp?"
         sID = sID & "ID=" & rs(0) & "&TABLA=OBJRSKREL&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vuelta2 & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo



        sinput2 = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput2 = sinput2  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinput2 = sinput2  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinput2 = sinput2  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinput2 = sinput2  & "none" & chr(34) & ";' "
        sinput2 = sinput2  & "  href='" & sID & "'  "
        sinput2 = sinput2  & "  TITLE='Asociar/cambiar documento operativo.'  ><b>&nbsp;" & sImagenI3 & "...&nbsp;</b></a>"

    else
      sInput2 = ""
    end if
    Celda2 = ofv.GenCelda("","ttR COLSPAN=2","","20%","","","","DOCUMENTO ASOCIADO","si")

    Fila2 = ofv.GenRow("","","","","","","",Celda2,"si")

    Celda2 = ofv.GenCelda("","vt COLSPAN=2","","20%","","","",sInput & sInput2,"si")

    Fila2 = Fila2 & ofv.GenRow("","AAA BGCOLOR='#F0F0F0' ","","","","","",Celda2,"si")

    Celda2 = ofv.GenCelda("","TTR ","","50%","","","","OBSERVACIONES","si")

    Celda2 = Celda2 & ofv.GenCelda("","TTR ","","50%","","","","EVENTO DE PERDIDA","si")

    Fila2 = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")


    Celda2 = ""
    'observacion

    IF cstr(rs(6) & "") = "0" THEN

      celda3 = "Debe asociar un documento operativo."

    ELSE

      celda3 = "<font class=at><textarea name=" & rs(3).name & " rows=3 cols=40 "
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


    Celda2 = ofv.GenCelda("","vt ","CENTER","50%","","","",celda3,"si")

    'Proceso

    IF cstr(rs(6) & "") = "0" THEN

        sInput = "-------------"

    ELSE



    if  bModificar then
        Strsql = "Select * From CSEVENTORSK Order By GENDESRSK"
        sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,Strsql,cn,0,6,"")
    else
        if  cstr(rs(5) & "") <> "" then
            Strsql = "Select GENDESRSK From CSEVENTORSK Where evento = 0" & rs(5)
            set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
            call ofv.cerrarconsulta(rsInstrumento)
        end if
        sInput = sInstrumento
    end if

   end if


    Celda2 = Celda2 & ofv.GenCelda("","Ut ","center","50%","","","",sInput,"si")

    Fila2 = Fila2 & ofv.GenRow("","","","","","","",Celda2,"si")
    stabla =  ofv.GenTabla("","AAA ","","98%","","0","","0","0","","",fila2,"si")



    Celda = Celda & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#336699;'  ","","95%","","","",stabla,"si")

         sID = "Edobjrskdet.asp?"
         sID = sID & "codrel=" & rs(0) & "&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo


        sinput2 = "<a class=nt style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput2 = sinput2  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinput2 = sinput2  & "#6699cc" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinput2 = sinput2  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinput2 = sinput2  & "none" & chr(34) & ";' "
        sinput2 = sinput2  & "  href='" & sID & "'  "
        sinput2 = sinput2  & "  TITLE='Asociar items de ponderacion.'  ><b>&nbsp;" & sImagenP & "&nbsp;</b></a>"

    Celda = Celda & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#336699;'  ","","5%","","","",sinput2,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objrskrel&volver=" & vuelta
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
    sAccion = "EDOBJRSK.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = sAccion & "&OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)

  end if

    sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,"no","")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>