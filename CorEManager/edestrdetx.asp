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

  CantRegMostrar = 10
  CantRegMover = cdbl(request.querystring("CantRegMover"))

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = request.querystring("ID")
  pmk = request.querystring("pmk")
  OSec = request.querystring("OSec")
  if OSec = "" then OSec = "NO"

  wposi = 0
  wpossal = 0

     session("comp31") = "TODOS"
     session("comp32") = "JPG"

'  vuelta = ofv.vueltaaspEx("../",CantRegMostrar,CantRegMover) & "_ID=" & ID

  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID
  vuelta2 = ofv.vueltaaspEx("../",CantRegMostrar2,CantRegMover2) & "_ID=" & ID
  vueltaC = ofv.vueltaaspExC("../",CantRegMover,CantRegMover2) & "_ID=" & ID

  eststt = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  eststa = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst1 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst2 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst3 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst4 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst5 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  esttxt = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#f7eedd;} "

  if pmk = "" then
     sfuncion = ""
  else
    sfuncion = " onload=" & chr(34) & "window.scrollTo(document.edestrdet" & pmk & ".offsetParent.offsetLeft,document.edestrdet" & pmk & ".offsetParent.offsetTop);" & chr(34)


  end if

  response.write ofv.menuheader("",sfuncion)

 ' response.write chr(13) & "<script LANGUAGE=Javascript >" & chr(13)
 ' response.write " var f='X'; " & chr(13)
 ' response.write "function ssss(t){" & chr(13)
 ' response.write " { if (t=='N') { " & chr(13)
 ' response.write "if (f!='X'){" & chr(13)
 ' response.write "window.blur();window.focus()}}" & chr(13)
  'response.write  "else {" & chr(13)
 ' response.write  "if (t.name!='dettexto'){ window.blur();}}" & chr(13)
 ' response.write "  </script>" & chr(13)

'  Formulario = ofv.Encabezados("Sec","Estructura","","")

  set cn = ofv.conectar(ofv.strconn2)


  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "Select * From Objestr Where objid = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bsecuen = ""
  bDescri = " "
  if not rs.eof then
     bsecuen = rs(4)
     bDescri = rs(3)
  end if
  call ofv.cerrarconsulta(rs)


'  response.write "<HR size=1 width=100% color=#336699>"

  response.write "<TABLE     width=100%  marginwidth=0  cellspacing=0  cellpadding=3 >"

'  response.write ofv.GenTabla("","aaa bordercolor=#eeeeee rules=rows bgcolor=#ffffff style={border-style:solid;border-width:1;}  ","","100%","","0","","0","0","","","","no")

    'null
    Celda =  ofv.GenCelda("","tt valign=top  ","","","1","","","&nbsp;","si")

    'Agregar Titulos
    if bModificar then
       sInput = ofv.GenerarInput("","+","submit","","20","bt title='Agregar contenido al Capitulo/Titulo/subtitulo' ","")
       Celda1 = ofv.GenCelda("","tt valign=top  ","","","1","","",sInput,"si")
       sAccion = "AgregarEstDet.asp?ID=" & ID & "&Tabla=Objestdet&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&IDx=0&IDt=X"
    else
       Celda1 = ofv.GenCelda("","","","","1","","","&nbsp;","si")
       sAccion = ""
    end if
    Formulario = ofv.GenForm("AgreDet1","",sAccion,"","post","",Celda1,"si")

    'Propiedades
    if bModificar then
'       sInput = ofv.GenerarInput("","P.","submit","","20","bt","")

       sInput = ofv.GenerarInput("","CAP","submit","","20","bt title='Propiedades' ","")
       Celda2 = ofv.GenCelda("","tt valign=top  ","","","1","","",sInput,"si")
       sAccion = "EdestrPrty.asp?volver=" & vuelta & "&Codobjeto=" & nCodigo & "&Tipobjeto="
       sAccion = sAccion & sTipo & "&ID=" & ID
    else
       Celda2 = ofv.GenCelda("","","","","1","","","&nbsp;","si")
       sAccion = ""
    end if

    Formulario = Formulario & ofv.GenForm("Proper","",sAccion,"Grama","post","",Celda2,"si")

    Celda = Celda & Formulario

'    Celda = Celda & ofv.GenCelda("","tx","","5","","","","CAP" ,"si")

      if OSec <> "SI" then

        'Secuencia Oculta
         sInput = ofv.GenerarInput("","S+","submit","","20","bt title='Mostrar secuencia' ","")
         Celda1 = amk & sInput & "&nbsp;&nbsp;" & bDescri
         sAccion = "edEstrDet.asp?ID=" & ID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sAccion = sAccion & "&OSec=SI" & "&cantregmostrar=" & cantregmostrar & "&CantRegMover=" & CantRegMover
         Formulario1 =  ofv.GenForm("secDet1","",sAccion,"","post","",Celda1,"si")



      else

        'Secuencia Visible
         sInput = ofv.GenerarInput("","S-","submit","","20","bt title='Ocultar secuencia' ","")
         Celda1 = amk & sInput & "&nbsp;&nbsp;" & bDescri
         sAccion = "edEstrDet.asp?ID=" & ID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sAccion = sAccion & "&OSec=NO" & "&cantregmostrar=" & cantregmostrar & "&CantRegMover=" & CantRegMover
         Formulario1 =  ofv.GenForm("secDet1","",sAccion,"","post","",Celda1,"si")



      end if



   'Estructura


    Celda = Celda & ofv.GenCelda("","tt valign=top colspan=9 ","","60%","1","","",Formulario1,"si")
    Celda = Celda & ofv.GenCelda("","tt valign=top ","","50%","1","","","&nbsp;","si")


      Fila = ofv.GenRow("","","","","","","",celda,"si")
      response.write Fila




  strsql = "SELECT * from Objestdet where detcon = " & ID & " and dettipo = 'A' "
  if bVersion = "NO" then
     strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
  end if
  strsql = strsql & " order by detsecuen"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  sant = "X"
  bseci = 0
  bsecix = 0
  redim preserve bsec(7)
  for i=0 to 7
      bsec(i) = 0
  next
  amkeya = 0
  do while not rs.eof

      do while not rs.eof

         SCLAVE = RS(0)
         Stip = RS(5)
         smodi = rs(29)
         stipmodi = rs(30)
         sForm = "edestrdet" & SCLAVE




      'Secuencia
      select case stip

             case "B"
               '1031 titulo a modificar colgar de otro titulo
               '1094 id del nuevo titulo

               IF rs(0) = 1031 THEN

                  STRSQLXX = "UPDATE OBJESTDET SET DETDEPEN = 1094, DETTIPO = 'C' WHERE OBJID = 1031"
                  CALL ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)


                  strsqlXX = "SELECT OBJID from Objestdet where DETDEPEN = 1031 AND DETTIPO NOT IN ('F','G','H','I') "
                  set rsX1 = ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)
                  DO WHILE NOT RSX1.EOF
'''''''''''''''''' C
                     STRSQLXX = "UPDATE OBJESTDET SET DETTIPO = 'D' WHERE OBJID = " & RSX1(0)
                     CALL ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)


                  strsqlXX = "SELECT OBJID from Objestdet where DETDEPEN = " & RSX1(0) & " AND DETTIPO NOT IN ('F','G','H','I') "
                  set rsX2 = ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)
                  DO WHILE NOT RSX2.EOF
'''''''''''''''''' D
                     STRSQLXX = "UPDATE OBJESTDET SET DETTIPO = 'E' WHERE OBJID = " & RSX2(0)
                     CALL ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)

                  strsqlXX = "SELECT OBJID from Objestdet where DETDEPEN = " & RSX2(0) & " AND DETTIPO NOT IN ('F','G','H','I') "
                  set rsX3 = ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)
                  DO WHILE NOT RSX3.EOF
'''''''''''''''''' E
                     STRSQLXX = "UPDATE OBJESTDET SET DETTIPO = 'F' WHERE OBJID = " & RSX3(0)
                     CALL ofv.crearconsultaEx(StrSqlXX,cn,1,parametros)

                  RSX3.MOVENEXT
                  LOOP


                  RSX2.MOVENEXT
                  LOOP



                  RSX1.MOVENEXT
                  LOOP

               END IF
      END SELECT


      'Secuencia
      select case stip
             case "A"
                  btip = "TIT"
                  btipp = "TITULO"
                  bseci = 0
                  bsecix = 0
                  bsec(1) = rs(4)

             case "B"
                  btip = "ST1"
                  btipp = "SUBTITULO1"
                  bseci = 1
                  bsecix = 1
                  bsec(2) = rs(4)



             case "C"
                  btip = "ST2"
                  btipp = "SUBTITULO2"
                  bseci = 2
                  bsecix = 2
                  bsec(3) = rs(4)
             case "D"
                  btip = "ST3"
                  btipp = "SUBTITULO3"
                  bseci = 3
                  bsecix = 3
                  bsec(4) = rs(4)
             case "E"
                  btip = "ST4"
                  btipp = "SUBTITULO4"
                  bseci = 4
                  bsecix = 4
                  bsec(5) = rs(4)
             case "F"
                  btip = "ST5"
                  btipp = "SUBTITULO5"
                  bseci = 5
                  bsecix = 5
                  bsec(6) = rs(4)
             case "G"
                  btip = "TXT"
                  btipp = "TEXTO"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = rs(4)
             case "H"
                  btip = "IMG"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = rs(4)
             case "I"
                  btip = "TAB"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = rs(4)


      end select
      Amk = "<a name=" & chr(34) & "X" & rs(0) & chr(34) & "> "
      amkey = rs(0)

     'Eliminar
      sHref = ""
      if smodi And stipmodi = "E" then
        sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        Celda = ofv.GenCelda("","","","","","","",sImage,"si")
      else
        if bModificar then
           if cstr(amkeya) = cstr(0) then
               vlk = vuelta
           else
               vlk = vuelta & "_pmk=" & amkeya
           end if
          if smodi And stipmodi = "A" then
            sHref = "asp/EliminarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vlk
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
          else
            sHref = "asp/ActualizarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vlk
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          end if
        end if
        if sHref <> "" then
          Celda = ofv.BotonEliminar("Detalle",sHref)
        else
          Celda = ofv.GenCelda("","","","","","","","","si")
        end if
      end if

      Formulario1 = Celda

      if bModificar then
         if stip <> "G" and stip <> "H" and stip <> "I" Then
          'Agregar Titulos
          sInput = ofv.GenerarInput("","A" ,"submit ","","20","bt title='Agregar contenido al Capitulo/Titulo/subtitulo' ","")
          Celda = ofv.GenCelda("","","","","","","",sInput ,"si")
          vlk = vuelta & "_pmk=" & amkey
          sAccion = "AgregarEstDet.asp?ID=" & ID & "&Tabla=Objestdet&volver=" & vlk
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&IDx=" & sclave
          sAccion = sAccion & "&IDt=" & stip
          Formulario1 = Formulario1 & ofv.GenForm("AgreDet2  ","",sAccion,"","post","",Celda,"si")
         else
          Celda = ofv.GenCelda("","","","20","","","","&nbsp;","si")
          Formulario1 = Formulario1 & ofv.GenForm("AgreDet3","","","","post","",Celda,"si")
         end if
         if stip <> "H" and stip <> "I" Then
          'Propiedades
          if rs("detdefault") = "NO" then
             sInput = ofv.GenerarInput("",Btip & "*","submit","","20","dtb title='Propiedades' ","")
          else
             sInput = ofv.GenerarInput("",Btip,"submit","","20","bt title='Propiedades' ","")
          end if
          Celda = ofv.GenCelda("","","","","","","",sInput,"si")
          sAccion = "EdestrPrtycont.asp?volver=" & vuelta  & "&Codobjeto=" & nCodigo
          sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Part=SI&TIPO=" & btipp & "&ID=" & sclave
          Formulario1 = Formulario1 & ofv.GenForm("Proper","",sAccion,"_blank","post","",Celda,"si")
         else
          sInput = ofv.GenerarInput("",Btip,"submit","","20","bt","")
          Celda = ofv.GenCelda("","","","","","","",sInput,"si")
          Formulario1 = Formulario1 & ofv.GenForm("AgreDet3","","","","post","",Celda,"si")
         end if
      else
        Celda = ofv.GenCelda("","aaa colspan=2","","20","","","","&nbsp;","si")
        Formulario1 = Formulario1 & Celda
      end if




      BsecD = ""
      for i=1 to bsecix
      BsecD = BsecD & bsec(i) & "."
      next

      if  bsecix = 7 then
          bspan = ""
      else
          Bspan = "ut colspan=" & 7 - bsecix
      end if




      if OSec <> "SI" then

        'Secuencia Oculta


         sInput = ofv.GenerarInput(rs(4).name,rs(4),"hidden","","2","dt","")
         sInput = sInput & ofv.GenerarInput(rs(5).name,rs(5),"hidden","","2","dt","")
         sInput = sInput & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","2","dt","")


         Celda = ""
         Formulario2 = ofv.GenCelda("","","","1","","","",sInput & "&nbsp;","si")

      else

        'Secuencia Visible


         if bModificar then
            sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"2","dtn","SI")
         else
            sInput = ofv.GenerarInput("",rs(4),"readonly","","2","dtn","SI")
         end if

         sInput = sInput & ofv.GenerarInput(rs(5).name,rs(5),"hidden","","2","dt","")
         sInput = sInput & ofv.GenerarInput(rs(6).name,rs(6),"hidden","","2","dt","")

         Celda = ofv.GenCelda("","cc","","120","","","",BsecD & sInput,"si")
         Formulario2 = ""

      end if


      for i=1 to bsecix
      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      next

      strsql = "SELECT * from Objestdet where detcon = " & ID & " and detdepen = " & sclave
      if bVersion = "NO" then
         strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
      end if
      strsql = strsql & " order by detsecuen"


         select case stip
                case "A"
      sant = "A"
      'Descripcion

      sicK = "<font face='corflaticon' size=4 color=#993333 >&#0163;w</font>&nbsp;"
      if bModificar then

        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","90","dt","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

      rs.movenext
      set rsa = rs
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "B"
      sant = "B"
      'Descripcion
      sicK = "<font face='corflaticon' size=3 color=#993333 >&#0140;w</font>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","90","dt","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

      rs.movenext
      set rsb = rs
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


                case "C"
      sant = "C"
      'Descripcion
      sicK = "<font face='corflaticon' size=3 color=#993333 >&#0141;w</font>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","90","dt","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

      rs.movenext
      set rsc = rs
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "D"
      sant = "D"
      'Descripcion
      sicK = "<font face='corflaticon' size=3 color=#993333 >&#0142;w</font>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","90","dt","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

      rs.movenext
      set rsd = rs
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "E"
      sant = "E"
      'Descripcion
      sicK = "<font face='corflaticon' size=3 color=#993333 >&#0143;w</font>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","90","dt","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

      rs.movenext
      set rse = rs
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "F"
      sant = "F"
      'Descripcion
      sicK = "<font face='corflaticon' size=3 color=#993333 >&#0144;w</font>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"90","dt","")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","90","dt","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

      rs.movenext
      set rsf = rs
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "G"

      'Texto



      wstexto = rs(8)

      if isnull(wstexto)  then
'         response.write "longv " & len(wstexto) & " " & wstexto & "<br>"
         sInput = space(80)
         sinput = replace(sinput," ","-")
      elseif len(wstexto) > 100 then
 '        response.write "longg " & len(wstexto) & " " & wstexto & "<br>"
         sinput = left(wstexto,100) & "<font face='tahoma' size=1 color=#993333 ><b>...(+)</b></font>"
      else
         if wstexto = space(len(wstexto)) then wstexto = replace(wstexto," ","-")
 '        response.write "longc " & len(wstexto) & " " & wstexto & "<br>"
         sinput = wstexto
      end if

    '  wstexto = "<font face='tahoma' size=2 color=#000080 >" & sInput  & "</font>"
      wstexto = sinput

      wstextoLNK = rs("detimagen")

       if isnumeric(wstextoLNK) and wstextoLNK <> "0" then
          sicK = "<font face='corflaticon' size=4 color=#008000 >&#0194;4</font>&nbsp;"
       else
          sicK = "<font face='corflaticon' size=4 color=#993333 >&#0157;4</font>&nbsp;"
       end if

 '     sicK = "<font face='corflaticon' size=4 color=#993333 >&#0157;4</font>&nbsp;"

      addqry = "cantregmostrar=" & cantregmostrar & "&cantregmover=" & cantregmover
      addqry = addqry & "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec

      sinput = "<a class=cc href=" & chr(34) & "edestrdetTXT.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sinput = sinput & "&IDTXT=" & rs(0) & "&vuelve="
      sinput = sinput & replace(addqry,"&","_") & chr(34) & "   title='Ingresar/Modificar el texto' >" & sicK & wstexto & "</a>"


      Celda = Celda & ofv.GenCelda("",bspan & " valign=middle  ","","","","","",sInput,"si")
      rs.movenext

                case "H"

      'Descripcion

      sicK = "<font face='corflaticon' size=3 >1</font>"
      if rs(9) <> "" then
         estia=" style={font-size:7pt;font-family:tahoma,arial;color:#eeeeee;background-color:#000080;font-weight:600;border-style:outset;border-width:2;}"

         slink = "<a class=cc href=Coedestimg.asp?A=" & rs(16) & "&B=" & rs(15)
         slink = slink & "&R=" & rs(12) & "&S=" & server.urlencode(rs(9))
         slink = slink & " target=_blank  title='Ver Imagen' >"
'         slink = slink & "<img src=imagenes/botedit.jpg border=0>"
         slink = slink & sick
         slink = slink & "</a>"


      else
         slink = ""
      end if

      if bModificar then
         sreduc = "<font class=at ><select  align=top class=at  "
         sreduc = sreduc & "onchange='" & sForm
         sreduc = sreduc & ".submit(); return true;' name=" & rs(12).name & " >"
         if not  isnull(rs(12)) then
            sreduc = sreduc & "<option value=" & rs(12) & " selected >" & rs(12) & "</option>"
         else
            sreduc = sreduc & "<option value=100 selected >100</option>"
         end if
         sreduc = sreduc & "<option value=120>120</option>"
         sreduc = sreduc & "<option value=110>110</option>"
         sreduc = sreduc & "<option value=100>100</option>"
         sreduc = sreduc & "<option value=90>90</option>"
         sreduc = sreduc & "<option value=80>80</option>"
         sreduc = sreduc & "<option value=70>70</option>"
         sreduc = sreduc & "<option value=50>50</option>"
         sreduc = sreduc & "<option value=40>40</option>"
         sreduc = sreduc & "<option value=30>30</option>"
         sreduc = sreduc & "<option value=20>20</option>"
         sreduc = sreduc & "<option value=10>10</option>"
         sreduc = sreduc & "<option value=5>5</option>"
         sreduc = sreduc & "</select></font>"
      else
         if isnull(rs(12))   then
            sreduc = "100"
         else
             sreduc = rs(12)
         end if

         sreduc = ofv.GenerarInput("",sreduc,"readonly","","3","dt","SI")
      end if

      if bModificar then
         sbord = "<font class=at><select  class=at onchange='" & sForm
         sbord = sbord & ".submit(); return true;' name=" & rs(15).name & " >"
         sSelect = ""
         if not  isnull(rs(15))  then
'            sbord = sbord & "<option value=" & rs(15) & " selected >" & rs(15) & "</option>"
            if rs(15) = "SI" then sSelect = " selected "
         end if
         sbord = sbord & "<option value=NO>NO</option>"
         sbord = sbord & "<option value=SI " & sSelect & " >SI</option>"
         sbord = sbord & "</select></font>"
      else
         if isnull(rs(15))  then
            sbord = "NO"
         else
             sbord = rs(15)
         end if

         sbord = ofv.GenerarInput("",sbord,"readonly","","2","dt","SI")
      end if

      if bModificar then
         salng = "<font class=at><select  class=at onchange='" & sForm
         salng = salng & ".submit(); return true;' name=" & rs(16).name & " >"
         if not  isnull(rs(16)) then
            salng = salng & "<option value=" & rs(16) & " selected >" & rs(16) & "</option>"
         end if
         salng = salng & "<option value=IZQUIERDA>IZQUIERDA</option>"
         salng = salng & "<option value=CENTRADO>CENTRADO</option>"
         salng = salng & "<option value=DERECHA>DERECHA</option>"
         salng = salng & "</select></font>"
      else
         if isnull(rs(16))  then
            salng = "IZQUIERDA"
         else
             salng =  rs(16)
         end if

         salng = ofv.GenerarInput("",salng,"readonly","","10","dt","SI")
      end if

 '     if bModificar then
 '       sArchivos = ofv.Getfiles("Estructura")
 '       sInput = ofv.GenerarCombo3(sForm,sArchivos,rs(9).name,rs(9))
 '     else
  '      sInput = ofv.GenerarInput("",chr(34) & rs(9) & chr(34),"readonly","","80","dt","SI")
 '     end if

      wstexto = rs(9)

      if wstexto <> "" then
         wstexto = chr(34) & wstexto & chr(34)
      else
         wstexto = "-----------------"
      end if

'''      wstexto = "<font face='tahoma' size=2 color=#000080 ><u>" &  wstexto   & "</u></font>"

      wstexto = "<u>" &  wstexto   & "</u>"

      addqry = "cantregmostrar=" & cantregmostrar & "&cantregmover=" & cantregmover
      addqry = addqry & "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec

      sinput = "<a class=cc href=" & chr(34) & "edestrdetIMG.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sinput = sinput & "&IDTXT=" & rs(0) & "&vuelve="
      sinput = sinput & replace(addqry,"&","_") & chr(34) & " title='Asignar/Modificar Imagen' >" & wstexto & "</a>"



      sicK = "<font face='corflaticon' size=4 color=#993333 >&#0163;4</font>&nbsp;"

      sreduc = "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;MOSTRAR&nbsp;AL&nbsp;" & sreduc
      sreduc = sreduc & "%&nbsp;&nbsp;"
      sreduc = sreduc & "ALINEACI&Oacute;N&nbsp;" & salng & "&nbsp;&nbsp;"
      sreduc = sreduc & "MARCO&nbsp;" & sbord & " "
      sreduc = sreduc & ""

      sCelda = sicK & sInput & "&nbsp;&nbsp;"  & slink
      sCelda = sCelda & "&nbsp;&nbsp;" & sreduc  & "&nbsp;&nbsp;"
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sCelda,"si")
      rs.movenext

               case "I"

      'Tablas
         estlink = "edTablas.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo
         estlink = estlink & "&OConte=SI&OConteK=" & rs(10) & "&OConteD=" & ofv.convertircar(rs(7)," ","_")
         estia=" style={font-size:7pt;font-family:tahoma,arial;color:#eeeeee;font-weight:600;text-decoration:none;background-color:#000080;border-style:outset;border-width:1pt;}"
         slink = "<a class=nn href=" & chr(34) & estlink & chr(34) & " target=_blank " & estia & " ><img src=imagenes/botedit.jpg border=0></a>"

      if bModificar then
        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"80","nn" & esttxt,"")
      else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","80","nn" & esttxt,"SI")
      end if

      sicK = "<font face='corflaticon' size=4 color=#993333 >2</font>&nbsp;"

      sCelda = sicK & sInput & "&nbsp;" & slink

      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sCelda,"si")


      rs.movenext

          end select

      'Version
      sSelect = ""
      if bVersion = "SI" then
      if smodi then sSelect = "selected"
      if sSelect <> "" then
        if bModificar then

'           if stipmodi = "A" then
'              sInput1 = "Agregado"
'           elseif stipmodi = "E" then
'              sInput1 = "Eliminado"
'           elseif stipmodi = "M" then
'              sInput1 = "Modificado"
'           end if

           if stipmodi = "A" then
              sInput1 = "<b>A</b>"
           elseif stipmodi = "E" then
              sInput1 = "<b>E</b>"
           elseif stipmodi = "M" then
              sInput1 = "<b>M</b>"
           end if

          sInput = "<font class=at><select name=modi"
          sInput = sInput & " onchange=" & sForm & ".submit(); >"
          sInput = sInput & "<option value=False>No</option>"
          sInput = sInput & "<option value=True " & sSelect & ">" & sInput1
          sInput = sInput & "</option></select></font>"
        else
          sInput = ofv.GenerarInput("modi","Modificada","readonly","","10","dt","")
        end if

        Celda = Celda & ofv.GenCelda("","dt","","","","","",sInput,"si")
      else
        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if
      else
        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")


      if bModificar then

        vlk = vuelta & "_pmk=" & amkey
        sAccion = "asp/ActualizarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vlk
'        sAccion = "asp/ActualizarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vueltaC
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",celda,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",celda,"si")
      end if


      if wpossal = cantregmover then
         if wposi = cantregmostrar then
            Fila = ""
            exit do
         else
            Fila = ofv.GenRow("","","","","","","",Formulario1 & Formulario & Formulario2,"si")
            response.write Fila
            wposi = wposi + 1
         end if
      else
         Fila = ""
         wpossal = wpossal + 1
      end if



      amkeya = amkey
    loop

         if wposi = cantregmostrar then
            exit do
         end if

      ssale = "no"
      do until ssale = "si"

         if     sant = "A" then
            set rs = rsa
            sant = "X"
         elseif sant = "B" then
            set rs = rsb
            sant = "A"
         elseif sant = "C" then
            set rs = rsc
            sant = "B"
         elseif sant = "D" then
            set rs = rsd
            sant = "C"
         elseif sant = "E" then
            set rs = rse
            sant = "D"
         elseif sant = "F" then
            set rs = rsf
            sant = "E"
         end if

      if (not rs.eof) or sant = "X" then
         ssale = "si"
      end if

      loop


 loop




  celda = ""


      for i=1 to 9
      Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
      next

      if OSec = "SI" then
         Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
      end if


      Celda = Celda & ofv.GenCelda("","","","79%","","","","&nbsp;","si")

      if OSec <> "SI" then
         Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
      end if

      Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")


      Fila = ofv.GenRow("","","","","","","",celda,"si")
      response.write Fila

response.write ofv.GenFTabla()



addqry = "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec




response.write  "<TABLE  name=btnas id=btnas width=2%  marginwidth=0  marginheight=0  cellspacing=1  cellpadding=2 ><tr>"

if cantregmover > 0 then
wcantmov = cantregmover - cantregmostrar
response.write  "<FORM  name=Bretro ID=Bretro action="
response.write  chr(34) & "edestrdet.asp?cantregmostrar=" & cantregmostrar & "&CantRegMover=" & wcantmov & addqry & chr(34) & "  method=post >"
response.write  "<TD ><input  type=SUBMIT  size=10%  value='Anterior' class=bt ></TD>"
response.write  "</FORM>"
end if

if not (wposi < cantregmostrar) then
wcantmov = cantregmover + cantregmostrar
response.write  "<FORM  name=Bavan ID=Bavan action="
response.write  chr(34) & "edestrdet.asp?cantregmostrar=" & cantregmostrar & "&CantRegMover=" & wcantmov & addqry & chr(34) & "  method=post >"
response.write  "<TD ><input  type=SUBMIT  size=10%  value='Siguiente' class=bt ></TD>"
response.write  "</FORM>"
end if

response.write  "</tr></table>"

'response.write  ofv.botoneraESSX("c",rs,cantregmover,cantregmostrar,"no","no",wposi,addqry,"","no","","","")

 call ofv.cerrarconsulta(rs)
 call ofv.cerrarconn(cn)

response.write  "</body></html>"
end if



%>