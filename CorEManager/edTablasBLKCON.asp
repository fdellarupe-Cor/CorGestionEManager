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
  stabla  = session("ESQTABLA")


  sSubTit  = request.querystring("SubTit")

'' response.write " stabla  " & stabla
'' response.write " sSubTit " & sSubTit

  if  sSubTit <> "" then
      session("SubTit") = sSubTit
  else
      sSubTit = session("SubTit")
  end if


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

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111 >"

'-------------------------------------------------

  set cn  = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    nestado = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

''response.write sTabla & " sTabla "

  Strsql  = "Select * From ObjTabdef Where objid=" & sTabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if not rs.eof then
     sForm = "edTablas" & rs(0)
     sid   = rs(0)


      Celda = ofv.GenCelda("","tt","","1%","","","","Tabla:&nbsp;","si")

      sInput = rs(5) & "."
      Celda = Celda & ofv.GenCelda("","ttt","","1%","","","",sInput,"si")

      sInput = rs(4)
      Celda = Celda & ofv.GenCelda("","ttt","","48%","","","",sInput,"si")

      sInput = "Bloque:"
      Celda = Celda & ofv.GenCelda("","tt","","1%","","","",sInput,"si")


      Strsql = "SELECT * From Objtabblk Where objid = " & sSubTit
      set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if  not rs2.eof then
          nid    = rs2(0)
          sInput = rs2(4) & "." & rs2(5)
          Celda  = Celda & ofv.GenCelda("","ttt","","98%","","","",sInput,"si")

      end if
      call ofv.cerrarconsulta(rs2)

      Fila = ofv.GenRow("","","","","25","","",Celda,"si")

  else
     sid = "NO"
  end if

  sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80%","","0","","0","0","","",fila,"si")


  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & ""

      Fila = ""

  if sid <> "NO" then
'''''''''''''''''''''''''''''''''''''''''''''''

  wcolspan  = "N"
  wrowspan  = "N"
  wcolspanc = 0
  wrowspanc = 0
  wcolspann = 0
  wcolspanw = "N"
  cstil = " style='background-color:#fafafa;border-right-style:solid;border-right-width:1;border-right-color:#cccccc;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "

  redim aColumnas(0)
  redim aColumnascol(0)
  redim aColumnasrow(0)


  if nID <> "no" then
    strsql = "Select tabperfil From Objtabdef Where objid = " & sID
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
    strsql = "Select * From Tabperfil Where Tabperfil = '" & rs2(0) & "'"
    call ofv.cerrarconsulta(rs2)
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    if not rs2.eof then
      nCols = rs2(7)
      if cint(nCols) > 10 then nCols = 10
      redim preserve aColumnas    (cint(nCols) - 1)
      redim preserve aColumnascol (cint(nCols) - 1)
      redim preserve aColumnasrow (cint(nCols) - 1)

      for ix = 0 to UBound(aColumnas)

          aColumnas (ix)    = rs2(ix + 16)
          aColumnascol (ix) = 0
          aColumnasrow (ix) = 0

      next
    end if



    end if
    call ofv.cerrarconsulta(rs2)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Encabezados titulos

    'simage1    = "<font face='corflaticon' size=3 color='#f7aa22'>&#052;</font>"
	simage1    = "<i class='las la-angle-right' ></i>"
    Celda      = ofv.GenCelda("","ttr","","","","","","&nbsp;","si")
    sInput     = "#"
    Celda      = Celda & ofv.GenCelda("","ttr colspan=2 ","","","","","",sInput,"si")
    sInput     = "Contenidos&nbsp;de&nbsp;Columnas"
    Celda      = Celda & ofv.GenCelda("","ttr colspan=" & UBound(aColumnas) + 2,"","","","","",sInput,"si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")


    if sOpcion = "T" then

       strsqlx = "Select * From Objtabcon Where tabblk = " & nID
       strsqlx = strsqlx &" and tabcontipo = '1' "
       strsqlx = strsqlx & " Order By tabconsec desc"
       set rsx = ofv.crearconsultaEx(StrSqlx,cn,1,volver)
       if not rsx.eof then
          nSecuencia = cint(rsx("tabconsec")) + 10
       else
          nSecuencia = 10
       end if
       call ofv.cerrarconsulta(rsx)
     Strsqlx = "insert into Objtabcon (codobjeto,tipobjeto,Tabblk,tabconsec,tabcontipo,TABCONID) values ("
     Strsqlx = Strsqlx & nCodigo & ",'" & sTipo & "'," & nID & "," & nSecuencia & ",1,'TB" & GTkey(nCodigo) & "')"
     call ofv.crearconsultaEx(StrSqlx,cn,1,vueltac)
    end if

    strsql = "SELECT Objtabcon.*, b.codinx2 From Objtabcon "

    'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objtabcon.objid = b.codinx2) "
   ' strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where   codsol = " & session("codSol") & ") b on (Objtabcon.objid = b.codinx2) "
    strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast(Objtabcon.objid as nvarchar)= b.codinx2) "
    strsql = strsql & " Where tabblk = " & nID & " and tabcontipo = '1' "
    strsql = strsql & " order by tabconsec "
 ''   response.write strsql & "<br>" & nid & "<br>"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    posicion = 1
    do until rs2.eof
      nAccion = rs2(0)


      for ix = 0 to UBound(aColumnas)

          aColumnascol (ix) = 0
          aColumnasrow (ix) = 0

      next


''''''''''''''''''''''''''
    strsql = "SELECT * From Objtabcmb Where TABobj = '" & rs2(26) & "'  "
    strsql = strsql & " order by tabcol "
 ''   response.write strsql & "<br>" & nid & "<br>"
    set rs21 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    do until rs21.eof

       if  rs21(8) <> "N" then aColumnasrow ((clng(rs21(6)) - 1)) = rs21(8)

       if  rs21(8) <> "N" then
           if  wcolspanw  = "N" then
               if  wrowspan  = "N" then
                   wrowspanc = rs21(8)
                   wcolspanc = rs21(7)
                   wcolspann = clng(rs21(6))
                   wcolspanw = "R"

                   if wcolspanc = "N" then wcolspanc = 1
                   if wrowspanc = "N" then wrowspanc = 1


               end if
           end if
       end if

       if  rs21(7) <> "N" then aColumnascol ((clng(rs21(6)) - 1)) = rs21(7)
       if  rs21(7) <> "N" then
           if  wcolspanw  = "N" then
               if  wcolspan  = "N" then
                   wcolspanc = rs21(7)
                   wcolspann = clng(rs21(6))
                   wcolspanw = "S"

                   if wcolspanc = "N" then wcolspanc = 1

               end if
           end if

       end if





       rs21.movenext
    loop
''''''''''''''''''''''''''




      Celda = ofv.GenCelda("","ut colspan=2 bgcolor='#eeeeee' ","","","","","","&nbsp;","si")



      'Orden y Tipo
      Celda = Celda & ofv.GenCelda("","uts bgcolor='#eeeeee'","right","","","","",simage1 & "&nbsp;","si")


''''''''''''''''''''''''''''''''''

      wcsc = 0


      'Columnas
      for ix = 0 to UBound(aColumnas)
        sSize = 10

        wxrs = ""
        wxcs = ""

        if  wcolspanw = "R" then
            if  wrowspan  = "N" and ix = (wcolspann-1) then
                wrowspan  = "S"
                wcolspan  = "S"

                wxrs = wrowspanc
                wxcs = wcolspanc

                if wxrs = 0 then wxrs = 1
                if wxcs = 0 then wxcs = 1

                wxrs = " ROWSPAN=" & wxrs & " "
                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101

                for ix2 = ix to UBound(aColumnas)

                    if ix2 > ix + (wcolspanc - 1) then
                       exit for
                    end if

                    if cint(aColumnas(ix2)) <> 0 then

                        sSize = sSize + cint(aColumnas(ix2))

                    end if


                next

            elseif  wrowspan  = "S" and ix = (wcolspann-1) then

                wxrs = ""

                if wcolspanc > 1 then
                   wcsc = wcolspanc
                else
                   wcsc = 1
                end if

            end if




        end if

        if  wcolspanw = "S"  then
            if  wcolspan  = "N" and ix = (wcolspann-1) then
                wcolspan  = "S"


                wxrs = ""
                wxcs = wcolspanc

                if wxcs = 0 then wxcs = 1

                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101

                for ix2 = ix to UBound(aColumnas)

                    if ix2 > ix + (wcolspanc - 1) then
                       exit for
                    end if

                    if cint(aColumnas(ix2)) <> 0 then

                        sSize = sSize + cint(aColumnas(ix2))

                    end if


                next



            end if
        end if

        if  wcsc = 0 or wcsc > 100 then

            sInput = rs2(ix + 6)

            Celda = Celda & ofv.GenCelda("","vt "  & cstil & wxrs & " " & wxcs ,"","","","","",sinput ,"si")


            if  wcsc > 100 then
                if wcolspanc > 1 then
                   wcsc = wcolspanc - 1
                else
                   wcsc = 0
                end if
            end if

        else

            wcsc = wcsc - 1

            if  wcsc = 0 then
                wcolspan  = "N"
            end if

            if  wcolspanw = "S" and wcolspan  = "N" then

                wcolspanw = "N"

                for ix2 = ix + 1 to UBound(aColumnas)

                    if aColumnascol (ix2) <> 0 then
                       wcolspanc = aColumnascol (ix2)
                       wcolspann = ix2 + 1
                       wcolspanw = "S"
                       exit for
                    end if

                next

            end if

        end if


      next


'''''''''''''''''''''''''''''''''

      Fila = ofv.GenRow("","","","","","","",Celda,"si")
      Formulario = Formulario & Fila

      Fila = ""

      if  wrowspan  <> "N" then

          wrowspanc = wrowspanc - 1

          if  wrowspanc = 0 then
              wrowspan  = "N"
              wcolspann = 0
              wcolspanw = "N"
              wcolspanc = 0
          end if


      end if

      rs2.movenext
      posicion = posicion + 1
    loop



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Encabezados Contenido
  '''  Formulario = ofv.Encabezados("Orden","Contenidos&nbsp;de&nbsp;Columnas")

  wcolspan  = "N"
  wrowspan  = "N"
  wcolspanc = 0
  wrowspanc = 0
  wcolspann = 0
  wcolspanw = "N"

  cstil = " style='border-right-style:solid;border-right-width:1;border-right-color:#cccccc;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "


    esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:#f7eedd;} "
    if sOpcion = "C" then

       strsqlx = "Select * From Objtabcon Where tabblk = " & nID
       strsqlx = strsqlx &" and tabcontipo = '2' "
       strsqlx = strsqlx & " Order By tabconsec desc"
       set rsx = ofv.crearconsultaEx(StrSqlx,cn,1,volver)
       if not rsx.eof then
          nSecuencia = cint(rsx("tabconsec")) + 10
       else
          nSecuencia = 10
       end if
       call ofv.cerrarconsulta(rsx)
     Strsqlx = "insert into Objtabcon (codobjeto,tipobjeto,Tabblk,tabconsec,tabcontipo,TABCONID) values ("
     Strsqlx = Strsqlx & nCodigo & ",'" & sTipo & "'," & nID & "," & nSecuencia & ",2,'TB" & GTkey(nCodigo) & "')"
     call ofv.crearconsultaEx(StrSqlx,cn,1,vueltac)
    end if


    strsql = "SELECT Objtabcon.*, b.codinx2 From Objtabcon "

    'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objtabcon.objid = b.codinx2) "
    strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & ") b on (Objtabcon.objid = b.codinx2) "
    strsql = strsql & " Where tabblk = " & nID & " and tabcontipo = '2' "

    strsql = strsql & " order by tabconsec "

' response.write "<br>" & strsql & "<br>"

    set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '         rs5.move 1
    posicion = 1
    nitem = 0
    do until rs5.eof


      nAccion = rs5(0)
      sForm = "edContenido" & rs5(0)
      nitem = nitem + 1

      for ix = 0 to UBound(aColumnas)

          aColumnascol (ix) = 0
          aColumnasrow (ix) = 0

      next


''''''''''''''''''''''''''
    strsql = "SELECT * From Objtabcmb Where TABobj = '" & rs5(26) & "'  "
    strsql = strsql & " order by tabcol "
  ''''  response.write  "<br><br>" & strsql & "<br>" & nid & "<br>"
    set rs21 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    do until rs21.eof

       if  rs21(8) <> "N" then aColumnasrow ((clng(rs21(6)) - 1)) = rs21(8)

       if  rs21(8) <> "N" then
           if  wcolspanw  = "N" then
               if  wrowspan  = "N" then
                   wrowspanc = rs21(8)
                   wcolspanc = rs21(7)
                   wcolspann = clng(rs21(6))
                   wcolspanw = "R"

                   if wcolspanc = "N" then wcolspanc = 1
                   if wrowspanc = "N" then wrowspanc = 1


               end if
           end if
       end if

       if  rs21(7) <> "N" then aColumnascol ((clng(rs21(6)) - 1)) = rs21(7)
       if  rs21(7) <> "N" then
           if  wcolspanw  = "N" then
               if  wcolspan  = "N" then
                   wcolspanc = rs21(7)
                   wcolspann = clng(rs21(6))
                   wcolspanw = "S"

                   if wcolspanc = "N" then wcolspanc = 1

               end if
           end if

       end if





       rs21.movenext
    loop
''''''''''''''''''''''''''
xclas = "ut"


      'Eliminar
      sHref = ""
      if bModificar then
        sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
  ''    Celda = ofv.BotonEliminar("Contenido",sHref)
        Celda = ofv.BotonEliminar("Eliminacion",sHref)
        Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
      else
        Celda = ofv.GenCelda("","","","","","","","","si")
      end if

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

      if (Nestado=5  or Nestado=6) and (rs5(33) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion

         ''''''agregado para color del cuadrito de Observaciones version 4.6

               if wComentario = "" then
        	        sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
	        ''''''''sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

         	   else
           	        sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
        	   end if
      ''''''
              Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")
               Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs ,"si")

      end if


      Celda = Celda & Celdaobs

      xmainscriptD = ""
      xmainscriptX = ""


      'Orden y Tipo
      sInput = ofv.GenerarInput(rs5(4).name & "L" & nitem,rs5(4),"agregar",sForm,"3","dt","")
      sInput2 = ofv.GenerarInput(rs5(5).name & "L" & nitem,rs5(5),"hidden",sForm,"10","dt","")
      Celda = Celda & ofv.GenCelda("","aaa style='border-right-style:solid;border-right-width:1;border-right-color:#cccccc;' valign=top ","","2%","","","",sInput & sInput2,"si")


      xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs5(4).name & "X" & nitem ,"0","hidden",sForm & "X" & nitem,"1","dt","")
      xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs5(5).name & "X" & nitem,"0","hidden",sForm & "X" & nitem,"1","dt","")

      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs5(4).name & "X" & nitem & "').value = document.getElementById('" & rs5(4).name & "L" & nitem & "').value; "
      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs5(4).name & "X" & nitem & "').name = '" & rs5(4).name & "'; "

      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs5(5).name & "X" & nitem & "').value = document.getElementById('" & rs5(5).name & "L" & nitem & "').value; "
      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs5(5).name & "X" & nitem & "').name = '" & rs5(5).name & "'; "



'''''''''''''''''''''''''''''''
      wcsc = 0


      'Columnas
      for ix = 0 to UBound(aColumnas)

        pooldata = ""

        sSize = 10

        wxrs = ""
        wxcs = ""

        if  wcolspanw = "R" then
            if  wrowspan  = "N" and ix = (wcolspann-1) then
                wrowspan  = "S"
                wcolspan  = "S"

                wxrs = wrowspanc
                wxcs = wcolspanc

                if wxrs = 0 then wxrs = 1
                if wxcs = 0 then wxcs = 1

                wxrs = " ROWSPAN=" & wxrs & " "
                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101

                for ix2 = ix to UBound(aColumnas)

                    if ix2 > ix + (wcolspanc - 1) then
                       exit for
                    end if

                    if cint(aColumnas(ix2)) <> 0 then

                        sSize = sSize + cint(aColumnas(ix2))

                    end if


                next

                 '  ix = ix2 - 1

            elseif  wrowspan  = "S" and ix = (wcolspann-1) then

                wxrs = ""

                if wcolspanc > 1 then
                   wcsc = wcolspanc
                else
                   wcsc = 1
                end if

            end if




        end if

        if  wcolspanw = "S"  then
            if  wcolspan  = "N" and ix = (wcolspann-1) then
                wcolspan  = "S"


                wxrs = ""
                wxcs = wcolspanc

                if wxcs = 0 then wxcs = 1

                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101

                for ix2 = ix to UBound(aColumnas)



                    if ix2 > ix + (wcolspanc - 1) then

                       exit for

                    end if

                    if cint(aColumnas(ix2)) <> 0 then

                        sSize = sSize + cint(aColumnas(ix2))

                      '''  pooldata = pooldata & rs5(ix2 + 6) & "-" & wcolspann

                    end if


                next




            end if
        end if

        if  wcsc = 0 or wcsc > 100 then

            if  cint(aColumnas(ix)) <> 0 then sSize = cstr(100 * cdbl(aColumnas(ix)) / 100)

          '''  sInput2 = ofv.GenerarInput(rs2(ix + 6).name,rs2(ix + 6),"agregar",sForm,sSize,"dt","")


            t3b = "<font class=at><textarea class=dt "
            t3b = t3b & " name=" & rs5(ix + 6).name & "L" & nitem & "  "
            t3b = t3b & " id=" & rs5(ix + 6).name & "L" & nitem & "  "
            t3b = t3b & " cols=" & sSize & " rows=3 "
            t3b = t3b & " onkeypress=" & chr(34)
            t3b = t3b & " if (this.value.length > 600) {alert('El texto es demasiado"
            t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
            xdatos =  rs5(ix + 6)
          ''''  xdatos = pooldata
            t3b = t3b & xdatos & "</textarea></font>"



      xmainscriptD = xmainscriptD & " <textarea class=dt id=" & rs5(ix + 6).name & "X" & nitem & " name=" & rs5(ix + 6).name & "X" & nitem & " style='display:none;' >0</textarea> "

   '''   xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs5(ix + 6).name & "X" & nitem ,"0","hidden",sForm & "X" & nitem,"1","dt","")

      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs5(ix + 6).name & "X" & nitem & "').value = document.getElementById('" & rs5(ix + 6).name & "L" & nitem & "').value; "
      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs5(ix + 6).name & "X" & nitem & "').name = '" & rs5(ix + 6).name & "'; "



            if aColumnascol (ix) <> 0 or aColumnasrow (ix) <> 0 then
               sImagenP = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#060;</FONT></B>"
            else
               sImagenP = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=2>&#060;</FONT></B>"
            end if


            sinput = "<a class='cc' href='EdTablasBLKcmb.asp?Tabla=" & sTabla
            sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SubTit=" & sSubTit
            sinput = sinput  & "&tipo=2&fila=" & rs5(26) & "&celda=" & (ix + 1)  & "&est=" & wcolspanw & "'"
            sinput = sinput  & " title='Propiedades de columna " & (ix + 1) & "' >" & sImagenP & "</a>"


            Celda = Celda & ofv.GenCelda("","ut valign=top " & cstil & wxrs & " " & wxcs ,"","","","","",sinput & t3b,"si")


            if  wcsc > 100 then
                if wcolspanc > 1 then
                   wcsc = wcolspanc - 1
                else
                   wcsc = 0
                end if
            end if

        else

            wcsc = wcsc - 1

            if  wcsc = 0 then
                wcolspan  = "N"
            end if

            if  wcolspanw = "S" and wcolspan  = "N" then

                wcolspanw = "N"

                for ix2 = ix + 1 to UBound(aColumnas)

                    if aColumnascol (ix2) <> 0 then
                       wcolspanc = aColumnascol (ix2)
                       wcolspann = ix2 + 1
                       wcolspanw = "S"
                       exit for
                    end if

                next

            end if

        end if


      next


'''''''''''''''''''''''''''''''''



      'Columnas





      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

      end if



      if bModificar then


      xmainscriptX = xmainscriptX & " document.getElementById('" & sForm & "X" & nitem & "').submit(); "


'''nitem
      '''   sInput = ofv.GenerarInput("boton1","Ok","buttom",sForm,"5","bt","")
         sInput = "<input type='button' value='OK' class='bt' onclick=" & chr(34) & xmainscriptX & chr(34)  &  " > "
         Formulariox = ofv.GenForm(sForm & "X" & nitem,"",sAccion,"","post","",xmainscriptD & sInput,"si")
      else
         sInput = "&nbsp;"
      end if



      Celda = Celda & ofv.GenCelda("","aaa  ","left","60%","","","",Formulariox,"si")


      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")





      if  wrowspan  <> "N" then

          wrowspanc = wrowspanc - 1

          if  wrowspanc = 0 then
              wrowspan  = "N"
              wcolspann = 0
              wcolspanw = "N"
              wcolspanc = 0
          end if


      end if



      rs5.movenext
      posicion = posicion + 1

    loop


    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","80%","","0","","-2","2","","",Formulario & fila,"si")

       Agregar = ""

    if bModificar then
      'Boton Agregar
       saccion = mid(vueltac,4)
       saccion = ofv.convertircar(saccion,"_","&")
       saccion = saccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OPC=C"
       Agregar = ofv.BotonAgregar(sAccion)

    end if

  sAccion  = "EdTablasBLK.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sVolver  = ofv.BotonVolver(sAccion)

  'Boton test
  sAccion   = "CoTablasCons.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Test      = ofv.BotonTest(sAccion)


  xxxx = ofv.botoneraESSX("0","","","","","","","","",Test,Agregar,sVolver,"")



    call ofv.cerrarconsulta(rs5)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   end if

  end if
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBS()

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si") '''''aaa onresize

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & xxxx


  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>