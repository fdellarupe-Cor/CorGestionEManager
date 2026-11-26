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

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

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
     ' response.write strsql
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
'  cstil = " style='border-style:solid;border-width:1;border-color:#99ccee;background-color:#fafafa;' "

  'cstil = " style='border-right-style:solid;border-right-width:1;border-right-color:#cccccc;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' "

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
  ''  Formulario = ofv.Encabezados("Orden","Titulos&nbsp;de&nbsp;Columnas")

    Celda = ofv.GenCelda("","ttr","","","","","","&nbsp;","si")
    sInput = "#"
    Celda = Celda & ofv.GenCelda("","ttr colspan=2 ","","","","","",sInput,"si")
    sInput = "Titulos&nbsp;de&nbsp;Columnas"
    Celda = Celda & ofv.GenCelda("","ttr colspan=" & (UBound(aColumnas)) + 2,"","","","","",sInput,"si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")

'''''''''''((UBound(aColumnas)) * 2) + 3

  ''  sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")
  ''  Formulario = ""

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

 '   strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objtabcon.objid = b.codinx2) "
   ' strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (Objtabcon.objid = b.codinx2) "
    strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast (Objtabcon.objid as nvarchar) = b.codinx2) "
    strsql = strsql & " Where tabblk = " & nID & " and tabcontipo = '1' "


    strsql = strsql & " order by tabconsec "
 '  response.write strsql & "<br>" & nid & "<br>"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    posicion = 1
    nitem = 0
    do until rs2.eof
      nAccion = rs2(0)
      sForm = "edtitulo" & rs2(0)
      nitem = nitem + 1

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

xclas = "ut"

      'Eliminar
      sHref = ""
      if bModificar then
        sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
    ''   Celda = ofv.BotonEliminar("Titulo",sHref)
         Celda = ofv.BotonEliminar("Eliminacion",sHref)
         Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
      else
        Celda = ofv.GenCelda("","","","","","","","","si")
      end if

      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","","","","","","si")

     'response.write nCodigo & " nCodigo " & "<br>"
     'response.write sTipo   & " sTipo " & "<br>"

''''' agregado para color del cuadrito de Observaciones version 4.6
      wComentario = ""
      strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and "
            strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"      ' and "
            strsql = strsql & " order by codcircuito, codusr, objid"
      'strsql = "SELECT * from OBJOBSCMT AND codsol = " & session("codSol") ""'& " and "
	''  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"      ' and "
	  'strsql = strsql & " order by codcircuito, codusr, objid"
	'response.write strsql
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if not rs.eof then
	     wComentario = rs(13)
	  end if
      call ofv.cerrarconsulta(rs)

      if (Nestado=5  or Nestado=6) and (rs2(33) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
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
      	       Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")

      end if


      Celda = Celda & Celdaobs

  ''    'agregado para compatibilidad con IE <11'
  ''       if bModificar then
  ''         sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
  ''         sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  ''      end if

      xmainscriptD = ""
      xmainscriptX = ""

      'Orden y Tipo
      sInput = ofv.GenerarInput(rs2(4).name & "L" & nitem ,rs2(4),"agregar",sForm,"3","dt","")

      sInput2 = ofv.GenerarInput(rs2(5).name & "L" & nitem,rs2(5),"hidden",sForm,"10","dt","")

      Celda = Celda & ofv.GenCelda("","ut valign=top ","","2%","","","",sInput & sInput2,"si")

      xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs2(4).name & "X" & nitem ,"0","hidden",sForm & "X" & nitem,"1","dt","")
      xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs2(5).name & "X" & nitem,"0","hidden",sForm & "X" & nitem,"1","dt","")

      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(4).name & "X" & nitem & "').value = document.getElementById('" & rs2(4).name & "L" & nitem & "').value; "
      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(4).name & "X" & nitem & "').name = '" & rs2(4).name & "'; "

      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(5).name & "X" & nitem & "').value = document.getElementById('" & rs2(5).name & "L" & nitem & "').value; "
      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(5).name & "X" & nitem & "').name = '" & rs2(5).name & "'; "


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
            if  wcsc = 0 and cint(aColumnas(ix)) <> 0 then sSize = cstr(aColumnas(ix))

            'sInput2 = ofv.GenerarInput(rs2(ix + 6).name & "L" & nitem,rs2(ix + 6),"agregar",sForm,sSize,"dt","")
			sInput2 = ofv.GenerarInput(rs2(ix + 6).name & "L" & nitem,rs2(ix + 6),"agregar",sForm,"150","dt","")
      ''      sInput2 = ofv.GenForm(sForm & "s3","",sAccion,"","post","",sInput,"si")


      xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs2(ix + 6).name & "X" & nitem ,"0","hidden",sForm & "X" & nitem,"1","dt","")


      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(ix + 6).name & "X" & nitem & "').value = document.getElementById('" & rs2(ix + 6).name & "L" & nitem & "').value; "
      xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(ix + 6).name & "X" & nitem & "').name = '" & rs2(ix + 6).name & "'; "




            if aColumnascol (ix) <> 0 or aColumnasrow (ix) <> 0 then
               'sImagenP = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#060;</FONT></B>"
			   sImagenP = "<i class='las la-sliders-h' style='font-size:1.5rem;color:" & Session("CMFPNBOT") & ";'></i>"
            else
               'sImagenP = "<B><FONT FACE=corflaticon COLOR=#6699cc SIZE=2>&#060;</FONT></B>"
			   sImagenP = "<i class='las la-sliders-h' style='font-size:1.5rem;color:" & Session("CMFPNBOT") & ";'></i>"
            end if


            sinput = "<a class='nt' href='EdTablasBLKcmb.asp?Tabla=" & sTabla
            sinput = sinput  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SubTit=" & sSubTit
            sinput = sinput  & "&tipo=1&fila=" & rs2(26) & "&celda=" & (ix + 1)  & "&est=" & wcolspanw & "'"
            sinput = sinput  & " title='Propiedades de columna " & (ix + 1) & "' >" & sImagenP & "</a>"


            Celda = Celda & ofv.GenCelda("","ut style='display:flex;align-items:center;' " & cstil  & wxrs & " " & wxcs ,"","","","","",sInput2 & sinput ,"si")




       ''     Celda = Celda & ofv.GenCelda("","uts " & wxrs ,"center","1","","","",sinput,"si")

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

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

      end if


      if bModificar then

 '''     xmainscriptD = xmainscriptD & " " & ofv.GenerarInput(rs2(ix + 6).name & "X" & nitem ,"0","hidden",sForm & "X" & nitem,"1","dt","")


    ''  xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(ix + 6).name & "X" & nitem & "').value = document.getElementById('" & rs2(ix + 6).name & "L" & nitem & "').value; "
     ''' xmainscriptX = xmainscriptX & "" & " document.getElementById('" & rs2(ix + 6).name & "X" & nitem & "').name = '" & rs2(ix + 6).name & "'; "


      xmainscriptX = xmainscriptX & " document.getElementById('" & sForm & "X" & nitem & "').submit(); "


'''nitem
      '''   sInput = ofv.GenerarInput("boton1","Ok","buttom",sForm,"5","bt","")
         sInput = "<input type='button' value='OK' class='bt' onclick=" & chr(34) & xmainscriptX & chr(34)  &  " > "
         Formulariox = ofv.GenForm(sForm & "X" & nitem,"",sAccion,"","post","",xmainscriptD & sInput,"si")
      else
         sInput = "&nbsp;"
      end if





      if  wrowspan  <> "N" then

          wrowspanc = wrowspanc - 1

          if  wrowspanc = 0 then
              wrowspan  = "N"
              wcolspann = 0
              wcolspanw = "N"
              wcolspanc = 0
          end if


      end if

      Celda = Celda & ofv.GenCelda("","ut","left","60%","","","",Formulariox,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

      rs2.movenext
      posicion = posicion + 1
    loop
    tabla= ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","80%","","0","","-2","2","","",Formulario & fila,"si")
    sHTML = sHTML & tabla
    'sHTML = sHTML & ofv.GenTabla("","aaa ","","50%","","0","","-2","2","","",Formulario,"si")

      Agregar = ""

    if bModificar then
      'Boton Agregar
       saccion = mid(vueltac,4)
       saccion = ofv.convertircar(saccion,"_","&")
       saccion = saccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OPC=T"
       Agregar = ofv.BotonAgregar(sAccion)

    end if

  sAccion  = "EdTablasBLK.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sVolver  = ofv.BotonVolver(sAccion)

  'Boton test
  sAccion   = "CoTablasCons.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  Test      = ofv.BotonTest(sAccion)

  xxxx =  ofv.botoneraESSX("0","","","","","","","","",Test,Agregar,sVolver,"")


    call ofv.cerrarconsulta(rs2)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


   end if

  end if
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBS()

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")''''aaa onresize

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & xxxx

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>