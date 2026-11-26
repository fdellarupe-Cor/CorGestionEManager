<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  dim wstipo(6)

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  bVersion = request.querystring("bVersion")
  baccion = request.querystring("baccion")
  nID = request.querystring("nID")



  wntitulo=" onload=" & chr(34) & "document.title='COR&nbsp;E-Manager&nbsp;Consulta&nbsp;de&nbsp;Procedimientos';"  & chr(34) & " "

  sHTML = "<html><body>" & corstyler2()

  shtml = shtml & CorPnlFnc("N","") & "<br><hr color='#000000' size=2 width=100% ><br>"


  set cn = ofv.conectar(ofv.strconn2)





  bColor = "False"
  if nID <> "no" then

    redim wsmodo(0)
    redim wsmodoB(0)

    ym = 1

    strsql = "select * from modo "
    strsql = strsql & " Order By tipomodo"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       do until rs2.eof
          tm = ym
          redim preserve wsmodo(ym)
          redim preserve wsmodoB(ym)
          wsmodo(ym) = rs2(0)
          wstipo(0) = rs2(2)
          for ymm = 1 to 6
              wstipo(ymm) = rs2(ymm + 3)
          next
          wsmodoB(ym) = wstipo
          ym = ym + 1
       rs2.movenext
       loop
    else
          tm = ym
          redim preserve wsmodo(ym)
          redim preserve wsmodoB(ym)
          wsmodo(ym) = "NO"
    end if





    strsql = "select * from objaccion "
    strsql = strsql & "WHERE codproced = '" &  nID & "' and codaccion = '" & baccion & "' "
    strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    strsql = strsql & " Order By secuen"

    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    posicion = 0

    do while not rs2.eof
      nAccion = rs2(0)
      nAccionID = rs2(18)
      sImage = "&nbsp;"
      if bVersion = "SI" and rs2(14) then
      if rs2(15) = "M" then
        sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
      elseif rs2(15) = "A" then
        sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
      end if
      end if


      Celda = ofv.GenCelda("","dt","","20","","","",sImage,"si")

      tenc = "NO"
      for ym=1 to tm
          if rs2(5) = wsmodo(ym) then
             tenc = "SI"
             for ymm=0 to 6
                 wstipo(ymm) = wsmodoB(ym)(ymm)
             next

             exit for
          elseif rs2(5) < wsmodo(ym) then
             tenc = "NO"
             exit for
          end if
      next

      if ym > tm then  tenc = "NO"

      if   tenc = "NO" then
             for ymm=0 to 6
                 wstipo(ymm) = 0
             next
      end if

        if rs2(4) <> "" then
          Strsql = "Select descri,modo From Tipoaccalt Where tipoaccionalt = " & rs2(4)
          set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsAcciones.eof then
             sAccion = rsAcciones(0)
             sModo   = rsAcciones(1)
          END IF
          call ofv.cerrarconsulta(rsAcciones)
        end if


      strsql12 = "SELECT archivo,archivocur from modo "
      strsql12 = strsql12 & "Where tipomodo = '" & sModo & "' "
      set rs12 = ofv.crearconsultaEx(StrSql12,cn,1,volver)
      if not rs12.eof then
         sArch   = rs12(0)
         sArch2   = rs12(1)
      end if
      call ofv.cerrarconsulta(rs12)


      sLink = "&nbsp;" ' ofv.GenImage("","","","","Iconos/" & sArch2,"1","")

      if wstipo(1) then
        'Acciones alternativas
        CantRegAMover = "A"
        vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)
        sID = "Accion=" & nAccionID & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
        sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
        sID = sID & "_Tipobjeto=" & sTipo & "_proced=" & nID & "_Menu=" & sMenux & "_Sector="
        sID = sID & ofv.convertircar(sSectorx,"_","@") & "_origen=" & origen

        sImage ="&nbsp;" 'ofv.GenImage("","","","","Iconos/" & sArch,"0","")
      end if
      Celda = Celda & ofv.GenCelda("","ut","","20","","","",sLink,"si")

      'Secuencia
      Celda = Celda & ofv.GenCelda("","ut","","20","","","","" & rs2(10) & "","si")


      'Acciones

      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccion,i-1)
        sAcciont = mid(sAccion,i+1)
      else
      sAcciont=""
      end if
      sInput = "<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont & ""
      Celda = Celda & ofv.GenCelda("","vt colspan=4","","","","","",sInput,"si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

      'Texto de Referencia
      sArma = rs2(11)
      if len(sArma) > 0 then
        sArma = replace(sArma,chr(13),"<br>")
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        SArma = "" & SArma & ""
        Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if
      sArma = rs2(16)
      if len(sArma) > 0 then
        sArma = replace(sArma,chr(13),"<br>")
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        SArma = "" & SArma & ""
        Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if
      sArma = rs2(17)
      if len(sArma) > 0 then
        sArma = replace(sArma,chr(13),"<br>")
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        SArma = "" & SArma & ""
        Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if

      'link
      if wstipo(3) then
        if clng(rs2(19)) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & rs2(19)
           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsx.eof then
              wtipoL = rsx(6)
              wtipoLc = rsx(7)
              wUrl = rsx(8)
              wUrlDesc = rsx(9)
              wtipoLNK = rsx(10)
           else
              wtipoL = 0
              wtipoLc = "NO"
              wUrl = "NO"
              wUrlDesc = "NO"
              wtipoLNK = "NO"
           end if
           call ofv.cerrarconsulta(rsx)

           if wtipoLNK = "I" then

              wNomL = wUrlDesc
              if not isnull(wNomL) then
                 wNomL = UCASE(wNomL)
              else
                 wNomL = "S/D"
              end if

              if not isnull(wUrl) then
                 sHref = instr(wUrl,":\\")
                 if sHref > 0 then
                    sHref = wUrl
                 else
                    sHref = "HTTP:\\" & wUrl
                 end if
              else
                 sHref = "HTTP:\\"
              end if
              sFuncion = " target=_blank "


           else

              wNomL    = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomL)
              else

                 wNomL = "S/D"

                 Strsql = "Select * From csGenObjetos where codobjeto = " & wtipoL
                 set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if not rsx.eof then
                     wNomL = rsx(1)
                 else
                     wNomL = "S/D"
                 end if
                 call ofv.cerrarconsulta(rsx)


              end if

              sActcon = "Consultar.asp?Codigo=" & wtipoL & "&Tipo=" & wtipoLc
              sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
              sHref = "&nbsp;"
             ' response.write sActcon
              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

           end if



            sImagen = wNomL

            SArma = wNomL

            sLink       = ofv.GenLink("","btbn",sHref,"","",sFuncion,sgllnk,"si")
            sLink2       = ofv.GenLink("","nt",sHref,"","",sFuncion,SArma,"si")

            Celdak = ""
            Celdak = Celdak & ofv.GenCelda("","btbn","","2","","","",sLink,"si")
            Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","300","","0","0","0","-1","","",Filak,"si")

        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt colspan=4","","","","","",SArma ,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


        end if
      end if

        if rs2(6) <> "" and rs2(6) > "0"  then
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        sLink = "&nbsp;"

        nCopias = "&nbsp;"
'''       if rs2(5) = "DC" then
      if wstipo(2) then
          'Copias
          if not isnull(rs2(13)) then
            if rs2(13) > "0" then
               if rs2(13) > "1" then
                  nCopias = "<b>&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copias</b>"
               else
                  nCopias = "<b>&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copia</b>"
               end if
            end if
          end if
        end if

'   'Componentes
'          Strsql = "Select nombre From csacomple Where codcomp = " & rs2(6)
'          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'          if  not rsComponentes.eof then
'              sComponente = rsComponentes(0)
'              SArma       = replace(sComponente," ","&nbsp;")  & nCopias & ""
'              sHref       = "EditarImagen2.asp?ID=" & rs2(6)'''
'
'              sLink       = ofv.GenLink("","btba",sHref,"_blank","","",sglcmp,"si")
'              sLink2      = ofv.GenLink("","nt",sHref,"_blank","","",SArma,"si")

'          end if
'          call ofv.cerrarconsulta(rsComponentes)


     'Componentes
          sComponente = "&nbsp;"
          sLink       = sglcmp
          sLink2      = "&nbsp;"


          Strsql = "Select * From objlnk Where ltipo = 'KL' and objid = 0" & rs2(6)
           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsx.eof then

              wtipoL   = rsx(6)
              wtipoLc  = rsx(7) & ""

              wNV1    = rsX(11) & ""
              wNV2    = rsX(12) & ""
              wNV3    = rsX(13) & ""

              IF  WNV1 = "" THEN WNV1 = "0"
              IF  WNV2 = "" THEN WNV2 = "0"
              IF  WNV3 = "" THEN WNV3 = "0"


          Strsql = "Select nombre From csacomple Where codcomp = " & wtipoL
          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsComponentes.eof then
              sComponente = rsComponentes(0)
              SArma       = replace(sComponente," ","&nbsp;")  & nCopias & ""
              sHref       = "EditarImagen2.asp?ID=" & wtipoL
              sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

              sLink       = ofv.GenLink("","btba",GenKeyLnks(sHref),"_blank","","",sglcmp,"si")
              sLink2      = ofv.GenLink("","nt",GenKeyLnks(sHref),"_blank","","",SArma,"si")

          end if
          call ofv.cerrarconsulta(rsComponentes)

          end if
          call ofv.cerrarconsulta(rsx)




            Celdak = ""
            Celdak = Celdak & ofv.GenCelda("","btba","","2","","","",sLink,"si")
            Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","300","","0","0","0","-1","","",Filak,"si")


        Celda = Celda & ofv.GenCelda("","dt colspan=4","","","","","",SArma ,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if


      if rs2(8) > "0" or rs2(9) > "0" then

        'Sector Continuador
        sContinuador = "&nbsp;"
        if rs2(8) > "0" then
          Strsql = "Select descri From Eco Where eco = " & rs2(8)
          set rsContinuadores = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsContinuadores.eof then sContinuador = "Continuador: " & rsContinuadores(0)
          call ofv.cerrarconsulta(rsContinuadores)

          Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")

          Celdak = ""
          Celdak = Celdak & ofv.GenCelda("","btbs","","2","","","",sglcnt,"si")
          Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sContinuador,"si")
          Filak = ofv.GenRow("","","","","","","",Celdak,"si")
          SArma = ofv.GenTabla("","TWINE","","300","","0","0","0","-1","","",Filak,"si")


          Celda = Celda & ofv.GenCelda("","dt colspan=4","","","","","",SArma,"si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

        end if


        'Perfil Continuador
        sPerfil = "&nbsp;"
        if rs2(9) > "0" then
          Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs2(9) & "'"
          set rsPerfiles = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsPerfiles.eof then sPerfil = "Perfil: " & rsPerfiles(0)
          call ofv.cerrarconsulta(rsPerfiles)

          Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")

          Celdak = ""
          Celdak = Celdak & ofv.GenCelda("","btbn","","2","","","",sglprf,"si")
          Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sPerfil,"si")
          Filak = ofv.GenRow("","","","","","","",Celdak,"si")
          SArma = ofv.GenTabla("","TWINE","","300","","0","0","0","-1","","",Filak,"si")


          Celda = Celda & ofv.GenCelda("","dt colspan=4","","","","","",SArma,"si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

        end if

      end if

        Celda = ""
        for xx=0 to 6
            Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
        next
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

      rs2.movenext
      posicion = posicion + 1
    loop
  end if
  sHTML = shtml & "<center>" & ofv.GenTabla("","","","90% style='border-style:solid;border-width:1;border-color:#dddddd;' ","","0","0","0","0","","",Fila,"si") & "</center>"

  if nID <> "no" then call ofv.cerrarconsulta(rs2)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>