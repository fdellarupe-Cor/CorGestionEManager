
<%


 FUNCTION GENFMTPROCEDIMIENTOS(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""

  dim wstipo(6)


  Fila = ""



  bColor = "False"
  if nID <> "no" then

''response.write nid

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





    strsql = "select * from OBJACCION "
    strsql = strsql & "WHERE codproced = '" &  nID & "'"
    if  sCambios = "NO" then
        strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    else
        strsql = strsql & " and modi <> 0  "
    end if
    strsql = strsql & " Order By secuen"
''''response.write strsql
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    posicion = 0

    do while not rs2.eof
      nAccion   = rs2(0)
      nAccionID = rs2(18)
      xfclas    = "dt"
      sImage    = "&nbsp;"
      if bVersion = "SI" and rs2(14) and SESSION("SHWMOD") <> "N" then
      if rs2(15) = "M" then
         xftitl = "Modificado."
         xfclas = "btbs"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif rs2(15) = "A" then
         xftitl = "Agregado."
         xfclas = "btba"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >6</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif rs2(15) = "E" then
         xftitl = "Eliminado."
         xfclas = "btbn"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >5</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      end if
      end if

      sPoint = "T" & nAccion

      'Acciones
      if rs2(4) <> "" then
         Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & rs2(4)
         set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rsAcciones.eof then
             sAccionM = rsAcciones(0)
             sModo   = rsAcciones(1)
         end if
         call ofv.cerrarconsulta(rsAcciones)
      end if


      Celda  = ofv.GenCelda(sPoint,"uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")


      ' estado observacion

      Celdaobs  = ""


      if Nestado=5 and origen <> "ED" then

               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")


      end if


      Celda = Celda & Celdaobs




      tenc   = "NO"
      for ym=1 to tm
          if sModo = wsmodo(ym) then
             tenc = "SI"
             for ymm=0 to 6
                 wstipo(ymm) = wsmodoB(ym)(ymm)
             next

             exit for
          elseif sModo < wsmodo(ym) then
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






      sLink = "&nbsp;"

      if wstipo(1) then
        'Acciones alternativas

        strsql12 = "SELECT * from modo "
        strsql12 = strsql12 & "Where tipomodo = '" & sModo & "' "
        set rs12 = ofv.crearconsultaEx(StrSql12,cn,1,volver)
        if not rs12.eof then
           sArch   = rs12(2)

        end if
        call ofv.cerrarconsulta(rs12)

        nCantRegAMover = "A"
        vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)
        sID = "Accion=" & nAccionID & "&Codproced=" & nID & "&Codobjeto=" & nCodigo
        sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaC & "_Codobjeto=" & nCodigo
        sID = sID & "_Tipobjeto=" & sTipo & "_proced=" & nID & "_Menu=" & sMenux & "_Sector="
        sID = sID & ofv.convertircar(sSectorx,"_","@") & "_origen=" & origen

        sImage   = ofv.GenImage("","","","","Iconos/" & sArch,"0","")
        sLink    = sImage

        slinkalt = acciones(nAccionID,nCodigo,nID,sTipo)


      end if
      Celda = Celda & ofv.GenCelda("","dt","","1%","","","",sLink,"si")

      'Secuencia
 '     Celda = Celda & ofv.GenCelda("","dt","","4%","","","",rs2(10),"si")



      sVerbo   = sAccionM
      sAcciont = sAccionM
      i        = instr(sVerbo," ")

      if i > 0 then
        sVerbo   = left(sAccionM,i-1)
        sAcciont = mid(sAccionM,i+1)
      end if

      sInput = rs2(10) & ".&nbsp;<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont
      Celda  = Celda & ofv.GenCelda("","qt colspan=5  ","","","","","","<b>" & sInput & "</b>","si")
      Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


      sImagen  = "NO"
      wasotxt  = "N"


      'link
      if wstipo(3) then
         wlnkptr = "0" & rs2(19)
         if not isnull(wlnkptr) then
          if clng(wlnkptr) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & rs2(19)
           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rsx.eof then


                wasotxt     = rsx(14) & ""
                wasotxtstr  = CLNG("0" & rsx(15))
                wasotxtcan  = CLNG("0" & rsx(16))
                wasotxtdat  = rsx(17) & ""

                IF wasotxt = "" THEN wasotxt = "N"


              wtipoL   = rsx(6)
              wtipoLc  = rsx(7) & ""
              wUrl     = rsx(8) & ""
              wUrlDesc = rsx(9) & ""
              wtipoLNK = rsx(10) & ""

              wNV1    = rsX(11) & ""
              wNV2    = rsX(12) & ""
              wNV3    = rsX(13) & ""

              IF  WNV1 = "" THEN WNV1 = "0"
              IF  WNV2 = "" THEN WNV2 = "0"
              IF  WNV3 = "" THEN WNV3 = "0"


           else
              wtipoL   = 0
              wtipoLc  = "NO"
              wUrl     = "NO"
              wUrlDesc = ""
              wtipoLNK = "NO"

              WNV1 = "0"
              WNV2 = "0"
              WNV3 = "0"

           end if
           call ofv.cerrarconsulta(rsx)

           if wtipoLNK = "I" then

              wNomLt    = wUrlDesc
              wNomL     = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomLt)
              else
                 wNomL = "S/D"
              end if

              if not isnull(wUrl) then
                 sHref = instr(wUrl,"://")
                 if sHref > 0 then
                    sHref = wUrl
                 else
                    sHref = "HTTP://" & wUrl
                 end if
              else
                 sHref = "javascript:void(0);"
              end if
              sFuncion = " target=_blank "


           else

              wNomLt    = wUrlDesc
              wNomL     = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomLt)
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
              sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

              sHref = "&nbsp;"

              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

           end if


           slnktit = wNomLt

            SArma = wNomL

            Slnkdata = wNomL

            sImagen  = "SI"

            sLink       = ofv.GenLink("","btbn title='" & slnktit & "' ",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),sgllnk,"si")
            sLink2       = ofv.GenLink("","nt title='" & slnktit & "' ",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),"NNNNNN","si")



        end if
       end if
      end if




''''''''''''''''''''''''''''''''''''''


      'Texto de Referencia
      sArma = rs2(11)
      if len(sArma) > 0 then


        IF wasotxt = "S" THEN
           if  wasotxtstr > 0  and wasotxtcan > 0 then

               sfontstyl = " style='font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("APLFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:800;' "

               sArma = left(sArma,wasotxtstr - 1) & "" & replace(sLink2,"NNNNNN","&nbsp;<u><font " & sfontstyl & "  >" & mid(sArma,wasotxtstr,wasotxtcan) & "</font></u>&nbsp;") & "" & right(sArma,len(sArma) - (wasotxtstr + wasotxtcan - 1))
           end if
        end if

        sArma = replace(sArma,chr(13),"<br>")
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if

      sArma = rs2(16)
      if len(sArma) > 0 then
        sArma = replace(sArma,chr(13),"<br>")
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if

      sArma = rs2(17)
      if len(sArma) > 0 then
        sArma = replace(sArma,chr(13),"<br>")
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","utg colspan=4 style='text-align:justify;' ","","","","","",SArma,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if


      'link
      if wstipo(3) then

        if  wasotxt  = "N" then
          if  simagen <> "NO" then

            sLink2 = replace(sLink2,"NNNNNN",slnkdata)

            Celdak = ""
            Celdak = Celdak & ofv.GenCelda("","btbn title='ver documento.' ","","2","","","",sLink,"si")
            Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")


             Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
             Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;","si")
             Celda = Celda & ofv.GenCelda("","dt2 colspan=3","","","","","",SArma,"si")
             Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

          end if

        end if


      end if




        if rs2(6) <> "" and rs2(6) > "0"  then
           Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")

           nCopias = "&nbsp;"

           if wstipo(2) then
        'Copias
              if not isnull(rs2(13)) then
                 if rs2(13) > "0" then
                    if rs2(13) > "1" then
                       nCopias = "&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copias"
                    else
                       nCopias = "&nbsp;En&nbsp;" & rs2(13) & "&nbsp;Copia"
                    end if
                 end if
              end if
           end if

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
            Celdak = Celdak & ofv.GenCelda("","btba title='ver complemento.' ","","2","","","",sLink,"si")
            Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")


          Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","dt2 colspan=3","","","","","",SArma,"si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if


      if rs2(8) > "0" or rs2(9) > "0" then
        Celda = ofv.GenCelda("","dt colspan=3","","20","","","","&nbsp;","si")
        'Sector Continuador
        sContinuador = "&nbsp;"
        if rs2(8) > "0" then
          Strsql = "Select descri From Eco Where eco = " & rs2(8)
          set rsContinuadores = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsContinuadores.eof then sContinuador = rsContinuadores(0) & "&nbsp;"
          call ofv.cerrarconsulta(rsContinuadores)
        end if


        'Perfil Continuador
        sPerfil = "&nbsp;"
        if rs2(9) > "0" then
          Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs2(9) & "'"
          set rsPerfiles = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsPerfiles.eof then sPerfil = rsPerfiles(0) & "&nbsp;"
          call ofv.cerrarconsulta(rsPerfiles)
        end if


        Celdak = ""
        Celdak = Celdak & ofv.GenCelda("","ut title='sector continuador.' ","","2","","","",sglcnt,"si")
        Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sContinuador & "&nbsp;&nbsp;","si")
        Celdak = Celdak & ofv.GenCelda("","ut title='perfil continuador.'  ","","2","","","",sglprf,"si")
        Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sPerfil,"si")

        Filak = ofv.GenRow("","","","","","","",Celdak,"si")
        SArma = ofv.GenTabla("","TWINE","","98%","","0","0","0","-1","","",Filak,"si")

        Celda = Celda & ofv.GenCelda("","dt ","","","","","","&nbsp;&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt colspan=3","","","","","",SArma,"si")

        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if

      if  wstipo(1) then
          Celda =  ofv.GenCelda("","dt colspan=3","","1%","","","","&nbsp;","si")
          Celda =  celda & ofv.GenCelda("","dt colspan=4","","","","","",sLinkalt,"si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
          sLinkalt = ""
      end if


        Celda    = ""

        Celda = Celda & ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","3%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","50%","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","dt","","","","","","&nbsp;","si")

        Fila = Fila & ofv.GenRow("","","","","5","","",Celda,"si")

      rs2.movenext
      posicion = posicion + 1
    loop
  end if



  FNSHTML = FNSHTML & ofv.GenTabla("","","","95%","","0","0","0","0","","",Fila,"si")


  GENFMTPROCEDIMIENTOS = "<div id=zglos>" & FNSHTML & "</div>"

 END FUNCTION


function acciones(nAccion,nCodigo,nProced,sTipo)
  dim FNSHTML2
  dim rs,rs2
  set cn = ofv.conectar(ofv.strconn2)

  strsqlx = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)


  strsql = "SELECT TIPOACCALT.tipoaccionalt, TIPOACCALT.descri, Objaccion.secuen "
  strsql = strsql & "FROM OBJACCION, TIPOACCALT where "
  strsql = strsql & "OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
  strsql = strsql & "and OBJACCION.codaccion = '" & nAccion & "'"
  strsql = strsql & " Order By secuen"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
     VTACC = rs(0)
     VDACC = rs(1)
     VSACC = rs(2)
  end if
  call ofv.cerrarconsulta(rs)




  strsql = "SELECT * from csaccaltsel "
  strsql = strsql & "Where idaccion = '" & nAccion & "' "
  strsql = strsql & " and (modi = 0 or "
  strsql = strsql & " (modi <> 0 and tipmodi <> 'E')) "
  strsql = strsql & " order by secuencia"
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)



  strsql = "SELECT * FROM TIPOACCALT WHERE TIPOACCALT.modo='OP' Or TIPOACCALT.modo='RF' "
  strsql = strsql & "ORDER BY descri"


  FNSHTML2 = ""

  FNSHTML2 = FNSHTML2 & ofv.GenTabla("panel","aaa style='border-style:solid;border-width:1;border-color:#eeeeee;' ","","80% ","","0","0","-1","2","","","","no")

  t2    = ofv.GenCelda("","utis colspan=3 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ","","","","","",sfidkC & "&nbsp;NO.","si")

  FNSHTML2 = FNSHTML2 & ofv.GenRow("","","","","","","",t2,"si")


  Vcont = 1
  Valor = ""
  VCTACC = ""
  VCDACC = ""
  VCSACC = ""

  if not rs.eof then
  do until rs.eof
    Valor = rs(0)
    if not isnull(rs(8)) then
       Continuador = rs(8)


        strsql = "SELECT TIPOACCALT.tipoaccionalt, TIPOACCALT.descri, Objaccion.secuen "
        strsql = strsql & "FROM OBJACCION, TIPOACCALT where "
        strsql = strsql & "OBJACCION.tipoaccion = TIPOACCALT.tipoaccionalt "
        strsql = strsql & "and OBJACCION.codaccion = '" & Continuador & "'"
        strsql = strsql & " Order By secuen"
       'response.write strsql
       set rs2 = ofv.crearconsultaEx(StrSql,cn,1,volver)
       if not rs2.eof then
          VCTACC = rs2(0)
          VCDACC = rs2(1)
          VCSACC = rs2(2)
       end if
       call ofv.cerrarconsulta(rs2)

    end if
    ID = rs(2)
      sImage = "&nbsp;"
      if bVersion = "SI" and rs(9) and SESSION("SHWMOD") <> "N" then
      if rs(10) = "M" then
        sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
      elseif rs(10) = "A" then
        sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
      end if
      end if

    t1 = ofv.GenCelda("","dt","","5%","","","",sImage,"si")

    t1 = t1 & ofv.GenCelda("","dt","","5%","","","","&nbsp;" & VSACC & "." & rs(6) & "&nbsp;","si")

    Vcont = Vcont + 1

     'Texto Accion Alt
      if rs(4) <> "" then
        Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & rs(4)
        'response.write strsql
        set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsAcciones.eof then
           sAccion = rsAcciones(0)
        end if
        call ofv.cerrarconsulta(rsAcciones)
      end if

      sVerbo = sAccion
      sAcciont = sAccion
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(sAccion,i-1)
        sAcciont = mid(sAccion,i+1)
      end if
      sInput = "<B><i>" & sVerbo  & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont & "</b>"

    t2 = ofv.GenCelda("","qt ","","90%","","","",sInput,"si")

    FNSHTML2 = FNSHTML2 & ofv.GenRow("","","","","","","",t1 & t2,"si")

    'Texto de Referencia
    if ofv.explorador = "MSIE" then
      nCol = 60
    else
      nCol = 60
    end if

    t3 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t3 = t3 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sArma = rs(7)
    if len(sArma) > 0 then
       sArma = replace(sArma,chr(13),"<br>")
    end if
'''    SArma = "<font color=#225588 face=arial-narrow size=1>" & sArma & "</font>"
    t3 = t3 & ofv.GenCelda("","et ","","","","","",SArma,"si")
    FNSHTML2 = FNSHTML2 & ofv.GenRow("","","","","","","",t3,"si")
    rs.movenext
  loop


    'Encabezado Continuador



    t4 = ofv.GenCelda("","","","2%","","","","&nbsp;","si")
    t4 = t4 & ofv.GenCelda("","","","","","","","&nbsp;","si")
      'Texto Continuador
    if VCDACC <> "" then
      sVerbo = VCDACC
      sAcciont = VCDACC
      i = instr(sVerbo," ")
      if i > 0 then
        sVerbo = left(VCDACC,i-1)
        sAcciont = mid(VCDACC,i+1)
      end if
      sInput = "<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
      sInput = sInput & sAcciont
      sContinuador = "<b>" & VCSACC & " - " & sInput & "</b>"
    end if
    t4 = t4 & ofv.GenCelda("","et","","","","","","Continua en: " & sContinuador,"si")
    if VCDACC <> "" then
       FNSHTML2 = FNSHTML2 & ofv.GenRow("","","","","","","",t4,"si")
    end if

else

    t6 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t6 = t6 & ofv.GenCelda("","","","","","","","&nbsp;","si")
 ''   SArma = "<h5><B>Sin&nbsp;datos&nbsp;disponibles</B></h5>"
    SArma = "&nbsp;"
    t6 = t6 & ofv.GenCelda("","dt","","","","","",SArma,"si")
    FNSHTML2 = FNSHTML2 & ofv.GenRow("","","","","","","",t6,"si")


end if

    t5 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    FNSHTML2 = FNSHTML2 & ofv.GenRow("","","","","","","",t5,"si")


     FNSHTML2 = FNSHTML2 & ofv.GenFTabla()

  call ofv.cerrarconsulta(rs)



  acciones=FNSHTML2
end function

%>