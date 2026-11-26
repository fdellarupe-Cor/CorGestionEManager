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

  nCodigo  = request.querystring("Codobjeto")
  sTipo    = request.querystring("Tipobjeto")
  sSectorx = request.querystring("Sector")
  sMenux   = request.querystring("Menu")
  origen   = request.querystring("origen")
  sperfil  = session("sperfil")
  sCambios = request.querystring("sCambios")
  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
  sNIV3   = request.querystring("NIV3")
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")

  INXMOD  = request.querystring("INXMOD")

  if  INXMOD = "" then INXMOD = "S"

  nCantRegAMover = request.querystring("CantRegAMover")

  if nCantRegAMover = "x" then
     xsproced = request.form("selcom")
  elseif nCantRegAMover = "A" then
     xsproced = request.querystring("proced")
  end if

'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK 
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "5" AND SNIV2 <> "0" THEN 

    xsproced =  SNIV2

    nCantRegAMover = "A"
    

 ELSE

    SPLFUNC = ""

 END IF 

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"


'''''''''''''''''''''''''''''''''''''''''''''''''


  if  sCambios = "" then sCambios = "NO"

  set cn = ofv.conectar(ofv.strconn2)
  '-------------------------------
  addperfil=""
  if smenux="PORSECTOR" THEN
     if sperfil="nsnc" or sperfil="" then 
      addperfil=""
     else
      set rspe = ofv.crearconsultaEx("select descri from tipoperfil where tipoperfil='" & sperfil & "'",cn,1,parametros)
       if not rspe.eof then  
        addperfil= " and  perfdescri='" & rspe(0) & "' "
        
        call ofv.cerrarconsulta(rspe)
       end if
     end if
  end if 
  '---------------------------------
  if len(origen) = 0 then origen = "CO"
  
  if sMenux = "PORSECTOR" then
    sSelecx = " and ECOdescri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' "
    sSelecx = sSelecx & addperfil
  else
    sSelecx = ""
  end if

  if nCantRegAMover = "x" then

    Strsql = "SELECT codproced,nombre,frecuencia  "
    Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & nCodigo
    Strsql = Strsql & " and codproced = '" & xsproced & "' " & sSelecx
    Strsql = Strsql & " order by secuen, secuencia,codproced "

  else
    if nCantRegAMover = "A" then

      Strsql = "SELECT codproced,nombre,frecuencia  "
      Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & nCodigo
      Strsql = Strsql & " and codproced = '" & xsproced & "' " & sSelecx
      Strsql = Strsql & " order by secuen, secuencia,codproced "

    else

      Strsql = "SELECT codproced,nombre,frecuencia  "
      Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & nCodigo
      Strsql = Strsql & sSelecx & " order by secuen, secuencia,codproced "
     CantRegAMover = "0" & request.querystring("CantRegAMover")
     CantRegAMover = cdbl(CantRegAMover)
     nCantRegAMover = CantRegAMover

    end if
  end if

  CantRegAMover2   = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar  = 1
  CantRegAMostrar2 = 4
  if ofv.explorador = "MSIE" then CantRegAMostrar2= 5

 sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  sMenu   = Session("menu")
  bModificar = session("Modificar")
  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_proced=" & xsproced

  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)

  

  strsqlx = "Select Version,estado From Objprty Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if  not rs.eof then
      if  rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1) 
  end if
  call ofv.cerrarconsulta(rs)

  'response.write StrSql & "4444"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      nID = rs(0)
     'response.write nid & "4444"

      bDescri = rs(1)

    Strsql = "SELECT *  "
    Strsql = Strsql & " from CSCEventoSectorPerfil Where codproced = '" & rs(0) & "' "
    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx.eof then


      EVSC = rsx(2)
      EVPS = rsx(6)
      EVDS = rsx(4)
      secd = rsx(3)
      perd = rsx(5)
 
     end if



   if origen <> "ED" then

    ' tomo todos los eventos

      stexto1 = "Seleccionar&nbsp;circuito"

      Celdat  = stexto1

      sAccion = "javascript:void(0);"
      sOnClik = " onclick=""document.all.TIND2.style.display='block';document.all.TIND1.style.display='none';"" "


      stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
      stexto1 = stexto1 & "  >" & Celdat & "</a>"

      Celdat  = "<td class=uts >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg style='cursor:pointer;' " & sOnClik & " >" & stexto1 & "</td>"
  ''    Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

      stexto1 = "Circuitos"

      Celdat  = "<td class=uts >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
    '  Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")

    eveant = ""

    Strsql = "SELECT *  "
    Strsql = Strsql & " from CSCEventoSectorPerfil Where obcodobjeto = " & nCodigo
    Strsql = Strsql & sSelecx 
    strsql = strsql & addperfil & " order by secuen, secuencia,codproced "
    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    do while not rsx.eof

       if eveant <> rsx(4) then

          Celdat  = "<td class=uts >" & sfidkA & "</td>"
          Celdat  = Celdat & "<td class=utg  >" & rsx(2) & ".&nbsp;" & rsx(4) & "</td>"
          Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

          eveant  = rsx(4)
       end if

          sAccion = "CoProcedimientos.asp?INXMOD=N&CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
          sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&proced=" & rsx(0)
          sOnClik1 = ""
          sOnClik2 = " onclick=""document.all.ASC" & rsx(0) & ".click();"" "

       if rsx(0)  = rs(0).value then
          sfidkD  = sfidkc
          stexto1 = "<u>" & rsx(6) & ".&nbsp;" & rsx(3) & "&nbsp;-&nbsp;" & rsx(5) & "</u>"

       else
          sfidkD  = sfidkB
          stexto1 = rsx(6) & ".&nbsp;" & rsx(3) & "&nbsp;-&nbsp;" & rsx(5)


       end if

          Celdat  = stexto1

          stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=nt href='" & sAccion & "' " & sOnClik1 & " "
          stexto1 = stexto1 & "   >" & Celdat & "</a>"

          Celdat  = "<td class=utis >" & sfidkD & "</td>"
          Celdat  = Celdat & "<td class=uti style='cursor:pointer;' " & sOnClik2 & " >" & stexto1 & "</td>"
          Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")



      rsx.movenext
   loop  



  else



  end if


    Strsql = "SELECT descri from frecuencia where codfrec = '" & rs(2) & "' "
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs1.eof then
       wsfrec = rs1(0)
    else
       wsfrec = "SIN&nbsp;DEFINIR"
    end if


    Strsql = "SELECT referencia from objproced where codproced = '" & rs(0) & "' "
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs1.eof then
       wsref = rs1(0)
    else
       wsref = "NO"
    end if

    if isnull(wsref) then wsref = "NO"


  else
    nID = "no"
  end if

  Fila = ""

'''''''''' arma tabla de indices '''''''''''''''''''''''''


   if origen <> "ED" then
      stexto1 = "Circuito&nbsp;Completo"

      Celdat  = stexto1

      sAccion = "CorVerTodoseventos.asp?volver=" & ofv.convertircar(vuelta,"_","&") & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "&Menu=" & sMenux & "&Sector=" & ofv.convertircar(sSectorx,"_","@")
      sAccion = sAccion & "&ncodigo=" & ncodigo
      saccion = saccion & "&proced=" & request.querystring("proced")

      stexto1 = "<A ID='AVTD' NAME='AVTD' class=utg href='" & sAccion & "' "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts >" & sfidk1 & "</td>"
      Celdat  = Celdat & "<td class=utg style='cursor:pointer;' onclick=""document.all.AVTD.click();"" >" & stexto1 & "</td>"
     ''' Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

   end if

  if nID <> "no" and origen <> "ED" then
      stexto1 = "Cursograma"

      Celdat  = stexto1

      sAccion = "Cursograma.asp?volver=" & vuelta & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo
      sAccion = sAccion & "_Menu=" & sMenux & "_Sector=" & ofv.convertircar(sSectorx,"_","@")
      sAccion = sAccion & "&Codproced=" & nID & "&Tipobjeto=" & sTipo

      stexto1 = "<A ID='ACRS' NAME='ACRS' class=utg href='" & sAccion & "' "
      stexto1 = stexto1 & " target='_blank'  >" & Celdat & "</a>"

      Celdat  = "<td class=uts >" & sfidk2 & "</td>"
      Celdat  = Celdat & "<td class=utg style='cursor:pointer;' onclick=""document.all.ACRS.click();"" >" & stexto1 & "</td>"
    '''  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


'    if bVersion = "SI" then'
'       if  sCambios = "NO" then
'           stexto1  = "Cambios"
'           xcambios = "SI"
'           sfidkD   = sfidk3
'       else
'           stexto1  = "Todos&nbsp;el&nbsp;Circuito"
'           xcambios = "NO"
'           sfidkD   = sfidk4
'       end if
'      sAccion = "CoProcedimientos.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
'      sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&proced=" & nid & "&scambios=" & xcambios
'      stexto1 = "<A ID='ACMB' NAME='ACMB' class=utg href='" & sAccion & "'   >" & stexto1 & "</a>"
'      Celdat  = "<td class=uts >" & sfidkD & "</td>"
'      Celdat  = Celdat & "<td class=utg style='cursor:pointer;' onclick=""document.all.ACMB.click();"" >" & stexto1 & "</td>"
'      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")
'   end if


  end if
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Celdat  = "<td class=uts colspan=2 style='border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</td>"
'  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

  sfidkD  = sfidkc
  stexto1 = stextoact

  Celdat  = "<td class=uts >" & sfidkD & "</td>"
  Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"
 ' Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")



' Detalle del procedimiento




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Celda = ofv.GenCelda(sPoint,"ttt","","10%","","","","EVENTO:","si")

    stexto1 = EVSC & "." & EVPS & "&nbsp;-&nbsp;" & EVDS

    if  sCambios = "SI" then stexto1 = stexto1 & "&nbsp;(Cambios)"

    Celda = Celda & ofv.GenCelda("","ttt colspan=3 ","","60%","","","",stexto1,"si")

    Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")


    Celda = ofv.GenCelda("","ut","","10%","","","","SECTOR:","si")

    Celda = Celda & ofv.GenCelda("","vti","","60%","","","",SECD,"si")



    Celda = Celda & ofv.GenCelda("","ut","","10%","","","","FRECUENCIA:","si")


    Celda = Celda & ofv.GenCelda("","vti","","20%","","","",wsfrec,"si")


    Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")




    Celda = ofv.GenCelda("","ut","","10%","","","","PERFIL&nbsp;:","si")


    Celda = Celda & ofv.GenCelda("","vti colspan=3 ","","60%","","","",perD,"si")


    Fila = FILA & ofv.GenRow("","","","","","","",celda,"si")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''






' Referencia del procedimiento
  if wsref <> "NO" then
     Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100% colspan=4","","","",replace(wsref,chr(13),"<br>"),"si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  end if


  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%   ","","0","","4","0","","",Fila,"si")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

''  response.write sHTML
''  sHTML = ""

  Fila = ""



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





    strsql = "select * from OBJACCION "
    strsql = strsql & "WHERE codproced = '" &  nID & "'"
    if  sCambios = "NO" then
        strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
    else
        strsql = strsql & " and modi <> 0  "
    end if 
    strsql = strsql & " Order By secuen"

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

      '	       sAccion = "genObs.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      '	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID
      '        xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      '	       sHref   = sAccion
     ' 	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
     ' 	       Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Observacion de " & BtipP & bpespec & "' ","","","1","","",sInput,"si")
     '          Celdaobs  = ofv.GenLink("",xclas & " title='Obs de " & BtipP & bpespec & "' ",sHref,"","","",Celdaobs,"si")



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
'      if rs2(5) = "CN" then
'      response.write "modo " & rs2(5) & " control " & wstipo(1) & "<br>"
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

'''''''''' arma tabla de indices '''''''''''''''''''''''''
      stexto1 = sInput
    '''  stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1

      Celdat  = stexto1

      stexto1 = "<A class=nt href='javascript:void(0);' "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=utis align=right >" & sfidkB & "</td>"
      Celdat  = Celdat & "<td class=utgi style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
    '  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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



  sHTML = sHTML & ofv.GenTabla("","","","95%","","0","0","0","0","","",Fila,"si")




  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"


''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

 ' call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''



  if nID <> "no" then call ofv.cerrarconsulta(rs2)



  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


  if CantRegAMover = "A" then

     if INXMOD = "S"  then

        sHTML = sHTML & "<script language='javascript' > clIndice(); </script> "

     END IF



  else 

     if INXMOD = "N"  then

        sHTML = sHTML & "<script language='javascript' > clIndice(); </script> "

     END IF


  'Botonera

      Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")

  end if




  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")

  'cambios
   Listar = ""

  if nID <> "no" and origen <> "ED" then

   if bVersion = "SI" and SESSION("SHWMOD") <> "N" then

      if  sCambios = "NO" then
          xcambios = "SI"
      else
          xcambios = "NO"
      end if

      xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
      sAccion = "CoProcedimientos.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
      sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&proced=" & nid & "&scambios=" & xcambios
      Listar  = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

   end if
  end if




  Agregar = Celt2
  Buscar  = Celt1
  Celda   = Listar

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

function acciones(nAccion,nCodigo,nProced,sTipo)
  dim shtml
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


  sHTML = ""

  sHTML = sHTML & ofv.GenTabla("panel","aaa style='border-style:solid;border-width:1;border-color:#eeeeee;' ","","80% ","","0","0","-1","2","","","","no")

  t2    = ofv.GenCelda("","utis colspan=3 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' ","","","","","",sfidkC & "&nbsp;NO.","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t2,"si")


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

    sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"si")

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
    sHTML = sHTML & ofv.GenRow("","","","","","","",t3,"si")
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
       sHTML = sHTML & ofv.GenRow("","","","","","","",t4,"si")
    end if

else

    t6 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t6 = t6 & ofv.GenCelda("","","","","","","","&nbsp;","si")
 ''   SArma = "<h5><B>Sin&nbsp;datos&nbsp;disponibles</B></h5>"
    SArma = "&nbsp;"
    t6 = t6 & ofv.GenCelda("","dt","","","","","",SArma,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t6,"si")


end if

    t5 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    t5 = t5 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t5,"si")


     sHTML = shtml & ofv.GenFTabla()

  call ofv.cerrarconsulta(rs)



  acciones=shtml
end function

%>