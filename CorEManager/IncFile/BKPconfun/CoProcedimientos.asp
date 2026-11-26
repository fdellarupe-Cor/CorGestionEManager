<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTPROCEDIMIENTOS.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



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






  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"



  sfidkD  = sfidkc
  stexto1 = stextoact





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


  sHTML = sHTML &  GENFMTPROCEDIMIENTOS(nCodigo,STIPO)





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



%>