<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTCONTROLES.inc" -->
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

  sCambios = request.querystring("sCambios")


  if  sCambios = "" then sCambios = "NO"

  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
  sNIV3   = request.querystring("NIV3")
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")

  nCantRegAMover = request.querystring("CantRegAMover")


  if nCantRegAMover = "x" then
     xsproced = request.form("selcom")
  elseif nCantRegAMover = "A" then
     xsproced = request.querystring("control")
  end if

'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "7" AND SNIV2 <> "0" THEN

    xsproced =  SNIV2

    nCantRegAMover = "A"


 ELSE

    SPLFUNC = ""

 END IF

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"



  sMenux = request.querystring("Menu")
  if sMenux = "PORSECTOR" then
     sSelecx = " and descri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' "
  else
     sSelecx = ""
  end if
  if nCantRegAMover = "x" then
     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo
     Strsql = Strsql & " and objid = " & xsproced & sSelecx
  else
  if nCantRegAMover = "A" then
     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo
     Strsql = Strsql & " and objid = " & xsproced & sSelecx
  else
     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo & sSelecx
     CantRegAMover = "0" & request.querystring("CantRegAMover")
     CantRegAMover = cdbl(CantRegAMover)
     nCantRegAMover = CantRegAMover
  end if
  end if

  CantRegAMover2  = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  if ofv.explorador = "MSIE" then
    CantRegAMostrar2 = 7
  else
    CantRegAMostrar2 = 5
  end if

  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_control=" & xsproced
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)





  sMenu = Session("menu")
  bModificar = session("Modificar")

  scontdesc = "NO"

  sHTML = ofv.FormHeader(session("FormN"))

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  Fila = ""

''''''''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)

  strsqlx = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

'''  response.write strsql & "<br>" & chr(13)

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     nID       = rs(0)
     scontdesc = rs(4)
     bDescri   = rs(1)
     if len(scontdesc) = 0 then scontdesc = "NO"


''''''''''''''''''
      stexto1 = "Seleccionar&nbsp;Control"
      stexto1 = "&nbsp;&nbsp;" & stexto1

      Celdat  = stexto1

      sAccion = "javascript:void(0);"
      sOnClik = " onclick=""document.all.TIND2.style.display='block';document.all.TIND1.style.display='none';"" "


      stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   '   Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

      stexto1 = "Controles"

      Celdat  = "<td class=uts >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   '   Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")



''''''''''''''''''''''

     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo & sSelecx

     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do while not rsx.eof

       if clng(rsx(0))   = clng(rs(0).value) then
          sfidkD  = sfidkC
          stexto1 = rsx(1)
          stextoact = stexto1
       else
          sfidkD  = sfidkA
          stexto1 = rsx(1)
          sAccion  = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
          sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & rsx(0)
          sOnClik1 = ""
          sOnClik2 = " onclick=""document.all.ASC" & rsx(0) & ".click();"" "

       Celdat  = stexto1

       stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
       stexto1 = stexto1 & "   >" & Celdat & "</a>"

       Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
       Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"
       Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

       end if

      rsx.movenext
   loop

''''''''''''''''''''''''''''''''


     call ofv.cerrarconsulta(rs)

  else

    nID = "no"

  end if

'' response.write vuelta & "<br>" & chr(13)

  Fila = ""

  stitu = "<big>" & bDescri & "</big>"

  if scambios = "SI" then
     stitu = stitu & "&nbsp;(Cambios)"
  end if

  Celda = ofv.GenCelda("","ttt","","60%","","","",stitu,"si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

' Referencia
  if scontdesc <> "NO" then
     Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100%","","","",replace(scontdesc,chr(13),"<br>"),"si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  end if


'''''''''''''''''''''''''''''''''''''''''''''''''



  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"


  sHTML = sHTML & GENFMTCONTROLES(nCodigo,STIPO)


  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

 ''' call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botonera

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")


  'cambios
   Listar = ""

  if origen <> "ED" then

   if bVersion = "SI" and SESSION("SHWMOD") <> "N" then

      if  sCambios = "NO" then
          xcambios = "SI"
      else
          xcambios = "NO"
      end if

    '  xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
      sAccion = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
      sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & nid & "&scambios=" & xcambios
      Listar  = genbotoncambios(sAccion) 'ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

   end if
  end if




  Agregar = Celt2
  Buscar  = Celt1
  Celda   = Listar

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)




  if nID <> "no" then call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>