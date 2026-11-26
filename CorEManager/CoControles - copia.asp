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
''response.write sMenux & " por sMenux " & "<br>"
  if sMenux = "PORSECTOR" then
     sSelecx = " and descri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' "
  else
     sSelecx = ""
  end if
  if nCantRegAMover = "x" then
     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo
     Strsql = Strsql & " and objid = " & xsproced & sSelecx
    'response.write " por aca -70- " & "<br>"
   else
  '''response.write nCantRegAMover & " nCantRegAMover " & "<br>"
  if nCantRegAMover = "A" then
     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo
     Strsql = Strsql & " and objid = " & xsproced & sSelecx
    'response.write " por aca -76- " & "<br>"
    'response.write Strsql & " strsql " & "<br>"
    else
     Strsql = "SELECT * from cscontrolsel "
     Strsql = Strsql & "Where codobjeto = " & nCodigo & sSelecx
    'Strsql = Strsql & " and seldes is NOT NULL and descri is NOT NULL "
     CantRegAMover = "0" & request.querystring("CantRegAMover")
     CantRegAMover = cdbl(CantRegAMover)
     nCantRegAMover = CantRegAMover
    'response.write " por aca -83- " & "<br>"
    'response.write Strsql & " strsql " & "<br>"
    'response.write CantRegAMover & " CantRegAMover " & "<br>"

  end if
  end if
 'response.write Strsql & " strsql " & "<br>"

  CantRegAMover2  = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  If BrowserName() = "IE" OR BrowserName() = "CHROME" then
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
'  shtml= shtml &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & chr(13)
 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  Fila = ""

''''''''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)

  strsqlx = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
 'response.write strsqlx & " strsqlx " & "<br>"
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

' response.write strsql & "<br>" & chr(13)

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
      sOnClik = " onclick=""document.getElementById('TIND2').style.display='';document.getElementById('TIND1').style.display='none';"" "


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
    'Strsql = Strsql & " and seldes is NOT NULL and descri is NOT NULL "
  '  response.write strsql & " strsql " & "<br>"
     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    'contar = 0
     do while not rsx.eof

     ' contar = contar + 1
     ' response.write contar & "contar " & "<br>"

       if clng(rsx(0))   = clng(rs(0).value) then
          sfidkD  = sfidkC
          stexto1 = rsx(1)
          stextoact = stexto1
         'response.write stexto1 & " stexto1 por si " & "<br>"
       else
          sfidkD  = sfidkA
          stexto1 = rsx(1)
         'response.write stexto1 & " stexto1 por else " & "<br>"
          sAccion  = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
          sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & rsx(0)
          sOnClik1 = ""
          sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

       Celdat  = stexto1

       stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
       stexto1 = stexto1 & "   >" & Celdat & "</a>"

       Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
       Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"

     ''response.write Celdat & " Celdat " & "<br>"

       Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

     ''response.write Filat & " Filat " & "<br>"

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

  stitu = bDescri

  if scambios = "SI" then
     stitu = stitu & "&nbsp;(Cambios)"
  end if

  Celda = ofv.GenCelda("","ttt","","60%","","","",stitu,"si")
  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

' Referencia
  if scontdesc <> "NO" then
     Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100%","","","",replace(scontdesc,chr(13),"<br>"),"si")
     Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")
  end if


'''''''''''''''''''''''''''''''''''''''''''''''''


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

  sHTML = sHTML &  ofv.GenTabla("","aaa","","100%","","0","","4","0","","",Fila,"si") & "<br>"


  sHTML = sHTML & GENFMTCONTROLES(nCodigo,STIPO)


  sHTML = sHTML & "</span></div>"

  sHTML = sHTML & "</div>"


  Celda = ofv.GenCelda("","ut","","100%","","","","&nbsp;","si")
  Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

 ''' call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botonera
  Celt1 = ""
 ''' Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")


  'cambios
   Listar = ""
'response.write origen & "222"
  if origen <> "ED" then

   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

        if  sCambios = "NO" then
          xcambios = "SI"
        else
          xcambios = "NO"
        end if

       ' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
        sAccion = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
        sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & nid & "&scambios=" & xcambios
        'Listar  = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")
         Listar  = genbotoncambios(sAccion)

      elseif SESSION("VIEWMOD") <> "N" then

        ' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)


      end if




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