<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTESQUEMAS.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
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
     xsproced = request.querystring("selcom")
  end if

'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK 
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "8" AND SNIV2 <> "0" THEN 

    xsproced =  SNIV2

    nCantRegAMover = "A"
    

 ELSE

    SPLFUNC = ""

 END IF 

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"


'''''''''''''''''''''''''''''''''''''''''''''''''



  sMenux = request.querystring("Menu")
  if sMenux = "PORSECTOR" then
     sSelecx = " and (descri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' "
     sSelecx = sSelecx & " or descri1 = '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "') "
  else
     sSelecx = ""
  end if
  if nCantRegAMover = "x" then
   Strsql = "SELECT * from csesquemasel "
   Strsql = Strsql & "Where codobjeto=" & nCodigo
   Strsql= Strsql & " and objid = " & xsproced & sSelecx
  else
  if nCantRegAMover = "A" then
   Strsql = "SELECT * from csesquemasel "
   Strsql = Strsql & "Where codobjeto=" & nCodigo
   Strsql= Strsql & " and objid = " & xsproced & sSelecx
  else
   Strsql = "SELECT * from csesquemasel "
   Strsql = Strsql & "Where codobjeto=" & nCodigo & sSelecx
   CantRegAMover="0" & request.querystring("CantRegAMover")
   CantRegAMover=cdbl(CantRegAMover)
   nCantRegAMover = CantRegAMover
  end if
  end if
 


  CantRegAMover2 = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  if ofv.explorador = "MSIE" then
    CantRegAMostrar2 = 7
  else
    CantRegAMostrar2 = 5
  end if

  sMenu = Session("menu")
  bModificar = session("Modificar")
  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_selcom=" & xsproced
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)



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


 '' sHTML = sHTML & ofv.GenTabla("panel","","","100% border=0 ","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     nID       = rs(0)
     nref      = rs(2)
     bDescri   = rs(1)
     if len(nref) = 0 then nref = "NO"
''''''''''''''''''
      stexto1 = "Seleccionar&nbsp;Esquema"
      Celdat  = stexto1

      sAccion = "javascript:void(0);"
      sOnClik = " onclick=""document.all.TIND2.style.display='block';document.all.TIND1.style.display='none';"" "


      stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   ''   Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

      stexto1 = "Esquemas"

      Celdat  = "<td class=uts >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   ''   Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")



''''''''''''''''''''''
     Strsql = "SELECT * from csesquemasel "
     Strsql = Strsql & "Where codobjeto=" & nCodigo & sSelecx

     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do while not rsx.eof

       if clng(rsx(0))   = clng(rs(0).value) then
          sfidkD  = sfidkC
          stexto1 = rsx(1)
          stextoact = stexto1
       else
          sfidkD  = sfidkA
          stexto1 = rsx(1)
          sAccion  = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
          sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & rsx(0)
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

  else
      nID = "no"

  end if

  Fila = ""

  stitu = "<big>" & bDescri & "</big>"

  if scambios = "SI" then
     stitu = stitu & "&nbsp;(Cambios)"
  end if

  Celda = ofv.GenCelda("","ttt","","60%","","","",stitu,"si")
  Fila  = fila & ofv.GenRow("","","","","","","",Celda,"si")

' Referencia 
  if nref <> "NO" then
     nref = replace(nref,chr(13),"<br>")
  ELSE
     nref = "&nbsp;"
  end if
  Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100%","","","",nref,"si")
  Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")
'''''''''''''''''''''''''''''''''''''''''''''''''

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  sHTML = sHTML &  "<div id='mt2' style='display:none;height:350px;overflow:auto;" & styl & "' >"


  sHTML = sHTML & GENFMTESQUEMAS(nCodigo,STIPO)

  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

 '' call ofv.cerrarconn(cn)

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

      xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
      sAccion = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
      sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & nid & "&scambios=" & xcambios
      Listar  = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

   end if
  end if




  Agregar = Celt2
  Buscar  = Celt1
  Celda   = Listar

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>