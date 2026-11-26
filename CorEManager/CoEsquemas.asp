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
  If BrowserName() = "IE" OR BrowserName() = "CHROME" then
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
   shtml= shtml &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & chr(13)

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  Fila = ""

''''''''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)

''  response.write " nCodigo " & nCodigo & "</div>"

  strsqlx = "Select Version,estado,objetos.objun From Objprty "
  strsqlx = strsqlx & " inner join objetos on(objetos.codobjeto=Objprty.codobjeto) "
  strsqlx=strsqlx & " Where Objprty.codobjeto = " & nCodigo

''  response.write " strsqlx " & strsqlx & "</div>"

  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  Cobjun=1
  if not rs.eof then
     Cobjun=rs(2)
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  if Cobjun=1 and Session("EMPMUN") = "S" then
  sql="exec SpSectoresCorp " & nCodigo &",'" & sTipo & "'," & session("un")
'     response.write sql
  Set rsSecCorp = ofv.crearconsultaEx(sql,cn,1,parametros)

	 if not rsSecCorp.eof then

		 ASecCorp =rsSecCorp.getrows
	  end if
	Call ofv.cerrarconsulta(rsSecCorp)
     end if


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"
'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     nID       = rs(0)
     nref      = rs(2)
  '   bDescri   = rs(1)
     if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
	 		SECD  = rs(1)
	 		SECD = FindInArray(ASecCorp,6,rs("eco"),1,SECD)
	         bDescri=SECD & "-" & rs("evento") & "-" & rs("descri")
	 	else
	 		bDescri = rs(1)
    end if

     if len(nref) = 0 then nref = "NO"
''''''''''''''''''
      stexto1 = "Seleccionar&nbsp;Esquema"
      Celdat  = stexto1

      sAccion = "javascript:void(0);"
      sOnClik = " onclick=""document.getElementById('TIND2').style.display='';document.getElementById('TIND1').style.display='none';"" "


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

          if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
			SECD  = rsx(1)
			SECD = FindInArray(ASecCorp,6,rsx("eco"),1,SECD)

			stexto1=SECD & "-" & rsx("evento") & "-" & rsx("descri")
			else
			stexto1 = rsx(1)
		  end if

          sAccion  = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
          sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & rsx(0)
          sOnClik1 = ""
          sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

       Celdat  = stexto1



      stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
      Celdat  = Celdat & "<td class=ut  >" & stexto1  & "zz" & "</td>"
      Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")
      end if

      rsx.movenext
   loop

''''''''''''''''''''''''''''''''

  else
      nID = "no"

  end if

  Fila = ""

  stitu = bDescri

  if scambios = "SI" then
     stitu = stitu & "&nbsp;(Cambios)"
  end if

  Celda = ofv.GenCelda("","ttt","","60%","","","",stitu,"si")
  Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")

' Referencia
  if nref <> "NO" then
     nref = replace(nref,chr(13),"<br>")
  ELSE
     nref = "&nbsp;"
  end if
  Celda = ofv.GenCelda("","utg","","100%","","","",nref,"si")
  Fila  = fila & ofv.GenRow("","","","","25","","",Celda,"si")
'''''''''''''''''''''''''''''''''''''''''''''''''


  sHTML = sHTML &  "<div id='mt2' style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

  sHTML = sHTML &  ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","100%","","0","","4","0","","",Fila,"si")

  sHTML = sHTML & GENFMTESQUEMAS(nCodigo,STIPO)

  sHTML = sHTML & "</span></div>"

  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","ut","","100%","","","","&nbsp;","si")
  Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

 '' call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botonera
  Celt1 = ""
''  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")



  'cambios
   Listar = ""

  if origen <> "ED" then
   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

        if  sCambios = "NO" then
          xcambios = "SI"
        else
          xcambios = "NO"
        end if


        'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
        sAccion = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
        sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & nid & "&scambios=" & xcambios
        Listar  = genbotoncambios(sAccion) 'ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

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

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>