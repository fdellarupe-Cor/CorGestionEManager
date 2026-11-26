<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  xsolapa=request.querystring("xsolapa")

   if xsolapa <> "" then
    session("xsolapaF")=xsolapa
    else
    xsolapa= session("xsolapaF")
  end if
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_xsolapa=" & xsolapa

''RESPONSE.WRITE " --> " & vuelta & " vuelta "

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
 ''RESPONSE.WRITE REQUEST.QUERYSTRING


  OPCION=request.querystring("OPCION")
  stema=cint("0" & request.form("codtema"))
  ssubtema=cint("0" & request.form("subtema"))

  if xsolapa="" then xsolapa="PORESQUEMA"

   sImagenP = "&#0232;"
   zstyle="style='{font-family:" & Session("CMFPNBOTFTFC") & ";color:#666666;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;bapckground-color:#ff0000;font-weight:bold;;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:none;}'"
   sinput = "<a "& zstyle & " href='TIPOSDOCSAVENCER.asp?XSOLAPA=PORESQUEMA'"
   sinput = sinput  & " >AVISOS REVISION</a>"

''   sImagenP = "&#0232;"
''   sinput2 = "<a "& zstyle &"  href='TIPOSDOCSAVENCER.asp?XSOLAPA=PORTEMA'"
''   sinput2 = sinput2  & " >AGREGAR SUBTEMAS</a>"
''   FORMULARIO=sinput & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & sinput2 &"<br><br>"
  
  FORMULARIO=sinput & "<br><br>" 
  
''IF xsolapa="PORTEMA" THEN
''Formulario = FORMULARIO & ofv.Encabezados("Cantidad&nbsp;de&nbsp;Dias","TIPO&nbsp;DE&nbsp;DOCUMENTO")
''ELSE

  Formulario = FORMULARIO & ofv.Encabezados("Cantidad&nbsp;de&nbsp;Dias","TIPO&nbsp;DE&nbsp;DOCUMENTO")

''END IF


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

''RESPONSE.WRITE " --> " & xsolapa & " xsolapa "
''RESPONSE.WRITE " --> " & opcion  & " opcion "

  IF  xsolapa="PORESQUEMA" AND OPCION="si" THEN

    

    sql3="INSERT INTO tipodocsavencer (cantdias,tipodoc,valor1) values ('0','DOCUMENTO','0')"
  ''RESPONSE.WRITE SQL3
    CALL ofv.crearconsultaEx(SQL3,cn,1,parametros)

  end if

  Strsql = "Select * From tipodocsavencer WHERE TIPODOC='DOCUMENTO' Order By CANTDIAS"
  
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi
  ''response.write " " & sForm & " sForm "  
   
    sInput = rs(0)

    'Eliminar
    if bSuprimir and sinput > "0" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=tipodocsavencer&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

''''response.write " --> " & XSOLAPA & " XSOLAPA " 

'' AGREGADO 10102014 PARA DOCUMENTOS A VENCER
''
 if xsolapa="PORESQUEMA" then  

    SCOMSEPx = chr(9)
    scomKk = ""
    scomDd = ""
    scomKk = scomKk & SCOMSEPx & "5"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 5 DIAS"
	  scomKk = scomKk & SCOMSEPx & "10"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 10 DIAS"
	  scomKk = scomKk & SCOMSEPx & "20"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 20 DIAS"
	  scomKk = scomKk & SCOMSEPx & "30"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 30 DIAS"
    scomKk = scomKk & SCOMSEPx & "60"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 60 DIAS"
   
 '' ORREVISION  = dFechan 
 '' SORREVISION = "ANTES DE LOS 5 DIAS"

''' sInput = ofv.gencomboFX(sForm,"ORREVISION",SORREVISION,scomKk,scomDd,"")
    sInput = ofv.gencomboFX(sForm,rs(1).name,rs(1),scomKk,scomDd,"") 
    Celda = Celda & ofv.GenCelda("","VTI","","15%","","","",sInput,"si")   
  
'' FIN AGREGADO 10102014 PARA DOCUMENTOS A VENCER

    StrSql="SELECT * FROM TIPObjeto ORDER BY DESCRI"
	  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    
    SCOMSEP = "" 
	  scomK   = ""
	  scomD   = ""

    svalor1 = rs(3) & ""
 	  do while not rsx1.eof
	  	 
	  	 SCOMSEP = chr(9)
	     scomK = scomK & SCOMSEP & rsx1(0)
	     scomD = scomD & SCOMSEP & rsx1(1)
  
	  rsx1.movenext
	  loop

''''sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dtN","")
''''Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    
    sInput = ofv.gencomboFX(sForm,"VALOR1",svalor1,scomK,scomD,"")
    
 '''response.write " --> " & sInput & " sInput " 
    
    Celda = Celda & ofv.gencelda("","VTI","","30%","","","",sInput,"si")

    Fila = ofv.GenRow("","","","","","","",Celda,"si")

 end if

   if bModificar then
      	        sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      	        sAccion = sAccion & "&tabla=tipodocsavencer&volver=" & vuelta
   end if



    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
    fila=""
    posi = posi + 1
    rs.Movenext
  Loop

  sHTML =   sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""

  if bAgregar then
  
  ''response.write " --> " & bAgregar & " bAgregar " 	
  	
     sAccion = "TIPOSDOCSAVENCER.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=tipodocsavencer&XSOLAPA=" & XSOLAPA
     Agregar = ofv.BotonAgregar(sAccion)
  end if

  'str1 = "Select * From FITROBSQ "    ' orden pagina
  str2 = "Select * From TIPODOCSAVENCER  "    ' orden alfabetico
  Buscar = "" 'ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))



  'sAccion = "NivelePUESTOCons.asp?" & request.querystring
  Celda = "" 'ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>