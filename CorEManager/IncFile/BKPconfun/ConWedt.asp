<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  response.buffer = true

  sTitulo = request.querystring("Titulo")
  sTitulo = ofv.convertircar(sTitulo,"_"," ")

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = nCodigo

  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
  sNIV3   = request.querystring("NIV3")
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")

'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK 
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "2" AND SNIV2 <> "0" THEN 


       SPLFUNC = "<script language='javascript' > creawin('T" & SNIV2 & "');</script>"

   

 ELSE

    SPLFUNC = ""

 END IF 

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"


'''''''''''''''''''''''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)

 sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

 shtml = shtml & "<style>" & chr(13)
 shtml = shtml & " p{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";"
 shtml = shtml & " font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;"
 shtml = shtml & " margin-top:0px;margin-bottom:0px;}"
 shtml = shtml & "</style>" & chr(13)

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------

  Fila = ""

  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"



  sHTML = sHTML & GENFMTWEDT(nCodigo,STIPO)


'''''''
  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"


''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

  call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''



  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

  sHTML   = sHTML & SPLFUNC


  'Botonera


  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")

  Agregar = Celt2
  Buscar  = Celt1
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Response.flush





%>