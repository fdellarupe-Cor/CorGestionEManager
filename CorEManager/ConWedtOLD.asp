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



  sHTML = sHTML & ofv.GenTabla("","","","100%","",0,0,0,0,"","","","no")


''''''''''''''''''
      stexto1 = "Capitulos"
      sOnClik = ""
      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


''''''''''''''''''''''



  set cn = ofv.conectar(ofv.strconn2)

'''''''''''


  Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " Order By consecuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     xxxx = NWEDTVER()
  end if

  XNVER = "N"
  XNSEC = 0

 ' Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
 ' XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
 ' set fff = server.createobject("scripting.filesystemobject")

  do until rs.eof

     tobjetivo      = rs(3)

     if isnull(tobjetivo)   then tobjetivo    = " "

     tobjetivo      = "<BIG>" & replace(tobjetivo," ","&nbsp;") & "</BIG>"

  '   Set MyFile = fff.openTextFile(XferTarget0 & rs(5) & ".CSE", 1, -1)

     sObjetivo = ""
 '    do  while MyFile.AtEndOfstream <> true 
 '        sObjetivo = sObjetivo & MyFile.readline
 '    loop 
  '   MyFile.close

     sObjetivo=getReadfile(rs(5) & ".CSE","")

  if sObjetivo <> "" then


'''''''''' arma tabla de indices '''''''''''''''''''''''''
      sPoint  = "T" & rs(0)
      stexto1 = rs(3)
      Celdat  = stexto1

      stexto1 = "<A class=ut href='javascript:void(0);' "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & sfidk1 & "</td>"
      Celdat  = Celdat & "<td class=uti onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



      pmk = "<form name="  & chr(34) & sPoint & chr(34) &  " method=post >"
      pmk = pmk & "<input type=hidden value=a ></form>"

        h1 = ofv.GenCelda("","","","300","","","","&nbsp;" & pmk,"si")
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
        h1 = ofv.GenCelda("","ttt  ","","","","","",tOBJETIVO,"si")
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

        h1 = ofv.GenCelda("","uti","","100%","","","",replace(sObjetivo,"&nbsp;"," ",1,-1,1),"si")
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  end if



    rs.movenext
  loop

  sHTML = sHTML & ""
  sHTML = sHTML & ofv.GenfTabla()



  call ofv.cerrarconsulta(rs)


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



function NWEDTVER()

  XNVER = "N"
  XNSEC = 0

  Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
  XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "carpetaObjetivos\"
  XferTarget2 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaAlcance\"
  XferTarget3 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaConsideracion\"

  abrir1 = XferTarget1 & RtfObjetivo
  abrir2 = XferTarget2 & RtfAlcance
  abrir3 = XferTarget3 & RtfConsid

  set fff = server.createobject("scripting.filesystemobject")




'''''''''''''' Objetivo
''''
  Strsql = "Select ARCHIVO,titulo From ObjMain where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget1 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""
   ' if fff.FileExists(trim(Xarchivo)) then
   '    Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
   '    do while fff1.AtEndOfstream <> true
   '       xxx = xxx & fff1.readline
   '    loop
   '    fff1.close
   
       'xxx=getReadfile2(xarchivo & ".CSE","carpetaObjetivos\")
       xxx=getReadfile2(xarchivo,"carpetaObjetivos\")
       set XNPSR = server.createobject("CorRtfParser.RtfParser")
       XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

 '   end if 
  end if 
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Alcance
''''
  Strsql = "Select ARCHIVO,titulo From ObjAlcance where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget2 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""
    'if fff.FileExists(trim(Xarchivo)) then
    '   Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
    '   do while fff1.AtEndOfstream <> true
    '      xxx = xxx & fff1.readline
    '   loop
    '   fff1.close
      xxx=getReadfile2(xarchivo,"carpetaObjetivos\")
      set XNPSR = server.createobject("CorRtfParser.RtfParser")
      XNPSRFL   =  XNPSR.RtfToHtml(xxx)
      Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
      MyFile.Write XNPSRFL
      MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  '  end if 
  end if 
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Consideraciones generales
''''
  Strsql = "Select ARCHIVO,titulo From ObjConsid where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget3 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""
    'if fff.FileExists(trim(Xarchivo)) then
    '   Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
    '   do while fff1.AtEndOfstream <> true
    '      xxx = xxx & fff1.readline
    '   loop
    '   fff1.close
       set XNPSR = server.createobject("CorRtfParser.RtfParser")
       XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  '  end if 
  end if 
  call ofv.cerrarconsulta(rs1)

  if  XNVER = "S" then

      Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " Order By consecuen"
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if

end function

function wold()





  sPath = server.mappath(request.servervariables("SCRIPT_NAME"))
  sPath = mid(sPath,1,instrrev(sPath,"\",-1))

  nID = ID

  sObjetivo = ""
  sAlcance = ""
  sConsideracion = ""

  Strsql = "Select Archivo,titulo From Objmain Where codobjeto = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sObjetivo = rs(0)
     TObjetivo = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  Strsql = "Select Archivo,titulo From Objalcance Where codobjeto = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sAlcance = rs(0)
     TAlcance = rs(1)
  end if 
  call ofv.cerrarconsulta(rs)

  Strsql = "Select Archivo,titulo From Objconsid Where codobjeto = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sConsideracion = rs(0)
     TConsideracion = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  if isnull(tobjetivo)      then tobjetivo      = "OBJETIVO*"
  if isnull(TAlcance)       then TAlcance       = "ALCANCE*"
  if isnull(TConsideracion) then TConsideracion = "CONSIDERACIONES*"

  tobjetivo      = "<BIG>" & replace(tobjetivo," ","&nbsp;") & "</BIG>"
  TAlcance       = "<BIG>" & replace(TAlcance," ","&nbsp;") & "</BIG>"
  TConsideracion = "<BIG>" & replace(TConsideracion," ","&nbsp;") & "</BIG>"

  hoja=1
  topeflush=50 

'Objetivo
  if sObjetivo <> "" then
    set session("fso") = server.createobject("scripting.filesystemobject")
    rtfFile = sPath & "CarpetaObjetivos\" & sObjetivo
    Set session("tf") = session("fso").openTextFile(rtfFile, 1, -1)

    xxx = ""
    do while session("tf").AtEndOfstream <> true 
      xxx = xxx & session("tf").readline
    loop 
    session("tf").close
    if xxx <> "" then
 h1 = ofv.GenCelda("","","","300","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","ttt  ","","","","","",tOBJETIVO,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = server.createobject("CorRtfParser.RtfParser")
    aaa =  session("Parser").RtfToHtml(xxx)
 h1 = ofv.GenCelda("","vti","","100%","","","",replace(aaa,"&nbsp;"," ",1,-1,1),"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = nothing
    end if
  end if

'Alcance
  if sAlcance <> "" then
    set session("fso") = server.createobject("scripting.filesystemobject")
    rtfFile = sPath & "CarpetaAlcance\" & sAlcance
    Set session("tf") = session("fso").openTextFile(rtfFile, 1, -1)

    xxx = ""
    do while session("tf").AtEndOfstream <> true 
      xxx = xxx & session("tf").readline
    loop 
    session("tf").close
    if xxx <> "" then
 h1 = ofv.GenCelda("","","","300","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","ttt","","","","","",tALCANCE,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = server.createobject("CorRtfParser.RtfParser")
    aaa =  session("Parser").RtfToHtml(xxx)
 h1 = ofv.GenCelda("","vti","","100%","","","",replace(aaa,"&nbsp;"," "),"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = nothing
    end if 
  end if

'Consideraciones
  if sConsideracion <> "" then
    set session("fso") = server.createobject("scripting.filesystemobject")
    rtfFile = sPath & "CarpetaConsideracion\" & sConsideracion
    Set session("tf") = session("fso").openTextFile(rtfFile, 1, -1)

    xxx = ""
    do while session("tf").AtEndOfstream <> true 
      xxx = xxx & session("tf").readline
    loop 
    session("tf").close
    if xxx <> "" then
 h1 = ofv.GenCelda("","","","300","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","ttt","","","","","",tCONSIDERACION,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = server.createobject("CorRtfParser.RtfParser")
    aaa =  session("Parser").RtfToHtml(xxx)
 h1 = ofv.GenCelda("","vti","","100%","","","",replace(aaa,"&nbsp;"," "),"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = nothing
    end if
  end if

if sObjetivo = "" and sAlcance = "" and sConsideracion = "" then
       sestr = "<h5>Sin&nbsp;datos&nbsp;disponibles"
       sestr = sestr & "</h5>" 
 h1 = ofv.GenCelda("","et","","","","","",sestr,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
end if


    sHTML = sHTML & ofv.GenFTabla()


end function

%>