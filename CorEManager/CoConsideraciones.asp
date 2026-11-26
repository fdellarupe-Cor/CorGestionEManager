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

  sHTML = ofv.FormHeader(session("FormN")) '''ofv.FormBody2
  sHTML = sHTML & "<center><br>"

  sHTML = sHTML & "<div style='height:90%;overflow:auto;' >"

  sHTML = sHTML & ofv.GenTabla("","","","90% border=0 ","",0,0,0,0,"","","","no")

  set cn = ofv.conectar(ofv.strconn2)

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

  if isnull(tobjetivo) then tobjetivo = "OBJETIVO*"
  if isnull(TAlcance) then TAlcance = "ALCANCE*"
  if isnull(TConsideracion) then TConsideracion = "CONSIDERACIONES*"

  tobjetivo = replace(tobjetivo," ","&nbsp;")
  TAlcance = replace(TAlcance," ","&nbsp;")
  TConsideracion = replace(TConsideracion," ","&nbsp;")

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
 h1 = ofv.GenCelda("","tt style={border-style:outset;border-width:2;} ","","","","","","<b><p align=left >" & tOBJETIVO & "</p></b>","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = server.createobject("CorRtfParser.RtfParser")
    aaa =  session("Parser").RtfToHtml(xxx)
 h1 = ofv.GenCelda("","","","100%","","","","<BR>" & replace(aaa,"&nbsp;"," ",1,-1,1),"si")
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
 h1 = ofv.GenCelda("","TT style={border-style:outset;border-width:2;} ","","","","","","<b><p align=left >" & tALCANCE & "</p></b>","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = server.createobject("CorRtfParser.RtfParser")
    aaa =  session("Parser").RtfToHtml(xxx)
 h1 = ofv.GenCelda("","","","100%","","","","<BR>" & replace(aaa,"&nbsp;"," ",1,-1,1),"si")
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
 h1 = ofv.GenCelda("","TT style={border-style:outset;border-width:2;} ","","","","","","<b><p align=left >" & tCONSIDERACION & "</p></b>","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    set session("Parser") = server.createobject("CorRtfParser.RtfParser")
    aaa =  session("Parser").RtfToHtml(xxx)
 h1 = ofv.GenCelda("","","","100%","","","","<BR>" & replace(aaa,"&nbsp;"," ",1,-1,1),"si")
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

 h1 = ofv.GenCelda("","","","300","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","","","300","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
 h1 = ofv.GenCelda("","","","300","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    sHTML = sHTML & ofv.GenFTabla()

  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</center><br>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

Response.flush

%>