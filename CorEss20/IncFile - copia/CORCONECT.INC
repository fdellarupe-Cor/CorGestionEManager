

<%
'Session.LCID=3082 bc
 ' Response.AddHeader "pragma","no-cache"
 ' Response.AddHeader "cache-control","private"
 ' Response.CacheControl = "no-cache"
 ' Response.Expires = -1
 ' Response.Expiresabsolute = Now() - 1
 ' Response.AddHeader "X-UA-Compatible" ,"IE=EmulateIE7"


  Response.AddHeader "pragma","no-cache"
  Response.AddHeader "cache-control","private"
  Response.AddHeader "Cache-Control", "no-cache, no-store, must-revalidate"
  Response.CacheControl = "no-cache, no-store, must-revalidate"
  Response.AddHeader "Expires", "0" '// Proxies.
  Response.Expires = -1
  Response.Expiresabsolute = Now() - 1


%>

<!-- #INCLUDE FILE="Corstyles.inc" -->

<%

''<!-- #INCLUDE FILE="funcvarias.inc" -->
' ****************************************************************************************
' ** Cor E-Solution V48- Esquema de Configuracion -----------------       Mayo, 2020   **
' ****************************************************************************************

''funcvarias

set ofv = server.createobject("CorWebServerV37.FuncVarias")
'Set ofv = SERVER.CreateObject("CorWebServerNet5.funvarias")

'<!-- #INCLUDE FILE="funcvarias.inc" -->
'set ofv = new funcvarias

'Set ofv.Rq = Request
set ofv.Request = Request
Set ofv.Server = Server
Set ofv.Response = Response
Set ofv.Session = Session
'ofv.SetRequest(Request)
'response.write ofv.Browser3()
session("WEBACC")="N"
session("NUEVOMENU")="S"
session("WEBDESK")="S"

session("menuex")="S"
'12/7/2018 se agrego para habilitar link a formularios en compomentes -> solapa procedimiento/contenido
session("CORMODUFORM")="S"

'8/6/2020 se agrego para habilitar   solicitud Normativa
session("CORMODUSLTNORM")="S"

'8/6/2020 se agrego para habilitar  ABM FORMULARIOS
session("CORMODUSOLABM")="S"
session("CORMENUNW")="N"

 'Session("EMPMUN")=N -> mono empresa  ---  Session("EMPMUN")= S -> multi empresa

'Session("EMPMUN")  = "S"
Session("EMPMUN") = "N"


session("COPIA") = "S"
session("Mcorrecion")="6a"
session("MailComponent")="CDOSYS"
'session("MailComponent")="NOSEND"
Session("SMTPServer")="smtp.gmail.com"
Session("SMTPPort")=25
SESSION("MAILOPC") = "HTML"
'SESSION("MAILOPC") = "TEXTO"

''SESSION("WKFETPU") = "D"
''SESSION("WKFETPU") = "F"

'SESSION("URLTYP")  = "W"
SESSION("URLTYP")  = "V"
SESSION("URLPORT")  = "N"

SESSION("GLOSAUTOR")  = "S"

SESSION("TIPOFECHA")  =0  ' fechadorCor

SESSION("FULLDOCN") = "5"

Session("CORUNQ") = "S"

'En S solo para bnx mail paso wkf
 Session("CORUNQWKF") = "S"

SESSION("WKFETPU") = "U"
SESSION("WKFEUSR") = "COREMANAGER"
SESSION("WKFEMAI") = "@"

Session("Sound") = "Imagenes/xLogon.wav"

'' Session("Modal") = "MSNUA"
 Session("Modal")   = "MSNTT"
 Session("TOPAPL")  = "COREM20"
 Session("CMACC")   = "L"
 Session("CMRES")   = "A"
 Session("CONSET")  = "0"



 Session("EXPLOG")  = "N"
'Session("EXPLOG")  = "S"

 'Session("DELLOG")  = "S"
 Session("DELLOG")  = "N"
 session("USRPRV")  = "I"
' session("USRPRV")  = "BPN"
 Session("FULLGUI") = "S"
 SESSION("RDCGUID") = "S"
 SESSION("FILETIPO")  = "L"

 SESSION("SHWMOD")  = "S"
 SESSION("VIEWMOD")  = "S"

 'session("SUBLOGO") = "Logo Claridad.jpg"
 'session("SUBLOGO") = ""
 'reemplaza site del usuario si se encuentra en la tabla de usuarios en getinfoAD.inc
 'SESSION("SITECORP") = "1"

 if Session("SETLOG") = "S" then Session("EXPLOG") = "N"

%>
<!-- #INCLUDE FILE="CorDbSources.inc" -->
<!-- #INCLUDE FILE="CorLOGF.inc" -->

<!-- #INCLUDE FILE="CorRDR.inc" -->

<!-- #INCLUDE FILE="GetInfo.inc" -->

<%

 CALL GETFECH()

 session("XEXPIR") = XLHORDAT


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ** Cor E-Solution V3 - Esquema de Configuracion -----------------        Sept, 2006   **
'''  COR CONECT - Include
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' ****************************************************************************************
' ** Cor E-Solution V3 - Esquema de Configuracion -----------------        Sept, 2006   **
' ****************************************************************************************
FUNCTION APLCHECK()

  SESSION("APLCHK") = "S"

  IF (session("APLMDL") & "") = "" THEN

       session("APLMDL") = "CP-CP-CP-CP-WK-GP"

  END IF

 'session("APLMDL") = "CP-CP-CP-CP-WK-GP"
 'session("APLMDL") = "CP-CP-CP-CP-WK-BC"
 'session("APLMDL") = "CP-CP-CP-CP-WK-GP"
  session("APLMDL") = "CP-CP-CP-CP-WK-BP"
  WKFWORKA = SPLIT(session("APLMDL"),"-")

  session("APLMDLCS") = WKFWORKA(0)
  session("APLMDLEM") = WKFWORKA(1)
  session("APLMDLOR") = WKFWORKA(2)
  session("APLMDLPL") = WKFWORKA(3)
  session("APLMDLRQ") = WKFWORKA(4)
  session("APLMDCLI") = WKFWORKA(5)

 END FUNCTION


' ****************************************************************************************
' ** Cor E-Solution V3 - Esquema de Configuracion -----------------        Sept, 2006   **
' ****************************************************************************************

%>

