<%

  Response.AddHeader "pragma","no-cache"
  Response.AddHeader "cache-control","private"
  Response.CacheControl = "no-cache"
  Response.Expires = -1
  Response.Expiresabsolute = Now() - 1
'   Response.AddHeader "X-UA-Compatible" ,"IE=EmulateIE7"


%>



<!-- #INCLUDE FILE="Corstyles.inc" -->


<%

''<!-- #INCLUDE FILE="funcvarias.inc" -->
' ****************************************************************************************
' ** Cor E-Solution V34 - Esquema de Configuracion -----------------       Mayo, 2008   **
' ****************************************************************************************

''funcvarias

set ofv = server.createobject("CorWebServerV36.FuncVarias")
''set ofv = new funcvarias


Set ofv.Request = Request
Set ofv.Server = Server
Set ofv.Response = Response
Set ofv.Session = Session

session("WEBACC")="N"

session("WEBDESK")="S"

session("Mcorrecion")="6a"
session("MailComponent")="CDOSYS"
'session("MailComponent")="NOSEND"
Session("SMTPServer")="smtp.gmail.com"
Session("SMTPPort")=25
SESSION("MAILOPC") = "HTML"
'SESSION("MAILOPC") = "TEXTO"

''SESSION("WKFETPU") = "D"
''SESSION("WKFETPU") = "F"

SESSION("URLTYP")  = "W"
'SESSION("URLTYP")  = "V"
SESSION("URLPORT")  = "Y"

SESSION("GLOSAUTOR")  = "S"

SESSION("TIPOFECHA")  =0  ' fechadorCor

SESSION("FULLDOCN") = "5"

Session("CORUNQ") = "N"

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
 '
 Session("EMPMUN")  = "S"
 'Session("EMPMUN") = "N"

 'Session("EXPLOG")  = "N"
Session("EXPLOG")  = "S"

 Session("DELLOG")  = "S"
 'Session("DELLOG")  = "N"
 session("USRPRV")  = "GP"
' session("USRPRV")  = "I"
 Session("FULLGUI") = "S"
 SESSION("RDCGUID") = "S"
 SESSION("FILETIPO")  = "L"

 SESSION("SHWMOD")  = "S"
 SESSION("VIEWMOD")  = "S"

 session("SUBLOGO") = ""
 SESSION("SITECORP") = "1"

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
''       session("APLMDL") = "CP-CP-CP-CP-WK-BC"
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

