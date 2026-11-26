<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportesex.inc" -->
<!-- #INCLUDE FILE="../coress20/IncFile/CorStyles.inc" -->
<%

dim sbase

Session("usuario") = "GOOGLE"

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  Response.Write sHTML
else


set oPrtG = new ClsReportes

Set oPrtG.Ofv = Ofv


 Sbase =Left(request.servervariables("SCRIPT_NAME"), InstrRev(request.servervariables("SCRIPT_NAME"),"/"))

'  End If


   If session("URLPORT") = "N" Then
        Sbase = "http://" & request.servervariables("SERVER_NAME") & Sbase
   Else
        Sbase = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & Sbase
   End If



response.write corstyler()


response.write   oPrtG.Generar()


END IF

%>