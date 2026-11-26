<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0
nCodigo = request.querystring("Codigo")
sTipo   = request.querystring("Tipo")
sPerfil = request.querystring("Perfil")
sNIV1   = request.querystring("NIV1")
sNIV2   = request.querystring("NIV2")
sNIV3   = request.querystring("NIV3")
sNIV4   = request.querystring("NIV4")
sNIV5   = request.querystring("NIV5")
sVolver = request.querystring("Volver")
sDatos  = request.querystring("Datos")
sACC    = request.querystring("ACC")
sMAINAPL  = request.querystring("MAINAPL")
SEGCKTPSC = request.querystring("SEGCKTPSC") & ""
Ysearch=request.querystring("Ysearch")
if Ysearch = "" then
 session("yider")=""
end if


sDOCMODAL  = request.querystring("DOCMODAL")

IF sDOCMODAL  = "CL" THEN
   session("ModMain") = "CL"
ELSEIF sDOCMODAL  = "WK" THEN
   session("ModMain") = "WK"
END IF


IF sACC = "" THEN sACC = "NM"

if (session("corapli") = "CLI111557214260398409701") then
	response.redirect "EditarDoc.asp?ACC=" & sACC & "&Tipo=" & sTipo & "&MAINAPL=" & SMAINAPL & "&Perfil=" & sPerfil & "&Codigo=" & nCodigo & "&Funcion=CD1&SEGCKTPSC=" & SEGCKTPSC & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & sDatos
else
	response.redirect "EditarDoc.asp?ACC=" & sACC & "&Tipo=" & sTipo & "&MAINAPL=" & SMAINAPL & "&Perfil=" & sPerfil & "&Codigo=" & nCodigo & "&Funcion=CD&SEGCKTPSC=" & SEGCKTPSC & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & sDatos
end if


%>