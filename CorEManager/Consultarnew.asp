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

IF sACC = "" THEN sACC = "NM"

if (session("corapli") = "CLI111557214260398409701") then
	response.redirect "EditarDoc.asp?ACC=" & sACC & "&Tipo=" & sTipo & "&MAINAPL=" & SMAINAPL & "&Perfil=" & sPerfil & "&Codigo=" & nCodigo & "&Funcion=CD1" & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & sDatos
else
	response.redirect "EditarDoc.asp?ACC=" & sACC & "&Tipo=" & sTipo & "&MAINAPL=" & SMAINAPL & "&Perfil=" & sPerfil & "&Codigo=" & nCodigo & "&Funcion=CD" & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5 & sDatos
end if


%>