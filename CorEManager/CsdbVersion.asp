<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 3
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



  set cn = ofv.conectar(ofv.strconn2)



  Strsql = "Select * from csdbversion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  aax = ofv.convertircar(Ucase(rs(0))," ","")

  aaa = instr(AAX,"SQLSERVER6.5")
  aab = instr(AAX,"SQLSERVER7.0")
  aac = instr(AAX,"ACCESS")
  
  if aaa > 0 then 
response.write "Ms SqlServer 6.5"
  elseif aab > 0 then 
response.write "Ms SqlServer 7.0"
  else 
response.write "Ms Access"
  end if 

  Response.Write "<br>" & rs(0) & " " & aaa & aab & aac


end if


%>