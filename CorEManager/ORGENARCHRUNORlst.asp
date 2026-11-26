<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

  response.expires=0
  response.ContentType="text/plain"


  sopcion         = request.querystring("opcion") & ""
  lote            = request.querystring("lote")
  ARCHIVO         = request.querystring("ARCHIVO")


  sHTML = ""

  APLCHECK()
  set cn = ofv.conectar(ofv.strconn2)

  if sopcion = "1" then wwwtarch = "RIESGOOP" 
  if sopcion = "2" then wwwtarch = "IDCD" 
  if sopcion = "3" then wwwtarch = "IDPRES" 

  if sopcion = "1R" then wwwtarch = "RRIESGOOP" 
  if sopcion = "2R" then wwwtarch = "IDCDR" 
  if sopcion = "3R" then wwwtarch = "IDPRESRR" 

  wwwFile = ARCHIVO

  Response.AddHeader "Content-Disposition", "attachment; filename=" & wwwFile & ";"




  Strsql = "select * from  orrepositor where tlote = '" & lote & "' and tarchivo = '" & wwwtarch & "' order by tsecuen"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  Do While (Not rs.EOF) 

    response.write rs("tdatos") & chr(13) & chr(10)
    rs.Movenext
  Loop

%>