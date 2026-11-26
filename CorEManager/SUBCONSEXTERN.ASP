<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

  session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"
  BSQFORMULARIO=request.querystring("FORMULARIOS")
  ing=request.querystring("ing")

  sMenu = request.querystring("Menu")
  sNombre = request.querystring("Nombre")
  sTema = request.form("codtema")
  if sTema = "" or sTema = "nsnc" then sTema = 0
  sSubtema = request.form("codprodserv")
  if sSubtema = "" or sSubtema = "nsnc" then sSubtema = 0
  sEsquema = request.form("codobjeto")
  if sEsquema = "" or sEsquema = "nsnc" then sEsquema = 0
  xComp = ""
  xComp = request.form("XComplem")
  if BSQFORMULARIO= "" then BSQFORMULARIO="D"


  sNombren = sNombre

  sNombre = "Consulta por " & sNombre
  snombre = ucase(snombre)




  'response.write  smenu
 '' set cn = ofv.conectar(ofv.strconn2)

 xx4 = request.form("BUSCAR") & ""
' RESPONSE.WRITE sNombreN
' RESPONSE.WRITE xx4


  tt="&nbsp;"

  select case UCase(sNombreN)
    case "ESTADO"


    case "SECTOR"

    case "FECHA"

    case "INSTRUMENTO"


    case "PRODUCTO"


    case "NUMEROCIRCULAR"

    CASE "BUSCADOR"

      xx4 = request.form("BUSCAR") & ""

            WKFREDIR = ""
	       ' WKFREDIR = WKFREDIR & "http://googal2/search?client=default_frontend&output=xml_no_dtd&proxystylesheet=default_frontend&ie=UTF-8&oe=UTF-8"
	       ' WKFREDIR = WKFREDIR & ""
	       ' WKFREDIR = WKFREDIR & "&as_q=" & xx4 & ""
	       ' WKFREDIR = WKFREDIR & ""
	       ' WKFREDIR = WKFREDIR & "&num=10&site=claridad&btnG=B%C3%BAsqueda+de+Google&as_epq=&as_oq=&as_eq=&lr=&as_ft=i&as_filetype=&as_occt=any"
      		'WKFREDIR = WKFREDIR & "&as_dt=i&as_sitesearch=&sort=&access=p&as_lq="

      		 WKFREDIR = WKFREDIR & "http://googal2/search?q="
      		 WKFREDIR = WKFREDIR & ""
      		 WKFREDIR = WKFREDIR  & xx4 & ""
      		 WKFREDIR = WKFREDIR & ""
      		 WKFREDIR = WKFREDIR & "+site%3Ahttp%3A%2F%2Fwebaplic1.bancogalicia.com.ar%3A8109&site=default_collection&access=p&sort=date%3AD%3AL%3Ad1&client=default_frontend&lr=&entqr=0&oe=UTF-8&ie=UTF-8&proxystylesheet=default_frontend&num=10&ud=1&output=xml_no_dtd"


           RESPONSE.redirect wkfredir


CASE "BUSCADOREX"

      xx4 = request.form("BUSCAR") & ""

            WKFREDIR = ""


' RESPONSE.write "SEARCHex.ASP?ing="&ing&"&FORMULARIOS="& BSQFORMULARIO &"&YIDER=" & xx4

          RESPONSE.redirect "SEARCHex.ASP?ing="&ing&"&FORMULARIOS="& BSQFORMULARIO &"&YIDER=" & xx4



CASE "BUSCADOREXF"

      xx4 = request.form("BUSCAR") & ""

            WKFREDIR = ""




           RESPONSE.redirect "SEARCHexF.ASP?YIDER=" & xx4

  end select

 END IF



%>