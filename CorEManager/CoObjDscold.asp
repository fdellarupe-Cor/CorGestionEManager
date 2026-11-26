<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'  response.buffer = true

  sTitulo = request.querystring("Titulo")
  sTitulo = ofv.convertircar(sTitulo,"_"," ")

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  ID = nCodigo

  sHTML = ofv.FormBody2

  sHTML = sHTML & "<center><br>"

  sHTML = sHTML & ofv.GenTabla("","","","80% style={border-style:solid;border-width:1;} valign=middle margin=3  ","",0,0,0,0,"","","","no")

  set cn = ofv.conectar(ofv.strconn2)



  nID = ID

  stitulo = " "
  smuestra = "NO"
  sdescri = " "


  Strsql = "Select muestra,titulo,descri From Objdscabr Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then 
     stitulo = rs(1)
     smuestra = rs(0)
     sdescri = rs(2)
  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","1","0","","","","no")

  if isnull(smuestra) then smuestra = "NO" 
  if isnull(sdescri) then sdescri = " " 

'Objetivo
  if smuestra = "NO" then

     h1 = ofv.GenCelda("","cc","","","","","","SIN&nbsp;DESCRIPCION","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  else

'     h1 = ofv.GenCelda("","","","","","","","&nbsp;","si")
'    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
     h1 = ofv.GenCelda("","tt style={border-style:outset;border-width:2;} ","","","","","","<font face=tahoma >" & stitulo & ":</font>","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
'     h1 = ofv.GenCelda("","","","","","","","&nbsp;","si")
'    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
     h1 = ofv.GenCelda("","cc","","","","","","<br><p align=justify  style='{font-family:tahoma;font-size:10pt;}' >" & replace(sdescri,chr(13),"</p><p align=justify >") & "</p><br>","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  end if
'     h1 = ofv.GenCelda("","","","","","","","&nbsp;","si")
'    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")   
'     h1 = ofv.GenCelda("","","","","","","","&nbsp;","si")
'    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    sHTML = sHTML & ofv.GenFTabla()
  sHTML = sHTML & "</center><br>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>