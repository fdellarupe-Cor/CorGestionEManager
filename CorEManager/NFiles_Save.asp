<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->


<!--#INCLUDE FILE="asp/clsUpload.asp"-->


<%

Dim Upload
Dim FileName1
Dim FileName2
Dim Folder


response.buffer=true
sacc = Request.querystring("acc")
idpedido = Request.querystring("id")
seg=REQUEST.querystring("seg")



Set Upload = New clsUpload




folder= session("directorio") & "\"
'response.write "folder " & folder
 'response.write shtml
 'response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"

For i = 0 To Upload.Count - 1

		FileName = Upload.Fields(Upload(i).Name).FileName
		upload(Upload(i).Name).saveas folder &  filename
	'	response.write folder &  filename
Next



 '   response.write "<TABLE>"
Sfilename=""
Stamano=""
	For i = 0 To Upload.Count - 1

	   Sfilename= Sfilename & Upload.Fields(Upload(i).Name).filename  & "#zzzz#"

	   Stamano=Stamano & Upload.Fields(Upload(i).Name).Length  & "#zzzz#"

    Next

'Set Upload = Nothing

session("Sfilename")=Sfilename
session("Stamano")=Stamano
if sacc ="ALTA" OR sacc ="MODI"  THEN
set cn = ofv.conectar(ofv.strconn2)
For i = 0 To Upload.Count - 1
       NOMBRE=Upload.Fields(Upload(i).Name).filename
    '   if Session("EMPMUN")  = "S" then
    '   sdestino="emp" & session("sgu_site") &"/PEDIDOS/"
    '     sql=" insert into PEDIDOSARCH (IDPEDIDO,DESCRIPCION,ARCHIVO,POS) values (" & idpedido & ",'" & NOMBRE & "','" & sdestino & NOMBRE &"',10)"
     ''  else
       sql=" insert into PEDIDOSARCH (IDPEDIDO,DESCRIPCION,ARCHIVO,POS) values (" & idpedido & ",'" & NOMBRE & "','PEDIDOS/"  & NOMBRE &"',10)"
'end if
      call ofv.crearconsultaEx(sql,cn,1,parametros)
 Next
 call ofv.cerrarconn(cn)

Set Upload = Nothing

if seg & "" <> "" then
response.redirect "CREARPEDIDO.asp?idpedido=" & idpedido & "&seg=" & seg & "&PEDIDO=" & SACC
else
response.redirect "CrearNormativa.asp?idpedido=" & idpedido & "&seg=" & seg & "&PEDIDO=" & SACC
end if

ELSE
Set Upload = Nothing
response.redirect "repostF.asp"
END IF



%>


