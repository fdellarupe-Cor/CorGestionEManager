<!--#INCLUDE FILE="asp/clsUpload.asp"-->


<%

Dim Upload
Dim FileName1
Dim FileName2
Dim Folder


response.buffer=true



Set Upload = New clsUpload

  nCodigo = session("edcodobj")
  sTipo   = session("edtipobj")


folder= session("directorio") & "\"
response.write "folder " & folder
 'response.write shtml
 'response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"

For i = 0 To Upload.Count - 1

		FileName = nCodigo & sTipo & "-" & Upload.Fields(Upload(i).Name).FileName
		upload(Upload(i).Name).saveas folder &  filename
Next



 '   response.write "<TABLE>"
Sfilename=""
Stamano=""
	For i = 0 To Upload.Count - 1

	   Sfilename= Sfilename & Upload.Fields(Upload(i).Name).filename  & "#zzzz#"

	   Stamano=Stamano & Upload.Fields(Upload(i).Name).Length  & "#zzzz#"

    Next

Set Upload = Nothing

session("Sfilename") = filename
session("Stamano")   = Stamano






response.redirect "EdPrintPnl.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&opcion=UPD"



%>


