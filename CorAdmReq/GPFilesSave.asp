<!--#INCLUDE FILE="asp/clsUpload.asp"-->


<%

Dim Upload
Dim FileName1
Dim FileName2
Dim Folder

  GPITEM = Request.querystring("GPITEM")

response.buffer=true



Set Upload = New clsUpload




folder= session("directorio") & "\"
response.write "folder " & folder
 'response.write shtml
 'response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"

For i = 0 To Upload.Count - 1

		FileName = Upload.Fields(Upload(i).Name).FileName
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

session("Sfilename")=Sfilename
session("Stamano")=Stamano
response.redirect "GPGETpostF.asp?GPITEM=" & GPITEM



%>


