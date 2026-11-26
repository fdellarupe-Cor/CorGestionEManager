<!--#INCLUDE FILE="asp/clsUpload.asp"-->


<%

Dim Upload
Dim FileName1
Dim FileName2
Dim Folder


response.buffer=true


'response.write  session("directorio") & "222222222222"
'session("directorio")="\..\ArcComponentes"
Set Upload = New clsUpload

'response.write session("directorio") & "222"
 ruta =server.mappath(request.servervariables("script_name"))
 pos=instrrev(ruta,"\")
  ruta=mid(ruta,1,pos-1)
  pos=instrrev(ruta,"\")
 ruta=mid(ruta,1,pos)

folderz= ruta


For i = 0 To Upload.Count - 1


    IF Upload(i).Name <> "DIRC" THEN
        xdir=replace(Upload.Fields("DIRC").VALUE&"","/","\")
		FileName =  xdir & Upload.Fields(Upload(i).Name).FileName
		upload(Upload(i).Name).saveas folderz    & filename

		response.write folderz    & filename

	ELSE
Upload.Fields("DIRC").VALUE
	END IF
Next

Set objField = Upload.Fields("DIRC")

 '   response.write "<TABLE>"
Sfilename=""
Stamano=""
	For i = 0 To Upload.Count - 1
       IF Upload(i).Name <> "DIRC" THEN
	   Sfilename= Sfilename & xdir & Upload.Fields(Upload(i).Name).filename  & "#zzzz#"

	   Stamano=Stamano & Upload.Fields(Upload(i).Name).Length  & "#zzzz#"
       end if
    Next

Set Upload = Nothing

session("Sfilename")=Sfilename
session("Stamano")=Stamano
response.redirect "repostF.asp"



%>


