
<!--#INCLUDE FILE="asp/clsUpload.asp"-->

<%




Dim Upload
Dim FileName1
Dim FileName2
Dim Folder


response.buffer=true

ncodigo=request.querystring("codigo")
stipo=request.querystring("tipo")

Set Upload = New clsUpload
upload.Sbase64="SI"


UPLOAD.Sbase64NameArch= ncodigo &stipo & ".PDF"

 ruta =server.mappath(request.servervariables("script_name"))
 pos=instrrev(ruta,"\")
  ruta=mid(ruta,1,pos-1)
  pos=instrrev(ruta,"\")
 ruta=mid(ruta,1,pos)

folderz= ruta





 'response.write shtml
 'response.write "<center><hr><h2>Transferencia Ok.</h2><hr><br></center>"
'response.write folder
For i = 0 To Upload.Count - 1

		'FileName = Upload.Fields(Upload(i).Name).FileName
'response.write replace(Upload.Fields("DIRC").VALUE  &"","/","\")
		 IF Upload(i).Name <> "DIRC" THEN
		  xdir=replace(Upload.Fields("DIRC").VALUE  &"","/","\")
		 ' FileName =  xdir & Upload.Fields(Upload(i).Name).FileName
	  	 upload(Upload(i).Name).saveas folderz &  xdir & UPLOAD.Sbase64NameArch
'		 response.write folderz &  xdir & UPLOAD.Sbase64NameArch
		end if
Next





   response.redirect "../../../coremanager/CursogramaAdmin.asp?op=SI&Codobjeto=" & ncodigo & "&Tipobjeto=" & sTipo

Set Upload = Nothing



%>
