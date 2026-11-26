<%

 Response.Charset="ISO-8859-1"
'response.write server.mappath(request.servervariables("script_name"))
 ruta=server.mappath(request.servervariables("script_name"))

 pos=instrrev(ruta,"\")
 ruta=mid(ruta,1,pos-1)
 pos=instrrev(ruta,"\")
 ruta=mid(ruta,1,pos)

 XARCHIVO=request.querystring("XARCHIVO")
 folder=request.querystring("folder")
 OP= request.querystring("op")
 NEWFOLDER=request.querystring("NEWFOLDER")

 Set fso = SERVER.CreateObject("Scripting.FileSystemObject")
 Set Folder = fso.getfolder(ruta  & folder)



 select case OP

 case "RF"
         'response.write ruta & folder & OP
        set sfolder = Folder.SubFolders

          For Each F in sfolder
            response.write f.name & "|" '& "<BR>"
          next

case "DF"

     '  set sfolder = Folder.SubFolders
      ' fso.DeleteFile(ruta & arch)
       ARCH=REQUEST.FORM("ARCH")
      fso.DeleteFile(ruta & arch)

case "AF"

       set sfolder = Folder.SubFolders
       sFolder.add(NEWFOLDER)


 case "RFILES"

       Set Files = Folder.Files

          For Each F in files
             response.write f.name &"###" & f.DateCreated  & "|"
          next

 CASE "F"

       Xarchivo=RUTA & "CARPETASCMD\" & XARCHIVO
      ' RESPONSE.WRITE XARCHIVO
       if fso.FileExists(trim(Xarchivo)) then
              Set fff1 = fso.openTextFile(trim(Xarchivo), 1, -1)
              do while fff1.AtEndOfstream <> true
                 response.write fff1.readline
              loop
              fff1.close
      end if

 End Select
%>