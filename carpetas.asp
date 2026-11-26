<%

 Set fso = SERVER.CreateObject("Scripting.FileSystemObject")

 Set wwFolder = fso.getfolder("c:\webintranet\GrupoPetersen\CorEManager\Emp1\Estructura")
dim resulta
resulta=""
sub  Listar(byref Directorio,byref resulta,wd)

For Each Subdirectorio In Directorio.SubFolders
'List1.AddItem Directorio.Path & "\" & Subdirectorio.Name
XPATH=Directorio.Path
X=INSTR(XPATH ,wd)
XPATH=mid(XPATH,x,len(XPATH))
resulta=resulta  & XPATH  & "\" & Subdirectorio.Name & "<BR>"
Listar Subdirectorio, resulta,wd
Next
end sub


call listar(wwFolder,resulta,"\Emp1\")

response.write resulta
%>