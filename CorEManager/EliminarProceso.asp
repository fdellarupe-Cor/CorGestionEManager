<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipobjeto")
nCodigo = request.querystring("Codobjeto")
sOPCION = request.querystring("OPCION")


on error resume next

set cn = ofv.conectar(ofv.strconn2)



'''response.write "<br><br><br>" & ncodigo & "-" & stipo & "<br>"


if sOPCION = "ACC" then

   strsql = "Delete From objaccion Where objid = " & ncodigo
   call ofv.crearconsultaEx(StrSql,cn,1,volver)
   set errores=cn.errors

else


strsql = "Select * From Objetos Where codobjeto = " & nCodigo
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if not rs.eof then
   sObjeto = rs(3)
   stipobj = rs(2)
   sCodiden =  rs(1)
end if
call ofv.cerrarconsulta(rs)

        Redim preserve aTablas(33)

        aTablas(0) =  "Objaccioncontrol"
        aTablas(1) =  "Objaccion"
        aTablas(2) =  "Objalcance"
        aTablas(3) =  "Objconsid"
        aTablas(4) =  "Objmain"
        aTablas(5) =  "Objcontrol2"
        aTablas(6) =  "Objcontrol"
        aTablas(7) =  "Objesquema2"
        aTablas(8) =  "Objesquema"
        aTablas(9) =  "Objproced"
        aTablas(10) = "Objhtm"
        aTablas(11) = "Objtabcon"
        aTablas(12) = "Objtabblk"
        aTablas(13) = "Objtabdef"
        aTablas(14) = "Objestprty"
        aTablas(15) = "Objestdet"
        aTablas(16) = "Objestr"
        aTablas(17) = "Objsector"
        aTablas(18) = "Objeven"
        aTablas(19) = "Objprty"
        aTablas(20) = "Objsdk"
        aTablas(21) = "Objsdkdet"
        aTablas(22) = "Objwfl"
        aTablas(23) = "Objwfldet"
        aTablas(24) = "Objdist"
        aTablas(25) = "Objdistprty"
        aTablas(26) = "ObjDSCABR"
        aTablas(27) = "ObjLNK"
        aTablas(28) = "Objfaq"
        aTablas(29) = "Objprtyadv"
        aTablas(30) = "ObjTSEGAGR"
        aTablas(31) = "OBJNWEDT"
        aTablas(32) = "Objetos"
        aTablas(33) = "Objperfildoc"

        sPath = server.mappath(request.servervariables("SCRIPT_NAME"))
        sPath = mid(sPath,1,instrrev(sPath,"\",-1)-1)
        sPath = mid(sPath,1,instrrev(sPath,"\",-1))
        sPathObj = sPath & "CarpetaObjetivos\"
        sPathAlc = sPath & "CarpetaAlcance\"
        sPathCon = sPath & "CarpetaConsideracion\"


          for nz = 0 to Ubound(aTablas)

            if Ucase(aTablas(nz)) <> Ucase("objaccioncontrol") then
               strsql = "Delete From " & aTablas(nz) & " Where codobjeto = " & ncodigo
               call ofv.crearconsultaEx(StrSql,cn,1,volver)
               set errores=cn.errors
               if errores.count <> 0 then exit for
            else


               strsql = "Select objaccioncontrol.objid From objaccioncontrol, "
               strsql = strsql & " objproced WHERE objaccioncontrol.codproced = "
               strsql = strsql & " objproced.codproced "
               strsql = strsql & " AND objproced.codobjeto = " & ncodigo
               set rsdel = ofv.crearconsultaEx(strsql,cn,1,volver)
               do until rsdel.eof
                  strsql = "Delete From objaccioncontrol Where objid = " & rsdel(0)
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)
                  set errores=cn.errors
                  if errores.count <> 0 then exit for
               rsdel.movenext
               loop
               call ofv.cerrarconsulta(rsdel)
            end if
          next


          strsql = "Delete From Relaobj Where codojasoc = " & ncodigo
          strsql = strsql & " and  tipobjasoc = '" & stipo & "' "
          call ofv.crearconsultaEx(StrSql,cn,1,volver)
          set errores=cn.errors


          sArchivoObj = sPathObj & "Obj" & ncodigo & ".rtf"
          sArchivoAlc = sPathAlc & "Alc" & ncodigo & ".rtf"
          sArchivoCon = sPathCon & "Con" & ncodigo & ".rtf"
          Set fso = CreateObject("Scripting.FileSystemObject")
          on error resume next
          fso.DeleteFile sArchivoObj
          fso.DeleteFile sArchivoAlc
          fso.DeleteFile sArchivoCon

 end if
 



set errores=cn.errors
if errores.count = 0 then
  response.redirect Session("ASPSTR")
else
  response.write ofv.ErrorDB(cn,"../vuelve.asp")
end if

call ofv.cerrarconn(cn)

%>