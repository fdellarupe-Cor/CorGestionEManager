<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipobjeto")
nCodigo = request.querystring("Codobjeto")
sEliminar = request.querystring("Eliminar")

'on error resume next

set cn = ofv.conectar(ofv.strconn2)

nFecini = 0

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

dHoy = clng(date) - nFecini

'''response.write "<br><br><br>" & ncodigo & "-" & stipo & "<br>"

strsql = "Select * From Objetos Where codobjeto = " & nCodigo
set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
if not rs.eof then
   sObjeto = rs(3)
   stipobj = rs(2)
   sCodiden =  rs(1)
end if
call ofv.cerrarconsulta(rs)

sMsg = ""

  if sEliminar = "NSNC" then

'Validacion

    strsql = "Select estado From Objprty Where codobjeto = " & nCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then
       if cint(rs(0)) = 3 then
          strsql = "Select * From ObjLNK Where Ltipo = 'L' and codobjetoL = " & nCodigo
          set rs1 = ofv.crearconsultaEx(StrSql,cn,1,volver)
          If not rs1.eof then
             sMsg = "<br>  No se puede eliminar el esquema <B><U><FONT COLOR=#993333>" & sCodiden & " - " &  sObjeto & "</FONT></U></B> ya que existen documentos que lo referencian<br><br>"
             sMsg = sMsg & "Verifique en <b><U><FONT COLOR=#993333>PROPIEDADES - ESQUEMAS QUE REFERENCIAN EL DOCUMENTO</FONT><U></B>  <br><br>"
          end if
          call ofv.cerrarconsulta(rs1)
       end if
    end if
    call ofv.cerrarconsulta(rs)

   end if



if sEliminar <> "NO" then
  if sMsg = "" then
'''''''
  if sEliminar = "NSNC" then
    sMensaje = "<br>  Esta seguro que desea eliminar <B><U><FONT COLOR=#993333>" & sCodiden & " - " & sObjeto & "</FONT><U></B>  <br><br>"
    dato = "../Mensajes3.asp?volver="
    dato = dato & "asp/EliminarEsquema.asp&Tipobjeto=" & sTipo & "&Codobjeto="
    dato = dato & nCodigo & "&Botones=2&msg=" & ofv.convertircar(sMensaje," ","_") & chr(13)

'response.write "<br>" & dato & "<br>"

    response.redirect dato
  elseif sEliminar = "SI" then
    strsql = "Select estado From Objprty Where codobjeto = " & nCodigo
    set rs = ofv.crearconsultaEx(strsql,cn,1,volver)
    if not rs.eof then
        Desde = 0
        Hasta = 0
        iz = 0
        Redim preserve aCodigo(iz)
        aCodigo(iz) = nCodigo
        Redim preserve atipobj(iz)
        atipobj(iz) = stipobj

        do
          for nz = desde to hasta
            strsql = "Select codojasoc, tipobjasoc From Relaobj Where codobjmaster = "
            strsql = strsql & aCodigo(iz) & " and  tipobjmaster = '" & atipobj(iz) & "' "
            set rsRel = ofv.crearconsultaEx(strsql,cn,1,volver)
            do until rsRel.eof
              iz = iz + 1
              Redim preserve aCodigo(iz)
              aCodigo(iz) = rsRel(0)
              Redim preserve atipobj(iz)
              atipobj(iz) = rsRel(1)
              rsRel.Movenext
            loop
            call ofv.cerrarconsulta(rsRel)
          next
          if hasta = iz then exit do
          desde = hasta + 1
          hasta = iz
        loop

      if cint(rs(0)) = 4 then

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

        for iz = 0 to hasta
          for nz = 0 to Ubound(aTablas)

            if Ucase(aTablas(nz)) <> Ucase("objaccioncontrol") then
               strsql = "Delete From " & aTablas(nz) & " Where codobjeto = " & aCodigo(iz)
               call ofv.crearconsultaEx(StrSql,cn,1,volver)
               set errores=cn.errors
            '   if errores.count <> 0 then exit for
            else


               strsql = "Select objaccioncontrol.objid From objaccioncontrol, "
               strsql = strsql & " objproced WHERE objaccioncontrol.codproced = "
               strsql = strsql & " objproced.codproced "
               strsql = strsql & " AND objproced.codobjeto = " & aCodigo(iz)
               set rsdel = ofv.crearconsultaEx(strsql,cn,1,volver)
               do until rsdel.eof
                  strsql = "Delete From objaccioncontrol Where objid = " & rsdel(0)
                  call ofv.crearconsultaEx(StrSql,cn,1,volver)
                  set errores=cn.errors
           '       if errores.count <> 0 then exit for
               rsdel.movenext
               loop
               call ofv.cerrarconsulta(rsdel)
            end if
          next
         ' if errores.count <> 0 then exit for
          strsql = "Delete From Relaobj Where codojasoc = " & aCodigo(iz)
          strsql = strsql & " and  tipobjasoc = '" & atipobj(iz) & "' "
          call ofv.crearconsultaEx(StrSql,cn,1,volver)
          set errores=cn.errors
          if errores.count <> 0 then exit for
          sArchivoObj = sPathObj & "Obj" & aCodigo(iz) & ".rtf"
          sArchivoAlc = sPathAlc & "Alc" & aCodigo(iz) & ".rtf"
          sArchivoCon = sPathCon & "Con" & aCodigo(iz) & ".rtf"
          Set fso = CreateObject("Scripting.FileSystemObject")
          on error resume next
          fso.DeleteFile sArchivoObj
          fso.DeleteFile sArchivoAlc
          fso.DeleteFile sArchivoCon
        next
      else
        for iz = 0 to hasta
          strsql = "Update Objprtyp Set estado = 4,fecbaja = '" & dHoy & "', "
          strsql = strsql & "elimino = '" & session("Usuario") & "' "
          strsql = strsql & "Where codobjeto = " & aCodigo(iz)
          call ofv.crearconsultaEx(StrSql,cn,1,volver)
          set errores=cn.errors
        '  if errores.count <> 0 then exit for
          if cint(rs(0)) = 1 or cint(rs(0)) = 2 then
             strsql = "Update relaobj Set codobjmaster = 0,tipobjmaster = 'RYS'"
             strsql = strsql & "Where codojasoc = " & aCodigo(iz)
             strsql = strsql & " and  tipobjasoc = '" & atipobj(iz) & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
             set errores=cn.errors
           '  if errores.count <> 0 then exit for
          end if
        next
      end if
    end if
    call ofv.cerrarconsulta(rs)
  end if
  else

    dato = "../Mensajes3.asp?volver="
    dato = dato & "asp/EliminarEsquema.asp&Tipobjeto=" & sTipo & "&Codobjeto="
    dato = dato & nCodigo & "&Botones=1&msg=" & ofv.convertircar(sMsg," ","_") & chr(13)

  ''  dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
  ''  dato = dato & "asp/EliminarEsquema.asp&Tipobjeto=" & sTipo & "&Codobjeto="
  ''  dato = dato & nCodigo & "&Botones=1" & chr(13)
  '  response.redirect dato
'''''''
  end if
end if

set errores=cn.errors
if errores.count = 0 then
'  response.redirect "../vuelve.asp"
else
'  response.write ofv.ErrorDB(cn,"../vuelve.asp")
end if

call ofv.cerrarconn(cn)

%>