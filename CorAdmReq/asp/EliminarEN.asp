<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0



vuelta = request.querystring("volver")
nCodigo = request.querystring("Codobjeto")
sID = request.querystring("ID")
Tabla = request.querystring("Tabla")
sEliminar = request.querystring("Eliminar")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")

if sEliminar = "NSNC" then session("volver2") = volver
sMsg = ""

on error resume next

set cn = ofv.conectar(ofv.strconn6)

if sEliminar <> "NO" then
  if sMsg = "" then
    if sEliminar = "NSNC" then
      sMensaje = "Esta seguro que desea eliminar " & sObjeto
      dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
      dato = dato & "asp/EliminarEN.asp?ID=" & sID & "&Tipobjeto=" & sTipo & "&Codobjeto="
      dato = dato & nCodigo & "&Accion=" & nAccion & "&Codproced=" & nProced & "&vuelta="
      dato = dato & vuelta & "&Botones=2" & "&Tabla=" & Tabla & chr(13)
      response.redirect dato
    elseif sEliminar = "SI" then
           

      strsql= "Delete From " & Tabla & " Where objid = " & sID
    
      call ofv.crearconsultaEx(StrSql,cn,1,volver)
    

      Volver = session("volver2")
      
         if right(volver,3) = "asp" then
            volver = volver & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         else
            volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         end if
     

     set errores = cn.errors
     if errores.count = 0 then
        response.redirect volver
     else
        response.write ofv.ErrorDB(cn,volver)
     end if 

    end if
  else
    dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
    dato = dato & "asp/EliminarEN.asp?ID=" & sID & "&Tipobjeto=" & sTipo & "&Codobjeto="
    dato = dato & nCodigo & "&Accion=" & nAccion & "&Codproced=" & nProced & "&vuelta="
    dato = dato & vuelta & "&Botones=1" & "&Tabla=" & Tabla & chr(13)
    response.redirect dato
  end if

else


      Volver = session("volver2")
      if tabla = "Objaccioncontrol" then
         Volver = Volver & "?Codproced=" & nProced & "&Codobjeto=" & nCodigo & "&Accion="
         Volver = Volver & nAccion & "&Tipobjeto=" & sTipo & "&vuelta=" & vuelta
      else
         if right(volver,3) = "asp" then
            volver = volver & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         else
            volver = volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         end if
      end if

       response.redirect volver

end if



call ofv.cerrarconsulta(rs)
call ofv.cerrarconn(cn)

%>