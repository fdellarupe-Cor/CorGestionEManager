<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

  set cn = ofv.conectar(ofv.strconn6)

  sopcion = request.querystring("ELIMINAR")
  sCODIGO = request.querystring("ID")
  sNOMBRE = request.querystring("NOMBRE")
  sVLV    = request.querystring("VLV")
  CantRegMoverx = request.querystring("CantRegMover")

  IF CantRegMoverx = "" THEN CantRegMoverx = request.querystring("Campo")

  STIPO =  SESSION("STIPO")

 IF sVLV  = "" THEN
    sVLV = "N"
 END IF


 LISTA = ""
 DESTINO = ""

 IF STIPO  = "PRY" THEN
    STABLA = "PROYECTO"

 ELSE
    STABLA = "REQUERIMIENTO"

 END IF




  IF SOPCION = "" THEN
      sMensaje = "Esta seguro que desea eliminar " & sNombre
      dato = "Mensajes4.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="

      dato = dato & "PRYREQDELETE.asp?VLV=" & sVLV & "&ID=" & sCODIGO
      dato = dato & "&Botones=2" & "&Tabla=" & STABLA & "&Nombre=" & ofv.convertircar(sNombre," ","_")
      dato = dato & "&Campo=" & CantRegMoverx & chr(13)
      response.redirect dato

  ELSEIF SOPCION = "SI" THEN



   ''     STRSQL = "DELETE FROM " & STABLA & " WHERE CODIGO = '" & SCODIGO & "' "  
   ''     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

     IF STIPO  = "PRY" THEN

        STRSQL = "DELETE FROM PROYANTEC WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYECTO WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYETAPA WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYPLAN WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYPRTY WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYRESP WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYSKR WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYTAREAS WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYTAREASADV WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM PROYTAREASLST WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

     ELSE

        STRSQL = "DELETE FROM REQANTEC WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQAUTOR WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQCOMPO WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQCONTACTOS WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQDELEGA WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 
        STRSQL = "DELETE FROM REQINFO WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQPERFIL WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQRESPO WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQSECTORES WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        STRSQL = "DELETE FROM REQUERIMIENTO WHERE CODIGO = '" & SCODIGO & "' "  
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

    END IF

 

    if sVLV = "N" then
       RESPONSE.REDIRECT "PROYLISTDSP.ASP?CantRegMover=" & CantRegMoverX
    else
       RESPONSE.REDIRECT "pryadmincidenciasdet.ASP?CantRegMover=" & CantRegMoverX
    end if

  ELSE

    if sVLV = "N" then
       RESPONSE.REDIRECT "PROYLISTDSP.ASP?CantRegMover=" & CantRegMoverX
    else
       RESPONSE.REDIRECT "pryadmincidenciasdet.ASP?CantRegMover=" & CantRegMoverX
    end if

  END IF 

END IF

RESPONSE.WRITE SHTML


%>