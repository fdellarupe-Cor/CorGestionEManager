<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

''''ListaPuestosFun.asp

  SOPCION              = request.querystring("OPCION")
  session("MODORIGEN") = "LISTAPUESTOS"

  if sopcion = "VUELVE" then
     CantRegMoverx    = session("Regisvuel")
     Spuesto          = session("Regisperfil")

     response.redirect "LISTADEPUESTOS.asp?CantRegMover=" & CantRegMoverx

  else 
     CantRegMoverx          = cdbl(request.querystring("CantRegMover"))
     Spuesto                = request.querystring("PUESTO")
     session("Regisvuel")   = CantRegMoverx
     session("Regisperfil") = Spuesto

     if ID = "" then
        ID = session("elID")
     else
        session("elID") =cdbl(ID)
     end if  

     scodigo = 0

     set cn = ofv.conectar(ofv.strconn2)

     Strsql = "Select Ecoperfil.*,listadepuestos.descri From Ecoperfil, listadepuestos where "
     Strsql = Strsql & "Ecoperfil.codsector = listadepuestos.objid and codsector= '" & Spuesto & "'  and Ecoperfil.codperfil = '****PUESTO' "
    ' response.write strsql
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if Not rs.EOF then

        scodigo = rs(0) 
   
     else

        strsql = "Insert into Ecoperfil "
        strsql = strsql & " (CODPERFIL,CODSECTOR,DOTACION)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'****PUESTO','" & Spuesto & "','0'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "Select Ecoperfil.*,listadepuestos.descri From Ecoperfil, listadepuestos where "
        Strsql = Strsql & "Ecoperfil.codsector = listadepuestos.objid and codsector= '" & Spuesto & "'  and Ecoperfil.codperfil = '****PUESTO' "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if Not rs.EOF then
           scodigo = rs(0) 
        end if


     end if 

    response.redirect "SECTORPERFILESFUN.asp?PERFIL=" & scodigo

  end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


end if

Response.Write sHTML

%>