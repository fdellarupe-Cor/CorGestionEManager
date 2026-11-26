<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  SOPCION              = request.querystring("OPCION")
  session("MODORIGEN") = "TIPOPERFIL"

  if sopcion = "VUELVE" then
     CantRegMoverx    = session("Regisvuel")
     SPERFIL          = session("Regisperfil")

     response.redirect "tiposdeperfil.asp?CantRegMover=" & CantRegMoverx

  else 
     CantRegMoverx          = cdbl(request.querystring("CantRegMover"))
     SPERFIL                = request.querystring("PERFIL")
     session("Regisvuel")   = CantRegMoverx
     session("Regisperfil") = SPERFIL

     if ID = "" then
        ID = session("elID")
     else
        session("elID") =cdbl(ID)
     end if  

     scodigo = 0

     set cn = ofv.conectar(ofv.strconn2)

     Strsql = "Select Ecoperfil.*,Tipoperfil.descri From Ecoperfil, Tipoperfil where "
     Strsql = Strsql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and codsector=0 and Ecoperfil.codperfil = '" & SPERFIL & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if Not rs.EOF then

        scodigo = rs(0) 
   
     else

        strsql = "Insert into Ecoperfil "
        strsql = strsql & " (CODPERFIL,CODSECTOR,DOTACION)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SPERFIL & "','0','0'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        Strsql = "Select Ecoperfil.*,Tipoperfil.descri From Ecoperfil, Tipoperfil where "
        Strsql = Strsql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and codsector=0 and Ecoperfil.codperfil = '" & SPERFIL & "' "
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