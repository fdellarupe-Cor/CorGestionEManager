<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

nID = request.querystring("ID")
t = request.querystring("t")
Nombre = request.querystring("Nombre")
sOpcion = request.querystring("Opcion")
sAplicacion = request.querystring("Aplicacion")
menu = request.querystring("menu")
submenu = request.querystring("submenu")
sGrupoUsuario = request.querystring("grupousuario")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")

'response.write "gu " & sGrupoUsuario & " ID " & nid

xdefault = request.form("default")
xtipo = "E"

dim aatrm(9)
for xind=0 to 9
    aatrm(xind) = 1
next
xind = 0
xatrm = nid
xtab = "S"

set cn = ofv.conectar(ofv.strconn0)

strsql = "select atributo from atrm where id = " & nID
set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rsa.eof then xatrm = rsa(0)
call ofv.cerrarconsulta(rsa)



if t = "G" then
  stabla = "atrmGrupos"
  sCampo = "Grupo"
else
  stabla = "atrmUsuarios"
  sCampo = "Usuario"
end if

strsql = "SELECT * From " & stabla & " Where " & sCampo & " = '" & sGrupoUsuario & "' "
strsql = strsql & "And atributo = '" & xatrm & "'"
'response.write strsql
set rsa = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if rsa.eof then xtab = "N"
call ofv.cerrarconsulta(rsa)

xtipo = "A"
for each x in request.form
       if ucase(x) = "ESTADO" then
          xtipo = "E"
          EXIT FOR
       end if
next


'response.write stabla
select case xtipo
       case "E"
       aatrm(0) = request.form("ESTADO")
       aatrm(1) = request.form("ESTADOD")
     '  response.write "0: " & aatrm(0) & " 1: " & aatrm(1) & " a " &  xatrm & "<br>"
       if aatrm(0) = aatrm(1) then
          if xtab <> "N" then

             strsql = "SELECT * FROM " & stabla & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql  & "' And atributo = '" & xatrm & "'"
			 SET RSL= ofv.crearconsultaEx(StrSql,cn,1,volver)
			 xxxr=GENLOGEXTDATA("DELETE",stabla,RSL(0).NAME,RSL(0),"SKR")
             CALL OFV.CERRARCONSULTA(RSL)

         strsql = "delete From " & stabla & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql  & "' And atributo = '" & xatrm & "'"
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if
       else
'       response.write "0: " & aatrm(0) & " 1: " & aatrm(1) & " a " &  xatrm & "<br>"
          if xtab = "N" then
             strsql = "insert into " & stabla & " (" & sCampo & ",atributo,estado,agregar,"
             strsql = strsql  & "modificar,eliminar,consultar) values('"
             strsql = strsql  & sGrupoUsuario & "','" & xatrm & "'," & aatrm(0) & ",1,1,1,1)"
             'call ofv.crearconsultaEx(StrSql,cn,1,volver)

             STRSQL="PRC_EXECSQL '" & REPLACE(STRSQL,"'","''") & "'"
			 SET VV= ofv.crearconsultaEx(StrSql,cn,1,volver)
			 IF NOT ISNULL(VV(0)) THEN
			   IDINSERT=VV(0)
			 END IF
			 CALL OFV.CERRARCONSULTA(VV)
             xxxr=GENLOGEXTDATA("INSERT",stabla,"ID",IDINSERT,"SKR")

          else

             strsql = "SELECT * FROM " & stabla
			 strsql = strsql & " Where " & sCampo & " = '" & sGrupoUsuario
			 strsql = strsql & "' And atributo = '" & xatrm & "' "

			 SET RSL= ofv.crearconsultaEx(StrSql,cn,1,volver)

			 XRS=CSTR(RSL("ESTADO")&"")
             vNuevoValor= CSTR(aatrm(0)&"")

             strsql = "update " & stabla & " set estado = " & aatrm(0)
             strsql = strsql & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql & "' And atributo = '" & xatrm & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
             if vNuevoValor <> XRS then
			           XVALANT=	CBOOL(XRS)
			           XVALACT	=CBOOL(vNuevoValor)
			           xxxr=GENLOGEXTDATA("UPDATE",STABLA,"estado",RSL(0),"SKR")
             end if
            ' xxxr=GENLOGEXTDATA("UPDATE",stabla,"estado",RSL(0),"SKR")

             CALL OFV.CERRARCONSULTA(RSL)
          end if
        ' response.write "sql: " & strsql & "<br>"
       end if

       case "A"
       aatrm(0) = request.form("AGREGAR")
       aatrm(1) = request.form("AGREGARD")
       aatrm(2) = request.form("MODIFICAR")
       aatrm(3) = request.form("MODIFICARD")
       aatrm(4) = request.form("ELIMINAR")
       aatrm(5) = request.form("ELIMINARD")
       aatrm(6) = request.form("CONSULTAR")
       aatrm(7) = request.form("CONSULTARD")

       if aatrm(0) = aatrm(1) and aatrm(2) = aatrm(3) and aatrm(4) = aatrm(5) and aatrm(6) = aatrm(7) then
          if xtab <> "N" then

             strsql = "SELECT *  From " & stabla & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql  & "' And atributo = '" & xatrm & "'"
			 SET RSL= ofv.crearconsultaEx(StrSql,cn,1,volver)
			 xxxr=GENLOGEXTDATA("DELETE",stabla,RSL(0).NAME,RSL(0),"SKR")
             CALL OFV.CERRARCONSULTA(RSL)

             strsql = "delete From " & stabla & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql  & "' And atributo = '" & xatrm & "'"
             call ofv.crearconsultaEx(StrSql,cn,1,volver)
          end if
       else
          if xtab = "N" then
             strsql = "insert into " & stabla & " (" & sCampo & ",atributo,estado,agregar,"
             strsql = strsql  & "modificar,eliminar,consultar) values('"
             strsql = strsql  & sGrupoUsuario & "','" & xatrm & "',1,"
             strsql = strsql  & aatrm(0) & "," & aatrm(2) & "," & aatrm(4) & "," & aatrm(6) & ")"
             'call ofv.crearconsultaEx(StrSql,cn,1,volver)

             STRSQL="PRC_EXECSQL '" & REPLACE(STRSQL,"'","''") & "'"
			 SET VV= ofv.crearconsultaEx(StrSql,cn,1,volver)
			 IF NOT ISNULL(VV(0)) THEN
			   IDINSERT=VV(0)
			 END IF
			 CALL OFV.CERRARCONSULTA(VV)
             xxxr=GENLOGEXTDATA("INSERT",stabla,"ID",IDINSERT,"SKR")

          else


             strsql = "SELECT * FROM  " & stabla
			 strsql = strsql & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql & "' And atributo = '" & xatrm & "' "
			 SET RSL= ofv.crearconsultaEx(StrSql,cn,1,volver)
			 IDL=RSL(0)
             CALL OFV.CERRARCONSULTA(RSL)

             strsql = "update " & stabla & " set estado = 1,"
             strsql = strsql & " agregar = " & aatrm(0) & ","
             strsql = strsql & " modificar = " & aatrm(2) & ","
             strsql = strsql & " eliminar = " & aatrm(4) & ","
             strsql = strsql & " consultar = " & aatrm(6)
             strsql = strsql & " Where " & sCampo & " = '" & sGrupoUsuario
             strsql = strsql & "' And atributo = '" & xatrm & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,volver)

             XVALANT=	"ESTADO"
			 XVALACT=CBOOL("1")
             xxxr=GENLOGEXTDATA("UPDATE",stabla,"ID",IDL,"SKR")
             XVALANT=	"agregar"
			 XVALACT=CBOOL(aatrm(0))
             xxxr=GENLOGEXTDATA("UPDATE",stabla,"ID",IDL,"SKR")
             XVALANT=	"modificar"
			 XVALACT=CBOOL(aatrm(2))
             xxxr=GENLOGEXTDATA("UPDATE",stabla,"ID",IDL,"SKR")
             XVALANT=	"consultar"
			 XVALACT=CBOOL(aatrm(6))
             xxxr=GENLOGEXTDATA("UPDATE",stabla,"ID",IDL,"SKR")
             XVALANT=	"ELIMINAR"
			 XVALACT=CBOOL(aatrm(4))
             xxxr=GENLOGEXTDATA("UPDATE",stabla,"ID",IDL,"SKR")
          end if
       end if


end select



if cn.errors.count = 0 then
  volver = volver & "&Opcion=" & sOpcion & "&Aplicacion=" & sAplicacion & "&menu=" & menu
  volver = volver & "&submenu=" & submenu & "&t=" & t & "&Nombre=" & Nombre
  response.redirect volver
else
  response.write ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconn(cn)

%>