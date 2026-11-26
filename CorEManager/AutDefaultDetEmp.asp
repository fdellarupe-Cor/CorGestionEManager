<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
'response.write request.querystring
  vuelta = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")

  set cn = ofv.conectar(ofv.strconn2)
  set cn3 = ofv.conectar(ofv.strconn6)

  sHTML = ofv.FormHeader("Empresa")

ID = request.querystring("ID")
opcion = request.querystring("opcion")

IF session("EMPMUN") = "S" then

strsql = "select a.site from usuarios a inner join objwfldet b on (a.ID=b.usr) where b.objid=" & ID
set rv = ofv.crearconsultaex(strsql,cn,1,parametros)
site = ""

if not rv.eof then site = rv(0)

sql = "select * from objmempalc where codobjeto='0' and tipobjeto='0' and codempre='" & site & "' and idusuario='" & ID & "'"
set rs = ofv.crearconsultaex(sql,cn,1,parametros)

if rs.eof then

sql = "insert into OBJmempalc (codobjeto,tipobjeto,codempre,idusuario) values "
sql = sql & " ('0','0','" & site & "','" & ID & "')"
call ofv.crearconsultaex(sql,cn,1,parametros)

end if

call ofv.cerrarconsulta(rs)
call ofv.cerrarconsulta(rv)

end if

if opcion = "agre" then

idemp = request.querystring("idemp")
idus = request.querystring("idus")

'SI VIENE EN 1 INSERT SI 0 DELETE
aod = request.form(idemp)

if aod = 1 then

sql = "insert into OBJmempalc (codobjeto,tipobjeto,codempre,idusuario) values "
sql = sql & " ('0','0','" & idemp & "','" & idus & "')"
call ofv.crearconsultaex(sql,cn,1,parametros)

else

objid = request.querystring("objid")
'DELETE
sql = "delete from OBJmempalc where objid=" & objid
call ofv.crearconsultaex(sql,cn,1,parametros)

end if

end if

	Celda = ofv.gencelda("","TT","","20%","","","","EMPRESA","si")
	Celda = Celda & ofv.gencelda("","TT","","5%","","","","ALCANCE","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")

	sql = "select b.id,b.cliente,a.codobjeto,a.tipobjeto,a.objid from clienteplan b "
	sql = sql & " left join OBJmempalc a on (b.id=a.codempre and a.codobjeto='0' and a.tipobjeto='0' and a.idusuario=" & ID & ") "
	sql = sql & " order by b.cliente "
	'response.write sql & "<br>"
	set rs = ofv.crearconsultaex(sql,cn,1,parametros)

	if not rs.eof then

	do while not rs.eof

	sForm = "Fx" & rs(0)

	Celda = ofv.gencelda("","ttt","","20%","","","",rs(1),"si")

	if not isnull (rs(2)) then
	xEM = 1
	else
	xEM = 0
	end if

	sInput = ofv.combosino(sForm,xEM,rs(0))

	sAccion = "AutDefaultDetEmp.asp?opcion=agre&idus=" & ID & "&idemp=" & rs(0) & "&ID=" & ID & "&objid=" & rs(4) & ""
	'response.write saccion & "<br>"
	Formulario = ofv.genform(sForm,"",sAccion,"","post","",sInput,"si")

	Celda = Celda & ofv.gencelda("","ttt","center","5%","","","",Formulario,"si")

	Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


	rs.movenext
	loop

	end if

Tabla = ofv.GenTabla("","b border=1","center","25%","","","","","","","",Fila,"si")

sHTML = sHTML & Tabla

sAccion = "AutDefaultDet.asp"
Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,Celda)


  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>