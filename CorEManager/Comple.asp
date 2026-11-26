<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar



  XXXX = NEWEMPRE("NO")

  CantRegMostrarx = 12
  CantRegMoverx = "0" & request.querystring("CantRegMover")
  ''RESPONSE.WRITE "-" & CantRegMoverx & "-"
  IF  CantRegMoverx <> "" THEN 
      CantRegMoverx = cdbl(CantRegMoverx)
  ELSE
      CantRegMoverx = 0
  END IF
  vuelta = ofv.vueltaasp("",CantRegMostrarx,Cantregmoverx)

  sTipo = request.querystring("Tipo")
  sNombre = request.querystring("Nombre")

  if sTipo <> "" then 
     session("stipo") = stipo
     session("nombre") = snombre
  else
     stipo = session("stipo")
     snombre = session("nombre")
  end if 




  sImagenM = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

''''''sHTML = ofv.MENUheader("",FUNCION)

  sHTML = "<html><body onload='window.top.ldfobj();' >"


  sHref = ""
  Fila  = ""

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Tipocomp Where tipo='" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    simage = rs(2)
    sPerfil = rs(4)
  end if
  call ofv.cerrarconsulta(rs)


    Celda = ofv.GenCelda("","vt colspan=3 ","","","","","","<u>LISTA DE " & sNombre & "</u>","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


  Strsql = "Select * From Complementos Where tipocomp='" & sTipo & "' and compun = " & session("un") & " order by descri "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi< CantRegMostrarx)

 '' do until rs.eof
    'Eliminar
    if bSuprimir then
      sHref = "asp/ElComp.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tipo=" & sTipo & "&Eliminar=NSNC&Volver=COMPLE.ASP?CANTREGMOVER=" & CANTREGMOVERX
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    sHref = "EditarComp.asp?Codcomp=" & rs(0) & "&Tipocomp=" & rs(2) & "&Nombre="
    sHref = sHref & sNombre & "&Complemento=" & ofv.convertircar(rs(3)," ","_") & "&Perfil="
    sHref = sHref & sPerfil
    sImagen = ofv.GenImage("","","","","Iconos\" & simage,"0","")
    sLink = ofv.GenLink("","nt",sHref,"_blank","","",sImagen,"si")
    Celda = Celda & ofv.GenCelda("","","","","","","",sLink,"si")

    sLink = ofv.GenLink("","nt",sHref,"_blank","","",rs(3),"si")
    Celda = Celda & ofv.GenCelda("","","","120%","","","",sLink,"si")

    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    posi=posi+1
    rs.movenext
  loop	
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si")

  'Agregar
  Agregar = ""
  if bAgregar then
    sInput = ofv.GenerarInput("","Agregar","submit","","10","bt","")
    Celda = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "Agregar2cess.asp?Tipo=" & sTipo & "&Nombre=" & sNombre & "&Volver=" & vuelta
    Agregar = ofv.GenForm("Agregar","",sAccion,"","post","",Celda,"si")
  end if

  'consulta
'  sinput = "<a  href=ListadoPorComponente.asp?"&request.querystring & " target=_new ><b>Listado</b></a>"
'  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

    sInput = ofv.GenerarInput("","Listado","submit","","10","bt","")
    Celda = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "ListadoPorComponente.asp?tipo=" & stipo & "&nombre=" & snombre
    Celda = ofv.GenForm("lista","",sAccion,"_blank","post","",Celda,"si")
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,celda,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>