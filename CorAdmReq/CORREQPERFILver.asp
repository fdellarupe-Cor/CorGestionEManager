<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = request.querystring("CantRegMover")


  wperfil = request.querystring("perfil")


  IF wperfil <> "" THEN 
     SESSION("wperfil") = wperfil
  ELSE
     wperfil = SESSION("wperfil")
  END IF


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_perfil=" & Wperfil

  sHTML = ofv.MenuHeader("","")




  'Encabezados
  Formulario = ""

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>L</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=4>&#064;</FONT></B>"



  dim stabla(12)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)






  posi = 0

  Strsql = "Select * From TIPOREQ where ID = " & wperfil
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     fila = ""


      Celda = ofv.GenCelda("","vt style='border-BOTTOM-style:ridge;border-BOTTOM-width:2;border-BOTTOM-color:#f7aa22;' ","","","","","","PERFIL: ","si")
      sInput = rs(1) & "&nbsp;-&nbsp;" & rs(2)

      Celda = celda & ofv.GenCelda("","vt colspan=2 style='border-BOTTOM-style:ridge;border-BOTTOM-width:2;border-BOTTOM-color:#f7aa22;' ","","90%","","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")



      Formulario = Fila

      sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si")



      fila = ""

      Celda = ofv.GenCelda("","vt colspan=2 ","","","","","","USUARIOS","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
      Formulario = Fila

    StrSql = "Select id,Descripcion From Usuarios where sector = 0" & RS(34) & " AND  NIVEL <= 0" & RS(35) & " Order By descripcion"
    set rsx1 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
    if rsx1.eof then

       sLink = "<b><font size=2>NO EXITEN USUARIOS PARA LA SELECCION<font></b>"
       celdas = ofv.GenCelda("","vt VALIGN=MIDDLE ","CENTER","80%","","","",sLink,"si")
       Formulario =  Formulario & ofv.GenRow("","","","","40","","",celdas,"si")

    else

  Do While (Not rsx1.EOF)



       sLink = "<b><font size=2>" & rsx1(0) & "&nbsp;-&nbsp;" & rsx1(1) & "<font></b>"
       celdas = ofv.GenCelda("","vt","","80%","","","",sLink,"si")
       Formulario =  Formulario & ofv.GenRow("","","","","10","","",celdas,"si")


  rsx1.movenext
  loop
  end if

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#dddddd,direction=125);' ","","80%","","0","","0","0","","",Formulario,"si")





  end if




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML


%>