<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader("Autorizantes Default")

  sHTML = sHTML & "<center>"

  'Encabezados


  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From DefWfl"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rs.eof then

     strsql = "Insert into DefWFL (" & rs(1).name & "," & rs(2).name & "," & rs(3).name & ") "
     strsql = strsql & " Values (0,0,0)"

     call ofv.crearconsultaEx(StrSql,cn,1,volver)
     Strsql = "Select * From DefWfl"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if


  if Not rs.EOF then
    sForm = "form" & posi

    fila = ""

    Celda =  ofv.gencelda("","ttt colspan=2 ","","","","","",ucase("Autorizantes default: "),"si")
    Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")


    'Activa
    sInput = ofv.combosino("form" & posi,rs(1),rs(1).name)

    Celda =  ofv.gencelda("","ut","","","","","",ucase("Activado: "),"si")
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    Fila = fila & ofv.genrow("","","","","22","","",Celda,"si")

    'Orden estricto
    sInput = ofv.combosino("form" & posi,rs(2),rs(2).name)

    Celda = ofv.gencelda("","ut","","","","","",ucase("Orden estricto: "),"si")
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    Fila = fila & ofv.genrow("","","","","22","","",Celda,"si")


    if rs(1) then

    Celda = ofv.gencelda("","ut","","","","","",ucase("Lista: "),"si")

    'sImagenP = "&#0232;"
    sinput = "<a " & sFuncAnc & " href='AutDefaultDet.asp' "
    sinput = sinput  & " title='Lista de autorizantes' ><i class='las la-arrow-right' ></i></a>"
    Celda = Celda & ofv.gencelda("","ft","","","15%","","",sInput,"si")


 '      sinput = "<a  href=AutDefaultDet.asp"
 '      sinput = sinput  &  " ><b>Lista de Autorizantes</b></a>"
 '      Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
 '   else
 '      Celda = Celda & ofv.gencelda("","","","","","","","&nbsp;","si")
 '   end if

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    end if

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & rs(0)
      sAccion = sAccion & "&tabla=DefWfl&volver=" & vuelta
    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

  end if
  sHTML = sHTML & ofv.GenTabla("","aaa","","60%","","0","","0","1","","",Formulario,"si")

  sHTML = sHTML & "</center>"


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>