<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
''  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")


   celdas = ""



  'Encabezados
  formulario = formulario & ofv.GenCelda("","ttt colspan=4","center","","","","","PERFILES DE REQUERIMIENTOS","si")
  formulario = formulario & ofv.Encabezados("CODIGO","DESCRIPCION","&nbsp;")


  StrSql = "Select * From TIPOREQ "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "ut"


      Celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
 


       sInput = rs(1)
       celdas = celdas & ofv.GenCelda("",sc,"","15%","","","",sInput,"si")

       sInput = rs(2)
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")




       sinput = "<a  href=CORREQTIPOPROP.asp?PERFIL=" & RS(0)
       sinput = sinput  & " ><b>TIPOS&nbsp;ASOCIADOS</b></a>" 

       celdas = celdas & ofv.GenCelda("","TT","center","20%","","","",sInput,"si")


    filas = ofv.GenRow("","","","","","","",celdas,"si")


      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")


    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""


  Buscar = ""
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML





%>