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
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")


    sImagenf = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=2>4</FONT></B>"
    sImageng = "<B><FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=4>&#0232;</FONT></B>"

  celdas = ""
  filas  = ""

  'Encabezados
  celdas = ofv.GenCelda("","TTT colspan=5  ","center","","","","","SOLICITAR NUEVO REQUERIMIENTO","si")
  filas  = FILAS & ofv.GenRow("","","","","22","","",celdas,"si")

  celdas = ofv.GenCelda("","TTR  ","","2%","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","TTR  ","","","","","","CODIGO","si")
  celdas = celdas & ofv.GenCelda("","TTR  ","","","","","","DESCRIPCION","si")
  celdas = celdas & ofv.GenCelda("","TTR  ","","","","","","DESTINATARIO","si")
  celdas = celdas & ofv.GenCelda("","TTR  ","","","","","","GENERAR","si")
  filas  = FILAS & ofv.GenRow("","","","","22","","",celdas,"si")


       celdas = ofv.GenCelda("","aaa colspan=5 ","center","5%","","","","&nbsp;","si")


       filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

       sc = "UT"

  StrSql = "Select * From TIPOREQnv "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rs(0)


       sc = "UT"



       celdas = ofv.GenCelda("","UTS","","1","","","",sImagenf,"si")

       sInput = rs(1)
       celdas = celdas & ofv.GenCelda("",sc,"","5%","","","",sInput,"si")

       sInput = rs(2)
       celdas = celdas & ofv.GenCelda("",sc,"","40%","","","",sInput,"si")


       sInput = getsect(rs(11))
       celdas = celdas & ofv.GenCelda("",sc,"","40%","","","",sInput,"si")

       sinput = "<a CLASS=NT href=PROYLISTDSP.ASP?OPC=NO&TIPO=REQ&RESP=O&OPCION=OK&TREQ=" & RS(1)
       sinput = sinput  & " >" & sImageng & "</a>" 

       celdas = celdas & ofv.GenCelda("","UTS","center","5%","","","",sInput,"si")


       filas = FILAS & ofv.GenRow("","","","","20","","",celdas,"si")

 

    posi = posi + 1
    rs.Movenext
  loop

       celdas = ofv.GenCelda("","aaa colspan=5 ","center","5%","","","","&nbsp;","si")


       filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & "<CENTER>"
  sHTML = sHTML  & ofv.GenTabla("","aaa style='border-style:SOLID;border-width:1;' ","","70%","","0","","-2","-1","","",filas,"si")
  sHTML = sHTML  & "</CENTER>"

  'Botonera


  Agregar = ""
  Buscar  = ""
  Celda   = ""


     sAccion = "PROYLISTDSP.ASP?OPC=NO&TIPO=REQ&RESP=O"
     Celda = ofv.BotonVolver(sAccion)
 

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,Buscar,celda)






  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if


response.write sHTML

fuNction getsect(XSECTORIG)

        StrSql = "SELECT * FROM ECO where ECO.eco = " & XSECTORIG
        set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             getsect = "SIN ASIGNAR"
          Else
             getsect = rsX3(1)
          End If
          call ofv.cerrarconsulta(rsX3)

end function



%>