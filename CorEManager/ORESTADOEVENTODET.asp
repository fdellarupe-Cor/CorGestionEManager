<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  SNIVEL1  = REQUEST.QUERYSTRING("NIVEL1")


  IF SNIVEL1 <> "" THEN
     SESSION("NIVEL1") = SNIVEL1
  ELSE
     SNIVEL1 = SESSION("NIVEL1")
  END IF



  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ""

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  WKTEVENTO  = "NO DEFINIDO"
  WKSETTABLA = "ORESTADOS"

  Strsql = "Select * From ORESTADOS WHERE OBJID = 0" & SNIVEL1
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(2)


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","ESTADO:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si") & "<BR>"

    scomK = "0"
    scomD = "SIN DEFINIR"

    Strsql = "Select * From ORESTADOS WHERE TIPO > 0 AND OBJID <> " & SNIVEL1
    set rsXX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    DO UNTIL rsXX.eof 

        scomD = scomD  & chr(9) & RSXX(2)
        scomK = scomK  & chr(9) & RSXX(0)

    RSXX.MOVENEXT
    LOOP


    Formulario = ""



  FOR POSI=1 TO 10

    sForm = "form" & posi

    WKFCOLOR = "#EEEEEE"

    IF rs(3 + POSI) > 0 THEN WKFCOLOR = "#99CCDD"


    Celda = ""

    'BOTON
    sInput = "BOTON&nbsp;" & POSI
    Celda = Celda & ofv.gencelda("","VT","","10%","","","",sInput,"si")

    'descripcion


    sInput = ofv.gencomboFX(sForm,rs(3 + POSI).name,rs(3 + POSI),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UTI","","","","","",sInput,"si")

    Fila = ofv.genrow("","AAA BGCOLOR='" & WKFCOLOR & "' ","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


  NEXT


  END IF
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""

  sAccion = "ORESTADOEVENTO.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","","","",VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>