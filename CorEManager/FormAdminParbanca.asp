<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("BANCA","ESTADO")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = " Select * From FORMADMbanca  Order By DESCRIPCION "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    sInput = rs(0)
    XRSESTADO=RS("ESTADO")
    'Eliminar
    if bSuprimir and sinput > "0" then

      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=FORMADMbanca&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if





    'Descripcion
    IF XRSESTADO=3 THEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"200","dt","")
    ELSE
    sInput=RS(2)
    END IF


    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    SCOMSEP = chr(9)
	    scomK = "3" & SCOMSEP & "4"
	     scomD = "VIGENTE" & SCOMSEP & "NO&nbsp;VIGENTE"
	     if bModificar then
	     	sInput = ofv.gencomboFX(sForm,"ESTADO",RS("ESTADO"),scomK,scomD,"")
	     ELSE
	        IF  RS("ESTADO")= 3 THEN
	        sInput="VIGENTE"
	        ELSE
	        sInput="NO VIGENTE"
	        END IF
         END IF
     Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'Celda = Celda & ofv.gencelda("","","","","","","","","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar  then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=FORMADMbanca&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarBANCA.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=FORMADMBANCA"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRIPCION From FORMADMBANCA Order By DESCRIPCION "    ' orden pagina
  str2 = "Select DESCRIPCION From FORMADMBANCA Order By DESCRIPCION "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sinput = "<a  href=TipodeFormuCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "bancaCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>