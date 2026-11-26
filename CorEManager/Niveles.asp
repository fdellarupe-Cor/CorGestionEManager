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
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  aplcheck()


  if session("APLMDCLI")="BC" THEN
  Formulario = ofv.Encabezados("Id","Descripcion","Grafica","ESTADO")
  ELSE
   Formulario = ofv.Encabezados("Id","Descripcion","ESTADO")
  END IF


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
'response.write bModificar & "  "   & Session("Form")
  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From NivelJerar Order By codnivel"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    sInput = rs(0)

    'Eliminar
    if bSuprimir and sinput > "0" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Niveljerar&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Nivel
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","cc","right","50","","","",sInput,"si")

    'Descripcion
    if rs("estado")=3 then
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    ELSE
    sInput=RS(1)
    END IF

    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    if session("APLMDCLI")="BC" THEN
    ssino = rs(2) & ""

    if ssino = "" then ssino = "N"

    scomK = "S" & chr(9) & "N"
    scomD = "SI" & chr(9) & "NO"


    if bModificar AND RS("ESTADO")=3 then
        sInput = ofv.gencomboFX(sForm,rs(2).name,ssino,scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
    end if

   '''' sInput = ofv.combosino(sForm,ssino,rs(2).name)
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    END IF

    'Celda = Celda & ofv.gencelda("","","","","","","","","si")

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

     Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Niveljerar&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

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
    sAccion = "AgregarNiveles.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=Niveljerar"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Niveljerar Order By codnivel "    ' orden pagina
  str2 = "Select descri From Niveljerar Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sinput = "<a  href=NivelesCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "NivelesCons.asp?" & request.querystring
''Celda = ofv.BotonConsultar(sAccion)
  Reportes=BtnReportes(16,"","xform",sAccion)

  boton=   ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,"no",Reportes)
  sHTML = sHTML & boton

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>