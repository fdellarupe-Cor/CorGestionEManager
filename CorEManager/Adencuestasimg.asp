<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMoverx = request.querystring("retorno")
  sopcion = request.querystring("opcion")

  if CantRegMoverx = "" then  CantRegMoverx = request.querystring("Cantregmover")


  ID = request.querystring("ID")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID

  set cn = ofv.conectar(ofv.strconn2)

  if sopcion = "SI" then

        strsql = "Insert into emencuestasIMG ("
        strsql = strsql & "ENCcodigo,IMGSECUEN,IMGTIPO,IMGDIRECTORIO,IMGIMAGEN"
        strsql = strsql & ") "
        strsql = strsql & " values ("
        strsql = strsql & "'" & ID & "','10','E','',''"
        strsql = strsql & ")"

      '  RESPONSE.WRITE STRSQL & "<BR>"

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  end if




  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Fila = ofv.Encabezados("#","tipo","directorio","imagen")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From emencuestasimg where ENCcodigo = '" & id & "' Order By IMGsecuen, objid"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)
    Formulario = ""
    sForm = "form" & posi

     clase = "ut "

    'Eliminar
   if bSuprimir then
     sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_") 
     sHref = sHref & "&Tabla=emencuestasimg&volver=" & vuelta & "&Eliminar=NSNC" 
     Celda = ofv.BotonEliminar("imagen: " & cstr(rs(0)),sHref) 
  else 
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si") 
  end if

    'secuen
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.GenCelda("",clase,"","1%","","","",sInput,"si")

    'tipo
    scomK = "I" & chr(9) & "E"
    scomD = "EManager" & chr(9) & "Externa"


    sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
    Celda = Celda & ofv.GenCelda("",clase,"","15%","","","",sInput,"si")


    sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"40","dt","")
    Celda = Celda & ofv.GenCelda("",clase,"","25%","","","",sInput,"si")


    sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"70","dt","")
    Celda = Celda & ofv.GenCelda("",clase,"","45%","","","",sInput,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=emencuestasimg&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    posi=posi+1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","0","0","0","","",Fila,"si")

  Agregar = ""
  if bAgregar then
    sAccion = "adencuestasimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id & "&opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  buscar = ""

  sAccion = "AdencuestasProp.asp?Retorno=" & CantRegMoverx & "&ID=" & id
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>