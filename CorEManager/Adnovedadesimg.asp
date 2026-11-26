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
  session("idx")=""
  if CantRegMoverx = "" then  CantRegMoverx = request.querystring("Cantregmover")

'response.write  request.querystring
  ID = request.querystring("ID")
  IDX=REQUEST.QUERYSTRING("IDX")
  seditar=REQUEST.QUERYSTRING("editar")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID & "_editar=" & seditar

  set cn = ofv.conectar(ofv.strconn2)

  if sopcion = "SI" then

        strsql = "Insert into NOVEGENPARIMG ("
        strsql = strsql & "NOVEGENPAR,SECUEN,TIPO,DIRECTORIO,IMAGEN"
        strsql = strsql & ") "
        strsql = strsql & " values ("
        strsql = strsql & "'" & ID & "','10','I','',''"
        strsql = strsql & ")"

      '  RESPONSE.WRITE STRSQL & "<BR>"

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  end if

  if sopcion = "UPD" then
  strsql = "UPDATE  NOVEGENPARIMG SET IMAGEN='" & REQUEST.FORM("IMAGEN") & "' WHERE OBJID=" & IDX
  call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF



  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Fila = ofv.Encabezados("#","tipo","directorio","imagen")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From novegenparimg where novegenpar = '" & id & "' Order By secuen, objid"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)
    Formulario = ""
    sForm = "form" & posi

     clase = "ut "

    'Eliminar
   if bSuprimir then
     sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
     sHref = sHref & "&Tabla=novegenparimg&volver=" & vuelta & "&Eliminar=NSNC"
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


    'sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"40","dt","")
    'Celda = Celda & ofv.GenCelda("",clase,"","25%","","","",sInput,"si")
    if rs(3)="E" then
	    sInput = ofv.GenerarInput(rs(5).name,rs(5)&"","text",sForm,"70","dt","")
	    else
	    sinput=rs(5)
    end if



    'sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"70","dt","")
    Celda = Celda & ofv.GenCelda("",clase,"","45%","","","",sInput,"si")
    if rs(3)="I" then
	      sInput  = "Seleccionar"
		 ' Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")

		  sHref    = "EdimagenCMPExNOVED.asp?XP=N&OPC=NO&codobjeto=" & id & "&IDX=" &RS(0) & "&editar=" & seditar
		  'sImagenP = "&#0232;"
		  'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
		  sinput   = "<a " & sFuncAnc & " href='" & sHref
		  sinput   = sinput & "' title='Seleccionar imagen' ><i class='las la-arrow-right'></i></a>"

	      Celda   = Celda & ofv.GenCelda("","uti","center","80%","","","",sInput,"si")
	     ELSE
	      Celda   = Celda & ofv.GenCelda("","uti","","80%","","","","-----","si")
     END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=novegenparimg&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

    posi=posi+1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","aaa","","100%","","0","0","0","0","","",Fila,"si")

  Agregar = ""
  if bAgregar then
    sAccion = "adnovedadesimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id & "&opcion=SI" & "&editar=" & seditar
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  buscar = ""

  sAccion = "AdnovedadesProp.asp?Retorno=" & CantRegMoverx & "&ID=" & id & "&editar=" & seditar
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>