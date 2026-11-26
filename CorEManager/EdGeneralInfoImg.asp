<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sOPCION = request.querystring("OPCION")
  sfamily = request.querystring("family")
  IDX=REQUEST.QUERYSTRING("IDX")
'  response.write REQUEST.QUERYSTRING
  session("idx")=""
  if sfamily = "" then sfamily = "D"
    vuelta  = ofv.vueltaasp2("../") & "?family=" & sfamily
  '   vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn2)
  if nCodigo&""  <> "" then
  session("nxCodigo")=nCodigo
  else
   ncodigo=session("nxCodigo")
  end if

  if sTipo&""  <> "" then
    session("nxsTipo")=sTipo
    else
     sTipo=session("nxsTipo")
  end if


      session("comp31") = "TODOS"
      session("comp32") = "TODOS"


  if sopcion = "SI" then

        strsql = "Insert into objinfoIMG ("
        strsql = strsql & "CODOBJETO,TIPOBJETO,OBJSECUEN,OBJTIPO,OBJDIRECTORIO,OBJIMAGEN"
        strsql = strsql & ") "
        strsql = strsql & " values ("
        strsql = strsql & "'" & ncodigo & "','" & stipo & "','10','I','',''"
        strsql = strsql & ")"

      '  RESPONSE.WRITE STRSQL & "<BR>"

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  end if


  if sopcion = "UPD" then

      ''    RESPONSE.WRITE "imagen>" & request.form(1)

        '  strsql = "Insert into objinfoIMG ("
        '  strsql = strsql & "CODOBJETO,TIPOBJETO,OBJSECUEN,OBJTIPO,OBJDIRECTORIO,OBJIMAGEN"
        '  strsql = strsql & ") "
        '  strsql = strsql & " values ("
        '  strsql = strsql & "'" & ncodigo & "','" & stipo & "','10','E','',''"
        '  strsql = strsql & ")"

          STRSQL = "UPDATE  objinfoIMG SET OBJIMAGEN='" & request.form("OBJIMAGEN")&"'"
         ' STRSQL=STRSQL & ",CODOBJETO=" & ncodigo & ",TIPOBJETO='" & stipo & "'"
          STRSQL =STRSQL & " WHERE OBJID=" &IDX
     '     RESPONSE.WRITE STRSQL
          call ofv.crearconsultaEx(StrSql,cn,1,parametros)


  end if




  sHTML = ofv.FormHeader(session("FormN"))


'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


  'Encabezados
  Fila = ofv.Encabezados("#","tipo","directorio","imagen")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  Strsql = "Select * From objinfoIMG where CODOBJETO = 0" & ncodigo & " and tipobjeto = '" & stipo & "'  Order By OBJsecuen, objid"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)
    Formulario = ""
    sForm = "form" & posi

     clase = "ut "

    'Eliminar
   if bSuprimir then
     sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
     sHref = sHref & "&Tabla=OBJINFOIMG&volver=" & vuelta & "&Eliminar=NSNC"
     ' sHref    = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
     Celda = ofv.BotonEliminar("imagen: " & cstr(rs(0)),sHref)
  else
    Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
  end if

    'secuen
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.GenCelda("",clase,"","1%","","","",sInput,"si")

    'tipo
    scomK = "I" & chr(9) & "E"
    scomD = "EManager" & chr(9) & "Externa"


    sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD,"si")
    Celda = Celda & ofv.GenCelda("",clase,"","15%","","","",sInput,"si")


    sInputD = ofv.GenerarInput(rs(5).name,"","hidden","","10","","") 'ofv.GenerarInput(rs(5).name,"N","hidden",sForm,"40","dt","")
    'Celda = Celda & ofv.GenCelda("",clase,"","25%","","","",sInput,"si")

    if rs(4)="E" then
    sInput = ofv.GenerarInput(rs(6).name,rs(6)&"","text",sForm,"70","dt","")
    else
    sinput=rs(6)
    end if

    Celda = Celda & ofv.GenCelda("",clase,"","45%","","","",sInput & sInputD,"si")

     if rs(4)="I" then
      sInput  = "Seleccionar"
	 ' Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")

	  sHref    = "EdimagenCMPex.asp?XP=N&OPC=NO&codobjeto=" & nCodigo & "&tipobjeto=" & stipo & "&family=" & sfamily & "&idx=" & rs(0)
	  sImagenP = "&#0232;"
	  sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
	  sinput   = "<a class='ft1n' " & sFuncAnc & " href='" & sHref
	  sinput   = sinput & "' title='Seleccionar imagen' >" & sImagenP & "</a>"

      Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
     ELSE
      Celda   = Celda & ofv.GenCelda("","uti","","80%","","","","-----","si")
     END IF
      Fila = Fila & ofv.genrow("","","","","30","","",Celda,"si")





    sAccion = ""
    if bModificar then

    '  sInput1 = "&nbsp;"
    '  sInput2 = ofv.GenerarInput("","Actualizar","submit","","10","btsb","")

     ' Celda = ""

     ' Celda = celdax & ofv.GenCelda("","ut colspan=5 style='border-top-style:solid;border-top-color:#888888;border-top-width:1;' ","right","","","","",sInput2,"si")

     ' Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJINFOIMG&volver=" & vuelta
      sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi=posi+1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","0","0","0","","",Formulario,"si")

  Agregar = ""
  if bAgregar and posi = 0 then
    sAccion = "EDGENERALINFOIMG.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "&opcion=SI" & "&family=" & sfamily
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  buscar = ""

  sAccion = "EDGENERALINFO.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "&family=" & sfamily
  Celda = ofv.BotonVolver(sAccion)



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>