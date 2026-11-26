<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0




  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sOpc = request.querystring("opc")
  vuelta = ofv.vueltaasp2("../") & "?opc=SI"



'  sHTML = ofv.FormHeader(session("FormN"))

'  if sOpc <> "SI" then
'     sfuncion = ""
'  else
'     sfuncion = " onload=" & chr(34) & "onload=document.nulo.submit();" & chr(34)
'  end if

  sHTML = ofv.menuheader("",sfuncion)

  'Encabezados
  Formulario = ""


  dim stabla(12)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0

  Strsql = "Select * From Objnoved Where codobjeto =" & nCodigo & " and tipobjeto ='" & stipo & "' " 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if not rs.eof then 

     fila = ""    


      Celda = ofv.GenCelda("","tt ","","","","","","Informa en Novedades: ","si")
      sInput = ofv.combosino("defa","1","dato")
      sInput = OFV.GenForm("defa", "", "EDobjnovedef.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo  , "_self", "post", "", sInput, "si")

      Celda = celda & ofv.GenCelda("","cc  ","","70%","","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      Celda = ""
      for ii=1 to 2  
          Celda = celda & ofv.GenCelda("","aaa ","","1","","","","&nbsp;","si")
      next

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")   
      Formulario = Fila

    sForm = "form" & posi


     fila = "" 

    esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:8pt;background-color:#f7eedd;} "
    'Encabezado Referencia
    Celda = ofv.GenCelda("","tt colspan=2 ","","100%","","","","Texto de Novedad","si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    'Texto de Referencia
    celda = "<font class=at><textarea name=" & rs(3).name & " rows=25 cols=120 " & esttxt
      if bModificar then
        celda = celda & " onblur=" & sForm & ".submit() "
        celda = celda & " onkeypress=" & chr(34)
        celda = celda & " if (this.value.length > 950) {alert('El texto es demasiado"
        celda = celda & " extenso'); return false;}" & chr(34)
      else
        celda = celda & " readonly"
      end if

    celda = celda & " >" & rs(3) & "</textarea></font>"

    Celda = ofv.GenCelda("","cc colspan=2 ","center","","","","",celda,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","tt  ","","","","","","dias: ","si")
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"60","dt","")
      Celda = celda & ofv.GenCelda("","cc","","","","","",sInput,"si")
     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","tt  ","","","","","","tipo distribucion: ","si")
      sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"60","dt","")
      Celda = celda & ofv.GenCelda("","cc","","","","","",sInput,"si")
     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objnoved&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo



    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    
 

 

  else

     fila = ""    

      Celda = ofv.GenCelda("","tt  ","","","","","","Informa en Novedades: ","si")
      sInput = ofv.combosino("defa","0","dato")
      sInput = OFV.GenForm("defa", "", "EDobjnovedef.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo  , "_self", "post", "", sInput, "si")

      Celda = celda & ofv.GenCelda("","cc ","","70%","","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      Celda = ""
      for ii=1 to 2  
          Celda = celda & ofv.GenCelda("","aaa ","","10%","","","","&nbsp;","si")
      next

      Celda = Celda & OFV.gencelda("", " ", "center", "", "", "", "", "&nbsp;", "si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")   
      Formulario = Fila

  end if
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si")


  'sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>