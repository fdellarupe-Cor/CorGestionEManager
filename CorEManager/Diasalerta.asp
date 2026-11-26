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

  'Encabezados de la tabla
  Formulario = ofv.Encabezados("Id","Color","Mayor a","Dias","Menor o Igual a","Dias" )


  ofv.ObtenerAtributos Session("Form"),""
  'bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  'bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

			'********************************************************
			'************ Variables para los colores del semaforo
			'*******************************************************
			
	   	strVerde = "SELECT * FROM diasalerta "
			strVerde = strVerde & "WHERE color = 'Green'"
		
			Set rsVerde = ofv.crearconsultaEx(strVerde,cn,1,parametros)
			
			MayorVerde = rsVerde(2)
			MenorVerde = rsVerde(3)
			
			'response.write "Mayor:" & MayorVerde & "Menor:" & MenorVerde
			
			strAmarillo = "SELECT * FROM diasalerta "
			strAmarillo = strAmarillo & "WHERE color = 'Yellow'"
		
			Set rsAmarillo = ofv.crearconsultaEx(strAmarillo,cn,1,parametros)
			
			MayorAmarillo = rsAmarillo(2)
			MenorAmarillo = rsAmarillo(3)
			
			'response.write "Mayor:" & MayorAmarillo & "Menor:" & MenorAmarillo

		
			strRojo = "SELECT * FROM diasalerta "
			strRojo = strRojo & "WHERE color = 'Red'"
		
			Set rsRojo = ofv.crearconsultaEx(strRojo,cn,1,parametros)
			
			MayorRojo = rsRojo(2)
			MenorRojo = rsRojo(3)
			
			'response.write "Mayor:" & MayorRojo & "Menor:" & MenorRojo
			
			
'********
'***** Armado de la tabla *****
'********

  Strsql = "Select * From diasalerta order by id"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then rs.move(CantRegMoverx)

  posi = 0

  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

  	sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=motivos&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'ID
    sInput = rs(0) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","cc","right","50","","","",sInput,"si")

    'COLOR
    sInput = " "
    Celda = Celda & ofv.gencelda("","aaa","","","",rs(1),"",sInput,"si")

	'MAYOR QUE
	
	'IF (rs(1)& "") = "Red" THEN
		'sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"50","dt readonly","si")
		'else
		
			'**** INCIDENTE 514 *****'
			''Se modifica el campo MAYOR A ya que no graba bien las modificaciones
			
			'CODIGO VIEJO'
			'sInput = ofv.GenerarInput(rs(2).name,rs(2),"readonly",sForm,"50","dt","")
			'end if
			
			'CODIGO NUEVO'
			sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"50","dt","")
			
			'*** FIN INCIDENTE 514 ***'
		
		Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
	
	'MAYOR QUE dias
    sInput = "Dias"
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

   'MENOR O IGUAL QUE
   	'IF (rs(1)& "") = "Red" THEN
		'sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"50","dt readonly","")
		'else
		sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"50","dt","")
		'end if
    
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
 
    'MENOR O IGUAL QUE dias
    sInput = "Dias"
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")
	
	Fila = ofv.genrow("","","","","","","",Celda,"si")



	'********
	'*** Actualizar 
	'*********
	
    sAccion = ""
	
    if bModificar  then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=diasalerta&volver=" & vuelta
    end if
	'response.write ofv.convertircar(rs(3)," ","*")
	
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  
  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

	'response.write "El Valor Mayor Verde es: " & MayorVerde & "<br>"
	'response.write "El Valor MenorIgual Amarillo es: " & MenorAmarillo & "<br>"

'cn.open
'set rs = Server.CreateObject("ADODB.Recordset")
'rs.open "diasalerta",cn',adOpenForwardOnly,adLockReadOnly,adCmdTable
	'For each campo in rs.fields
	'response.write "Nombre: " & campo.Name & "<br>"
	'response.write "Value: " & campo.Value & "<br>"
	'next
	'rs.close
	'set rs= nothing
	

  'Botonera
  'Agregar = ""
  'if bAgregar then
    'sAccion = "AgregarMotivos.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=motivos"
    'Agregar = ofv.BotonAgregar(sAccion)
  'end if

  'str1 = "Select descripcion From motivos order by id"    ' orden pagina
  'str2 = "Select descripcion From motivos order by id"    ' orden alfabetico
  'Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,"no","")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

	
'shtml = "<script type='text/javascript'>alert('holaaaaa');</script>"
'sHTML = sHTML & "<script language='javascript'>alert('Menor Rojo No puede ser superior a Mayor Amarillo')</script>"
'shtml = "<script type='text/javascript'>alert('Javaescriiiiiiiiiiiiii');</script>"
'Response.Write sHTML

'shtml = "<script type='text/javascript'>prompt('Javaescriiiiiiiiiiiiii');</script>"

Response.Write sHTML
%>