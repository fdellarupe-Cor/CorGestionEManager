<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

 
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
'  response.buffer = true




  nCodigo = request.querystring("Codobjeto")  
 

	 Strsql = "Select Objsector.objid,Eco.descri From Objsector, Eco where eco.estado=3 and "
     Strsql = Strsql & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo

     response.write "SQL del SECTOR:   " & StrSql


  set cn = ofv.conectar(ofv.strconn2)
'response.querystring




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar




   'Sector
   if bModificar then

     Strsql = "Select Objsector.objid,Eco.descri From Objsector, Eco where eco.estado=3 and "
     Strsql = Strsql & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo

     
     ''sInput = ofv.GenerarCombo("edProcedimientos",rs(5).name,rs(5).value,Strsql,cn,0,1,"sb")
     
     Set rsSectorn = ofv.crearconsultaEx(strsql, cn, 1, parametros)
       'sInput = ofv.GenerarCombo("edProcedimientos",rs(6).name,rs(6).value,Strsql,cn,0,1,"sb")

       scomK = ""'"null" & Chr(9)
       scomD = "" '"" & Chr(9)
       SCOMSEP =""
     do while not rsSectorn.eof
        scomK=scomK & SCOMSEP & rsSectorn(0)
        scomd=scomd & SCOMSEP & rsSectorn(1)
        SCOMSEP = chr(9)
        rsSectorn.movenext
       loop
           sInput = ofv.gencomboFX("edProcedimientos",RS(5).Name,RS(5),scomK,scomD,"")
   else
     sSector = ""
     if rs(5) <> "" then

       Strsql = "Select descri From Eco, Objsector where "
       Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.objid = " & rs(5)
       set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsSector.eof then sSector = rsSector(0)
       call ofv.cerrarconsulta(rsSector)
     end if
'     sInput = ofv.GenerarInput("",sSector,"readonly","","40","","")
     sInput = sSector
   end if
   Celda = Celda & ofv.GenCelda("","tu","","20%","","","",sInput,"si")



'-----------------------------------------------------------------------------------------'

   'Perfil
   if bModificar then

     Strsql = "SELECT TIPOPERFIL.tipoperfil,TIPOPERFIL.descri "
     Strsql = Strsql & "FROM Objsector, ECO, ECOPERFIL, TIPOPERFIL "
     Strsql = Strsql & " where Eco.eco = Objsector.codsector and "
     Strsql = Strsql & " ECO.eco = ECOPERFIL.codsector and TIPOPERFIL.estado=3 and"
     Strsql = Strsql & " ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
     Strsql = Strsql & " and objsector.objid = " & rs(5).value
     Set rsPerfil = ofv.crearconsultaEx(strsql, cn, 1, parametros)
     'sInput = ofv.GenerarCombo("edProcedimientos",rs(6).name,rs(6).value,Strsql,cn,0,1,"sb")
      scomK = ""'"null" & Chr(9)

    scomD = "" '"" & Chr(9)
  	SCOMSEP =""
	  do while not rsperfil.eof
	   scomK=scomK & SCOMSEP & rsperfil(0)
	   scomd=scomd & SCOMSEP & rsperfil(1) 
	   SCOMSEP = chr(9)
	   rsPerfil.movenext
      loop
         sInput = ofv.gencomboFX("edProcedimientos", RS(6).Name, RS(6), scomK, scomD, "")
   else
     sPerfil = ""
     if rs(6) <> "" then
       Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
       set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsPerfil.eof then sPerfil = rsPerfil(0)
       call ofv.cerrarconsulta(rsPerfil)
     end if
'     sInput = ofv.GenerarInput("",sPerfil,"readonly","","40","","")
     sInput = sPerfil
   end if
   Celda = Celda & ofv.GenCelda("","tu","","20%","","","",sInput,"si")

   fila = ofv.GenRow("","","","","","","",celda,"si")

	if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID2 & "&Tabla=Objproced&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      tabla= ofv.GenTabla("","","","100%","","0","","-1","2","","",formulario1 & fila,"si")
      Formulario =   ofv.GenForm("edProcedimientos","",sAccion,"","post","",tabla,"si")
      'Formulario = Formulario & ofv.GenForm("edProcedimientos","",sAccion,"","post","",Celda,"si")
    else
      tabla= ofv.GenTabla("","","","100%","","0","","-1","2","","",formulario1 & fila,"si")
      Formulario =  ofv.GenForm("edProcedimientos","",sAccion,"","post","",tabla,"si")
      'Formulario = Formulario & ofv.GenForm("edProcedimientos","","","","post","",Celda,"si")
    end if

	response.write formulario






%>