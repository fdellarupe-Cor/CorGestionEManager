<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  SOPCION = request.querystring("OPCION")
  Regisvuelta = request.querystring("Regisvuelta")
  Regisperfil = request.querystring("Regisperfil")
  NOM = request.querystring("NOM")
  NOMU = request.querystring("NOM")
  NOM = ofv.convertircar(NOM,"*","&nbsp;")
  ID = request.querystring("ID")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  if Regisvuelta = "" then
    Regisvuelta = session("Regisvuel")
    Regisperfil = session("Regisperfil")
  else
    session("Regisvuel") = cdbl(Regisvuelta)
    session("Regisperfil") = cdbl(Regisperfil)
  end if

  if NOMU = "" then
    NOM = session("elNom")
    NOM = ofv.convertircar(NOM,"*","&nbsp;")
    NOMU = session("elNom")
  else
    session("elNom") = Nom
  end if

  if ID = "" then
    ID = session("elID")
  else
    session("elID") =cdbl(ID)
  end if

  sHTML = ofv.FormHeader(" ")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  sInput = "Sector:&nbsp;" & NOM

  Celda = ofv.gencelda("","ttt colspan=3","","","","","",sInput,"si")

  Formulario = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Formulario & ofv.Encabezados("#","ITEM","DESCRIPCION")

  set cn = ofv.conectar(ofv.strconn2)


''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA CONTROLES MITIGANTES - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN


        strsql = "Insert into ECODESCRIPCION "
        strsql = strsql & " (ECO,SECUEN,ECOITEM,ECODSC)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & ID & "','0','0',"
        strsql = strsql &  "'SIN DEFINIR'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF

          xTrColor = "#FFFFFF"

  Strsql = "Select * From EcoDESCRIPCION where "
  Strsql = Strsql & " ECO = 0" & ID & " ORDER BY SECUEN "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sForm = "form" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=ECODESCRIPCION&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("DESCRIPCION",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")

    'ITEM
    Sql1 = "Select * From TIPODSCSECTOR order by DESCRIPCION"
    sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","ut VALIGN='TOP' ","","30%","","","",sInput,"si")

    'DESCRIPCION
     sInput = "<font class=at><textarea class=dt name=" & rs(4).name & "  id=" & rs(4).name & " rows=3 cols=80 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 1950) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(4)&"")  &"</textarea></font>"
    Celda = Celda & ofv.gencelda("","uti  ","","","","","",sInput,"si")




    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#eeeeff"
       ELSE
          xTrColor = "#FFFFFF"
       END IF


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ECODESCRIPCION&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","-1","1","","",Formulario,"si")


  'Botonera
  Agregar = ""
  Buscar = ""

  if bAgregar then
     sAccion = "SECTORDESCRIPCION.asp?OPCION=SI&NOM=" & NOMU & "&ID=" & ID
     Agregar = ofv.BotonAgregar(sAccion)

  end if

  sAccion = "SectoresProp.asp?CantRegMover=" & regisvuelta   & "&Wperfil=" & regisperfil
  retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",agregar,retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>