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
  SOPCION = REQUEST.QUERYSTRING("OPCION")


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("CODIGO","DESCRIPCION","ENVIA A BCRA","*","-")

     AEstado = True
     AConsultar = True
     AEliminar = True
     AModificar = True
     AAgregar = True

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)



 ''' response.write "atrm"  & bagregar & "-" & bModificar & "-" & bSuprimir & "-" & Session("Form") & "<br>"


  IF SOPCION = "SI" THEN

        strsql = "Insert into ORESTADOS "
        strsql = strsql & " (CODIGO,DESCRIPCION,TIPO,ENVIA,"
        strsql = strsql &  "NEXT011,NEXT012,NEXT013,NEXT014,NEXT015,"
        strsql = strsql &  "NEXT016,NEXT017,NEXT018,NEXT019,NEXT0110"
        strsql = strsql &  ")"
        strsql = strsql & "  Values "
        strsql = strsql &  "('AAA','AAA - NUEVO ESTADO DE EVENTO','2','N',"
        strsql = strsql &  "'0','0','0','0','0',"
        strsql = strsql &  "'0','0','0','0','0'"
        strsql = strsql &  ")"

        '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF




  WKSETTABLA = "ORESTADOS"


  Strsql = "Select * From " & WKSETTABLA  & " WHERE TIPO <> '0' ORDER BY DESCRIPCION"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir AND (RS(3) & "") > "1" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    ' COD BCRA
    IF (RS(3) & "") > "1" THEN
       sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"10","dt","")
    ELSE
       sInput = "<B>" & RS(1) & "</B>"
    END IF
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'ENVIA
    scomK = "N" & chr(9) & "X" & chr(9) & "S"
    scomD = "NO ENVIA" & chr(9) & "SOLO VALIDA" & chr(9) & "ENVIA"
    sInput = ofv.gencomboFX(sForm,rs(14).name,rs(14),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT","","20%","","","",sInput,"si")

    BTIT1 = "<A CLASS=CC HREF='ORESTADOSSDKDET.ASP?NIVEL1=" & RS(0) & "' >PERMISOS</A>"
    Celda = Celda & ofv.gencelda("","VT bgcolor='#F0F0F0' style='border-style:solid;border-width:1;border-color:#F7AA22;'  ","CENTER","10%","","","",BTIT1,"si")


    BTIT1 = "<A CLASS=CC HREF='ORESTADOEVENTODET.ASP?NIVEL1=" & RS(0) & "' >BOTONES</A>"
    Celda = Celda & ofv.gencelda("","VT bgcolor='#F0F0F0' style='border-style:solid;border-width:1;border-color:#F7AA22;'  ","CENTER","10%","","","",BTIT1,"si")




    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORESTADOEVENTO.asp?OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRIPCION From " & WKSETTABLA   & " WHERE TIPO <> '0' ORDER BY DESCRIPCION"    ' orden pagina
  str2 = "Select DESCRIPCION From " & WKSETTABLA   & " WHERE TIPO <> '0' ORDER BY DESCRIPCION"    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORESTADOEVENTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML





%>
