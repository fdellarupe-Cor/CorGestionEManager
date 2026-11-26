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
  sOPCION       = request.querystring("OPCION")

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("COD. BCRA","codigo","Descripcion","&nbsp;")

  ofv.ObtenerAtributos Session("Form"),""
' xxxx = ObtenerAtributos( Session("Form"),"","","")

  bAgregar =ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
'   bAgregar = cbool(AAgregar)
'
'  bModificar = cbool(AModificar)
'  bSuprimir = cbool(AEliminar)
'  bConsultar = cbool(AConsultar)
   xTrColor = "#FFFFFF"
   sImage = "<font face='corflaticon' size=2 Xcolor='#000088' style='border-style:SOLID;border-width:1;' >4</font>"
response.write session("xsql")
set cn = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO EVENTO EN BASE DE EVENTOS DE PERDIDA

  if sopcion = "SI" then



        strsql = "Insert into ORTIPOEVENTO "
        strsql = strsql & " (DESCRI,secuen"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'NUEVO TIPO EVENTO','0'"
        strsql = strsql & ")"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if


  WKSETTABLA = "ORTIPOEVENTO"

'  if session("APLMDLOR") <> "BE"  THEN
'     Sordenar="	Order By objid "
'    else
'     Sordenar = " Order By CAST(CODBCRA AS INT) "
'  end if
   Sordenar="	Order By CODBCRA "
  Strsql = "Select * From " & WKSETTABLA  & Sordenar


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'sec
    'sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
   ' Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

     ' COD BCRA

      'sInput = ofv.GenerarInput(rs(3).name,rs(3) ,"text",sForm,"15","dtn  maxlength=2","SI")

'     if session("APLMDLOR") <> "BE"  THEN
'	   sInput = ofv.GenerarInput(rs(3).name,rs(3) ,"text",sForm,"15","dtn  maxlength=2","SI")
'	  ELSE
    if bModificar then
	   SINPUT="<input name='"& rs(3).name   &"'  value='" & trim(rs(3)&"") & "' type=text MAXLENGTH=2 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(3)&"")&"',2);"&chr(34) & ">"
	else
	   SINPUT="<input readonly name='"& rs(3).name   &"'  value='" & trim(rs(3)&"") & "' type=text MAXLENGTH=2 class=dtn  >"
	end if
'	  end if


	  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"90","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    BTIT1 = "<A CLASS=CC HREF='ORIMPACTORIESGODET.ASP?TIPOEV=" & RS(0) & "' >CATEGORIAS</A>"
    Celda = Celda & ofv.gencelda("","VT bgcolor='#F0F0F0' style='border-style:solid;border-width:1;border-color:#F7AA22;'  ","CENTER","10%","","","",BTIT1,"si")


    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF


    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORimpactoriesgo.asp?opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & Sordenar    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & Sordenar    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORIMPACTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

function ObtenerAtributos(sForm,  sNodo,  sCampo,  sValor)

       ' On Error GoTo ObtenerAtributosErr

     Dim strsql,  atrok,   sSQL,  xcount , XMFI , XMFSEP
     Dim XMFTGRP



     Dim XTabPer(4)

     xcount = 0


     AEstado = True
     AConsultar = True
     AEliminar = True
     AModificar = True
     AAgregar = True

     atrok = "N"
     strsql = "Select Alias,usrEstado,usragregar,usrmodificar,usreliminar,usrconsultar "
     strsql = strsql & " From atrmXusuario Where "
    If sNodo <> "" Then
        strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
    Else
        strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
    End If
    strsql = strsql & " And usuario = '" & Session("segperfil") & "' And "
    strsql = strsql & "  Aplicacion = '" & Session("CorApli") & "'"
    response.write strsql & "<br>"
    Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    If rs.EOF Then
        atrok = "O"
  '  response.write "aaa<br>"

    Else
        atrok = "S"
        AEstado = rs(1)
        AConsultar = rs(5)
        AEliminar = rs(4)
        AModificar = rs(3)
        AAgregar = rs(2)

  '  response.write "bbb<br>"


    End If
    Call ofv.CerrarConsulta(rs)

  '  response.write "atrok " & atrok & "<br>"

    If atrok = "O" Then
        strsql = "Select Alias,grpEstado,grpagregar,grpmodificar,grpeliminar,grpconsultar "
        strsql = strsql & " From atrmXgrupo Where  "
        If sNodo <> "" Then
            strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
        Else
            strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
        End If
        strsql = strsql & " And Aplicacion = '" & Session("CorApli") & "'"

     If Session("DELLOG") = "S" Then

          XMFTGRP = Split(Session("SegGrpLst"), Chr(9))
          xcount = UBound(XMFTGRP) + 1
          XMFSEP = ""
          For XMFI = 0 To 4
              XTabPer(XMFI) = xcount
          Next

          strsql = strsql & " and grupo in ("

          For XMFI = 0 To UBound(XMFTGRP)
              strsql = strsql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
              XMFSEP = ","
          Next
          strsql = strsql & ") "
   response.write strsql & "<br>"
          Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          If rs.EOF Then
             atrok = "O"
   ' response.write "ccc<br>"

          Else
   ' response.write "ddd<br>"

             atrok = "S"

             If Not rs(5) Then XTabPer(1) = XTabPer(1) - 1 ''' consultar
             If Not rs(4) Then XTabPer(2) = XTabPer(2) - 1 ''' eliminar
             If Not rs(3) Then XTabPer(3) = XTabPer(3) - 1 ''' modificar
             If Not rs(2) Then XTabPer(4) = XTabPer(4) - 1 ''' agregar

          End If

          If atrok = "S" Then

               do while not rs.eof

                 If XTabPer(1) = 0 Then AConsultar = False
                 If XTabPer(2) = 0 Then AEliminar = False
                 If XTabPer(3) = 0 Then AModificar = False
                 If XTabPer(4) = 0 Then AAgregar = False

                 rs.movenext
                loop

          End If


      Else
        strsql = strsql & " And grupo = '" & Session("seggrp") & "'  "
    response.write strsql & "<br>"
        Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        If rs.EOF Then
            atrok = "O"

   ' response.write "eee<br>"
        Else

   ' response.write "fff<br>"
           atrok = "S"
           '  If Not rs(1) Then AEstado = False ''' consultar
           '  If Not rs(5) Then AConsultar = False ''' consultar
           '  If Not rs(4) Then AEliminar  = False ''' eliminar
           '  If Not rs(3) Then AModificar = False ''' modificar
           '  If Not rs(2) Then AAgregar   = False ''' agregar

           '  If Not rs(1) Then AEstado = 0 ''' consultar
           '  If Not rs(5) Then AConsultar = 0 ''' consultar
           '  If Not rs(4) Then AEliminar  = 0 ''' eliminar
           '  If Not rs(3) Then AModificar = 0 ''' modificar
           '  If Not rs(2) Then AAgregar   = 0 ''' agregar

           '  If rs(1) = 0 Then AEstado = 0 ''' consultar
           '  If rs(5) = 0 Then AConsultar = 0 ''' consultar
           '  If rs(4) = 0 Then AEliminar  = 0 ''' eliminar
           '  If rs(3) = 0 Then AModificar = 0 ''' modificar
           '  If rs(2) = 0 Then AAgregar   = 0 ''' agregar


  '  response.write AConsultar & "<br>"
  '  response.write AEliminar & "<br>"
  '  response.write AModificar & "<br>"
   ' response.write AAgregar & "<br>"


            AEstado = rs(1)
            AConsultar = rs(5)
            AEliminar = rs(4)
            AModificar = rs(3)
            AAgregar = rs(2)
        End If

       End If

        Call ofv.CerrarConsulta(rs)
    End If


    ObtenerAtributos = "ok"


End function

%>