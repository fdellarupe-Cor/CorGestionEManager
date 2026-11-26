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

  SESSION("OPCSET")=""
  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("CODIGO&nbsp;BCRA","CODIGO","Descripcion","ACUMULA","IMPACTO<BR>CONTABLE","PATRON","*")

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


 ' xxxx = ObtenerAtributos( Session("Form"),"","","")
 ' bAgregar = AAgregar
 ' bModificar = AModificar
 ' bSuprimir = AEliminar
 ' bConsultar = AConsultar


 ''' response.write "atrm"  & bagregar & "-" & bModificar & "-" & bSuprimir & "-" & Session("Form") & "<br>"


  IF SOPCION = "SI" THEN

        strsql = "Insert into ORTIPOCARACTER "
        strsql = strsql & " (DESCRI,ACUMULA,IMPCONT,PATRON,CODBCRA)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('NUEVO CARACTER DE EVENTO','OM','S','DEF','00')"

         '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF



  WKSETTABLA = "ORTIPOCARACTER"

  'if session("APLMDLOR") <> "BE"  THEN
   Sordenar="	Order By CODBCRA "
 ' else
 '  Sordenar = " Order By CAST(CODBCRA AS INT) "
 ' end if

   Strsql = "Select * From " & WKSETTABLA  & Sordenar
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
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


    ' COD BCRA
  '    if session("APLMDLOR") <> "BE"  THEN
'	   sInput = ofv.GenerarInput(rs(3).name,rs(3) ,"text",sForm,"15","dtn  maxlength=2","SI")
'	  ELSE

	   SINPUT="<input name='"& rs(3).name   &"'  value='" & trim(rs(3)&"") & "' type=text  MAXLENGTH=2 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(3)&"")&"',2);"&chr(34) & ">"


'	  end if
	 Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")


   'Tipo
       sInput = rs(0)
       Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'ACUMULA
    scomK = "OM" & chr(9) & "REC" & chr(9) & "DES" & chr(9) & "PER" & chr(9) & "NCN"

    scomD = "OTROS MONTOS" & chr(9) & "RECUPEROS" & chr(9) & "DESAFECTACIONES" & chr(9)
    scomD = scomD & "PERDIDAS" & chr(9) & "NO CONTABLES"
    sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT","","20%","","","",sInput,"si")


    XCOL = " BGCOLOR='#F7EEDD' "

    IF (RS(4) & "") = "N" THEN XCOL = " BGCOLOR='#DDDDDD' "

    'IMPCONT
    scomK = "S" & chr(9) & "N"
    scomD = "SI" & chr(9) & "NO"
    sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT" & XCOL,"","20%","","","",sInput,"si")

    XCOL = " BGCOLOR='#F7EEDD' "

    IF (RS(5) & "") = "ESP" THEN XCOL = " BGCOLOR='#F7AA22' "

    'PATRON
    scomK = "DEF" & chr(9) & "ESP"
    scomD = "DEFAULT" & chr(9) & "ESPECIAL"
    sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT" & XCOL,"","20%","","","",sInput,"si")

    BTIT1 = "<A CLASS=CC HREF='ORBASERIESGOSADD.ASP?MODAL=DOBLIGA&CantRegMover=" & cantregmoverx
    IF (RS(5) & "") = "DEF" THEN
       BTIT1 =  BTIT1 & "&OPCSET=DEF"
    ELSE
       BTIT1 =  BTIT1 & "&OPCSET=" & RS(0)
    END IF
    BTIT1 =  BTIT1 & "' >VER</A>"
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
    sAccion = "ORCARACTER.asp?OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & Sordenar    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & Sordenar    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORCARACTERCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  'shtml=shtml & "<script languaje=javascript> "
  'shtml=shtml & " function validarnum(obj,xform,vant,digitos) {"
  'shtml= shtml & " var re=new RegExp('^\\d{'+ digitos +'}$'); "
  'shtml= shtml & "if (obj.value.search(re)==-1) {alert('debe ser numerico y de ' + digitos +  ' posiciones'); obj.value=vant;return false;}"
  'shtml = shtml & "else{  var fr= document.getElementById(xform); fr.submit();}}"
  'shtml= shtml & "</script>"
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
    Else
        atrok = "S"
        AEstado = rs(1)
        AConsultar = rs(5)
        AEliminar = rs(4)
        AModificar = rs(3)
        AAgregar = rs(2)
    End If
    Call ofv.CerrarConsulta(rs)

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
          Set rs = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
          If rs.EOF Then
             atrok = "O"
          Else
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
        Set rs = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
        If rs.EOF Then
            atrok = "O"
        Else
           atrok = "S"
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
