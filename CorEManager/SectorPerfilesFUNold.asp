<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  sperfil = request.querystring("perfil")
  SSOLAPA = request.querystring("SOLAPA")
  SOPCION = request.querystring("OPCION")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  if SSOLAPA = "" then
     SSOLAPA = "DSC"
  end if


  sHTML = ofv.FormHeader(" ")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  set cn = ofv.conectar(ofv.strconn2)






  Strsql = "Select Ecoperfil.*,Tipoperfil.descri From Ecoperfil, Tipoperfil where "
  Strsql = Strsql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and Ecoperfil.OBJID = 0" & sperfil
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

  posI = 0
   IF Not rs.EOF THEN

    sForm = "form" & posi

   Celda = ofv.GenCelda("","UT","","30%","","","","AREA/SECTOR:","si")


  SSECTOR = RS(1)
  SSECTORDAT = RS(1)

   Strsql = "Select * From ECO where eco = 0" & RS(1)
   set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rsX.EOF then
      SSECTOR = RSX(1)
   END IF

    'SECTOR
    sInput = SSECTOR
    Celda = Celda & ofv.gencelda("","ut ","","70%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ofv.GenCelda("","UT","","","","","","PUESTO:","si")

    'Perfil
    sInput = rs(4)
    Celda = Celda & ofv.gencelda("","ut ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","UT","","","","","","DOTACION:","si")

    'Dotacion
    sInput = rs(3)
    Celda = Celda & ofv.gencelda("","ut ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

  END IF

  sHTML = sHTML & ofv.GenTabla("","","","","","0","","-1","1","","",FILA,"si")

  xTrColor = "#FFFFFF"


''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA CONTROLES MITIGANTES - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN

        IF SSOLAPA = "RPTA" THEN

       SSTIPO = REQUEST.QUERYSTRING("SSTIPO")



        strsql = "Insert into ECOPERFILDEP "
        strsql = strsql & " (ECOPERFIL,SECUEN,PTOTIPO,PTOSECTOR)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SPERFIL & "','0','" & SSTIPO & "',"
        strsql = strsql &  "'" & SSECTORDAT & "'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    ELSEIF SSOLAPA = "RPTB" THEN

       SSTIPO = REQUEST.QUERYSTRING("SSTIPO")



        strsql = "Insert into ECOPERFILDEP "
        strsql = strsql & " (ECOPERFIL,SECUEN,PTOTIPO,PTOSECTOR)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SPERFIL & "','0','" & SSTIPO & "',"
        strsql = strsql &  "'" & SSECTORDAT & "'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    ELSEIF SSOLAPA = "DSC" THEN


        strsql = "Insert into ECOPERFILDSC "
        strsql = strsql & " (ECOPERFIL,SECUEN,PTODSCITEM,PTODESCRI)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SPERFIL & "','0','0',"
        strsql = strsql &  "'SIN DEFINIR'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    ELSEIF SSOLAPA = "FUN" THEN


        strsql = "Insert into ECOPERFILFUN "
        strsql = strsql & " (ECOPERFIL,SECUEN,PTOACCION,PTODESCRI)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SPERFIL & "','0','0',"
        strsql = strsql &  "'SIN DEFINIR'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    ELSEIF SSOLAPA = "COM" THEN


        strsql = "Insert into ECOPERFILCOM "
        strsql = strsql & " (ECOPERFIL,SECUEN,PTOCOMPE,PTODESCRI)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & SPERFIL & "','0','0',"
        strsql = strsql &  "'SIN DEFINIR'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    END IF


 END IF


 IF SOPCION = "UPD" THEN

    IF SSOLAPA = "COM" THEN

       SDATO  = REQUEST.QUERYSTRING("DATO")
       SCOMPE = REQUEST.QUERYSTRING("COMPE")


        strsql = "UPDATE ECOPERFILCOM "
        strsql = strsql & " SET PTODESCRI = '" & SCOMPE & "' "
        strsql = strsql & " WHERE OBJID = 0" & SDATO


        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    END IF


 end if


    Fila = ""

   SFILTRO  = SSOLAPA

  xTrColor = "#FFFFFF"

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "



'''''''''''''''''''''''''''''

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "
       wkstylo2 = " style='background-color:#FADDAA;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";border-style:solid;border-color:#DDDDDD;border-width:1;' "

     BCOLORG = "#AAAAAA"
     BCOLORD = "#6699CC"

    BTIT1 = "&nbsp;DESCRIPCION&nbsp;PUESTO&nbsp;"
    BTIT2 = "&nbsp;REPORTA&nbsp;A&nbsp;"
    BTIT3 = "&nbsp;LE&nbsp;REPORTA&nbsp;"
    BTIT4 = "&nbsp;FUNCIONES&nbsp;DEL&nbsp;PUESTO&nbsp;"
    BTIT5 = "&nbsp;COMPETENCIAS&nbsp;DEL&nbsp;PUESTO&nbsp;"


  if SFILTRO  = "DSC" then

          WKTIT = BTIT1

     B1 = "N"
     B2 = "S"
     B3 = "S"
     B4 = "S"
     B5 = "S"


  elseif SFILTRO  = "RPTA" then

          WKTIT = BTIT2

     B1 = "S"
     B2 = "N"
     B3 = "S"
     B4 = "S"
     B5 = "S"

  elseif SFILTRO  = "RPTB" then

          WKTIT = BTIT3

     B1 = "S"
     B2 = "S"
     B3 = "N"
     B4 = "S"
     B5 = "S"

  elseif SFILTRO  = "FUN" then

          WKTIT = BTIT4

     B1 = "S"
     B2 = "S"
     B3 = "S"
     B4 = "N"
     B5 = "S"

  elseif SFILTRO  = "COM" then

          WKTIT = BTIT5

     B1 = "S"
     B2 = "S"
     B3 = "S"
     B4 = "S"
     B5 = "N"


  end if


  WKTITG = " " & WKTIT


    BTIT1 = "<A CLASS=CC HREF='SECTORPERFILESFUN.ASP?SOLAPA=DSC&PERFIL=" & SPERFIL & "' >" & BTIT1 & "</A>"
    BTIT2 = "<A CLASS=CC HREF='SECTORPERFILESFUN.ASP?SOLAPA=RPTA&PERFIL=" & SPERFIL & "' >" & BTIT2 & "</A>"
    BTIT3 = "<A CLASS=CC HREF='SECTORPERFILESFUN.ASP?SOLAPA=RPTB&PERFIL=" & SPERFIL & "' >" & BTIT3 & "</A>"
    BTIT4 = "<A CLASS=CC HREF='SECTORPERFILESFUN.ASP?SOLAPA=FUN&PERFIL=" & SPERFIL & "' >" & BTIT4 & "</A>"
    BTIT5 = "<A CLASS=CC HREF='SECTORPERFILESFUN.ASP?SOLAPA=COM&PERFIL=" & SPERFIL & "' >" & BTIT5 & "</A>"



    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI " & wkstylo1,"","55%","","","",WKTITG,"si")

    IF B1 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT1,"si")
    IF B2 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT2,"si")
    IF B3 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT3,"si")
    IF B4 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT4,"si")
    IF B5 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT5,"si")


    Fila = FILA & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda,"si")
    sHTML = sHTML & "<CENTER><BR>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ","","100% ","","0","","1","1","","",Fila,"si") & "</CENTER>"


    Celda = ""

    Celda = Celda & "<div id='DX5' style='z-index:1;margin-top:1;margin-left:1;display:BLOCK;"
    Celda = Celda & "width:100%;height:65%;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
    Celda = Celda & "background-color:#F9F9F9;"
    Celda = Celda & "' >" & Chr(13)





    sHTML = sHTML & Celda




'''''''''''''''''''''''''''

  if SFILTRO  = "DSC" then

    Formulario = ""

  Strsql = "Select * FROM ECOPERFILDSC WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"
    wkproccodA = "<a class=nt href='SECTORPERFILESFUN.asp?PERFIL=" & SPERFIL & "&SOLAPA=" & SSOLAPA  & "&OPCION=SI' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"

   Celda = ofv.GenCelda("","UT style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","ut title='Agregar Descripciones.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
   Celda = CeldA & ofv.GenCelda("","UTI COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","","","","","DESCRIPCIONES DEL PUESTO","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo2,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","30%","","","","ITEM","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","70%","","","","DESCRIPCION","si")
   Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")


    Formulario = Formulario & FILA

  posI = 0
   IF Not rs.EOF THEN

  Do While (Not rs.EOF)

    sForm = "formRB" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EcoperfilDSC&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("DESCRIPCION",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")

    'ACCION
    Sql1 = "Select * From TIPODSCFUN order by descriPCION"
    sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")

    'DESCRIPCION
     sInput = "<font class=at><textarea class=dt name=" & rs(4).name & "  id=" & rs(4).name & " rows=3 cols=80 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 950) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(4)&"")  &"</textarea></font>"
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")




    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoperfilDSC&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")



    posi = posi + 1
    rs.Movenext
  Loop



  ELSE

    sInput = " SIN DEFINIR DESCRIPCIONES DEL PUESTO. "
    Celda =  ofv.GenCelda("","uti COLSPAN=4 style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput ,"si")

    Fila = ofv.genrow("","","","","60","","",Celda,"si")

    Formulario = Formulario &  FILA


  END IF

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","","","95%","","0","","1","1","","",Formulario,"si") & "</center>"



  elseif SFILTRO  = "RPTA" then


    Formulario = ""

  Strsql = "Select * FROM ECOPERFILDEP WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " AND PTOTIPO ='RA' ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"
    wkproccodA = "<a class=nt href='SECTORPERFILESFUN.asp?PERFIL=" & SPERFIL & "&SOLAPA=" & SSOLAPA  & "&SSTIPO=RA&OPCION=SI' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"


   Celda = ofv.GenCelda("","UT style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","ut title='Agregar Puestos.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
   Celda = CeldA & ofv.GenCelda("","UTI COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","","","","","REPORTA A","si")
   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo2,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","AREA/SECTOR","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","PERFIL","si")
   Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")

    Formulario = Formulario & FILA

  posI = 0
   IF Not rs.EOF THEN

  Do While (Not rs.EOF)

    sForm = "formRA" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EcoperfilDEP&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("Dependencia",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")

    'SECTOR
    Sql1 = "Select * From ECO order by descri"
    sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","60%","","","",sInput,"si")

    'Perfil
    Sql1 = "Select * From TIPOPERFIL WHERE TIPOPERFIL IN (SELECT CODPERFIL FROM ECOPERFIL WHERE CODSECTOR = 0" & rs(4) & ") "
    Sql1 = Sql1 & " order by descri"
    sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","60%","","","",sInput,"si")



    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoperfilDEP&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")



    posi = posi + 1
    rs.Movenext
  Loop



  ELSE

    sInput = " SIN DEFINIR SECTOR/PUESTO A QUIEN REPORTA. "
    Celda =  ofv.GenCelda("","uti COLSPAN=4 style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput ,"si")

    Fila = ofv.genrow("","","","","60","","",Celda,"si")

    Formulario = Formulario &  FILA


  END IF

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","","","95%","","0","","1","1","","",Formulario,"si") & "</center>"




  elseif SFILTRO  = "RPTB" then



    Formulario = ""

  Strsql = "Select * FROM ECOPERFILDEP WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " AND PTOTIPO ='RB' ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"
    wkproccodA = "<a class=nt href='SECTORPERFILESFUN.asp?PERFIL=" & SPERFIL & "&SOLAPA=" & SSOLAPA  & "&SSTIPO=RB&OPCION=SI' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"

   Celda = ofv.GenCelda("","UT style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","ut title='Agregar Puestos.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
   Celda = CeldA & ofv.GenCelda("","UTI COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","","","","","PUESTO QUE LE REPORTA","si")
   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo2,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","AREA/SECTOR","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","50%","","","","PERFIL","si")
   Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")

    Formulario = Formulario & FILA

  posI = 0
   IF Not rs.EOF THEN

  Do While (Not rs.EOF)

    sForm = "formRB" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EcoperfilDEP&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("Dependencia",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")

    'SECTOR
    Sql1 = "Select * From ECO order by descri"
    sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","60%","","","",sInput,"si")

    'Perfil
    Sql1 = "Select * From TIPOPERFIL WHERE TIPOPERFIL IN (SELECT CODPERFIL FROM ECOPERFIL WHERE CODSECTOR = 0" & rs(4) & ") "
    Sql1 = Sql1 & " order by descri"
    'response.write sql1
    sInput = ofv.GenerarCombo(sForm & "a" ,rs(5).name,rs(5).value,sql1,cn,0,1,"")





    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","60%","","","",sInput,"si")

    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoperfilDEP&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")



    posi = posi + 1
    rs.Movenext
  Loop



  ELSE

    sInput = " SIN DEFINIR SECTOR/PUESTO QUE LE REPORTA. "
    Celda =  ofv.GenCelda("","uti COLSPAN=4 style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput ,"si")

    Fila = ofv.genrow("","","","","60","","",Celda,"si")

    Formulario = Formulario &  FILA


  END IF

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","","","95%","","0","","1","1","","",Formulario,"si") & "</center>"


  elseif SFILTRO  = "FUN" then

    Formulario = ""

  Strsql = "Select * FROM ECOPERFILFUN WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"
    wkproccodA = "<a class=nt href='SECTORPERFILESFUN.asp?PERFIL=" & SPERFIL & "&SOLAPA=" & SSOLAPA  & "&OPCION=SI' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"

   Celda = ofv.GenCelda("","UT style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","ut title='Agregar FUNCIONES.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
   Celda = CeldA & ofv.GenCelda("","UTI COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","","","","","FUNCIONES DEL PUESTO","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo2,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","30%","","","","FUNCION","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","70%","","","","DESCRIPCION","si")
   Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")


    Formulario = Formulario & FILA

  posI = 0
   IF Not rs.EOF THEN

  Do While (Not rs.EOF)

    sForm = "formRB" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EcoperfilFUN&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("Dependencia",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")

    'ACCION
    Sql1 = "Select * From TIPOACCIONFUN order by descriPCION"
    sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")

    'DESCRIPCION
     sInput = "<font class=at><textarea class=dt name=" & rs(4).name & "  id=" & rs(4).name & " rows=3 cols=80 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 450) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(4)&"")  &"</textarea></font>"
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")




    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoperfilFUN&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")



    posi = posi + 1
    rs.Movenext
  Loop



  ELSE

    sInput = " SIN DEFINIR FUNCIONES DEL PUESTO. "
    Celda =  ofv.GenCelda("","uti COLSPAN=4 style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput ,"si")

    Fila = ofv.genrow("","","","","60","","",Celda,"si")

    Formulario = Formulario &  FILA


  END IF

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","","","95%","","0","","1","1","","",Formulario,"si") & "</center>"


  elseif SFILTRO  = "COM" then

    Formulario = ""

  Strsql = "Select * FROM ECOPERFILCOM WHERE "
  Strsql = Strsql & " ECOPERFIL = 0" & sperfil & " ORDER BY SECUEN "
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    Fila = ""

    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"
    wkproccodA = "<a class=nt href='SECTORPERFILESFUN.asp?PERFIL=" & SPERFIL & "&SOLAPA=" & SSOLAPA  & "&OPCION=SI' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"

   Celda = ofv.GenCelda("","UT style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","ut title='Agregar COMPETENCIAS.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
   Celda = CeldA & ofv.GenCelda("","UTI COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;' " & wkstylo1,"","","","","","COMPETENCIAS DEL PUESTO","si")

   Fila  = ofv.genrow("","","","","25","","",Celda,"si")

   Celda = ofv.GenCelda("","UTI COLSPAN=2  " & wkstylo2,"","2%","","","","&nbsp;","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","30%","","","","COMPETENCIA","si")
   Celda = CeldA & ofv.GenCelda("","utI  " & wkstylo2,"","70%","","","","DESCRIPCION","si")
   Fila  = Fila & ofv.genrow("","","","","25","","",Celda,"si")


    Formulario = Formulario & FILA

  posI = 0
   IF Not rs.EOF THEN

  Do While (Not rs.EOF)

    sForm = "formRB" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EcoperfilCOM&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("Dependencia",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SECUEN
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"3","dtn","")
    Celda = Celda & ofv.gencelda("","UT VALIGN='TOP' ","","2%","","","",sInput,"si")

    'ACCION
    Sql1 = "Select * From TIPOCOMPETENCIA order by descriPCION"
    sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")


    Strsql = "Select * From tipocompetencia where objid = 0" & rs(3)
'' response.write strsql
    set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    if  not rsX2.eof then
        SCOMPETIPO = CSTR(rsX2(2) & "")
    END IF


    IF SCOMPETIPO = "1" THEN

       SCOMPEDAT = "<FONT COLOR='#993333' >SIN DEFINIR</FONT>"
       SCOMPEDET = rs(4) & ""

       IF NOT ISNUMERIC(SCOMPEDET) THEN SCOMPEDET = 0

      Strsql = "Select * From TIPOCOMPEDET where objid = 0" & SCOMPEDET
'' response.write strsql
      set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      if  not rsX2.eof then
        SCOMPEDAT = CSTR(rsX2(2) & "")
      END IF

        xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
        wkproccodA = "<a class=nt TARGET='fesquemas' href='GENPTOCOMPE.ASP?PERFIL=" & SPERFIL & "&DATO=" & RS(0) & "&COMPE=" & RS(3) & "' ><b>"
        wkproccodA = wkproccodA & SCOMPEDAT
        wkproccodA = wkproccodA & "</b></a>"
	Celda = Celda & ofv.GenCelda("","utI style='border-style:solid;border-color:#888888;border-width:1;cursor:pointer;' title='SELECCIONAR COMPETENCIA.' " & xsffun,"","","","","",wkproccodA,"si")



    ELSE



    'DESCRIPCION
     sInput = "<font class=at><textarea class=dt name=" & rs(4).name & "  id=" & rs(4).name & " rows=3 cols=80 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 450) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(4)&"")  &"</textarea></font>"
    Celda = Celda & ofv.gencelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")

   END IF




    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoperfilCOM&volver=" & vuelta & "_SOLAPA=" & SSOLAPA & "_perfil=" & Sperfil
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")



    posi = posi + 1
    rs.Movenext
  Loop



  ELSE

    sInput = " SIN DEFINIR COMPETENCIAS DEL PUESTO. "
    Celda =  ofv.GenCelda("","uti COLSPAN=4 style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput ,"si")

    Fila = ofv.genrow("","","","","60","","",Celda,"si")

    Formulario = Formulario &  FILA


  END IF

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","","","95%","","0","","1","1","","",Formulario,"si") & "</center>"


  end if





'''''''''''''''''''''''''''

    sHTML = sHTML & "</div>" & Chr(13)

  'Botonera
  Agregar = ""
  Buscar = ""

'  if bAgregar AND SSOLAPA <> "DSC" then
'     sAccion = "SECTORPERFILESFUN.asp?PERFIL=" & SPERFIL & "&SOLAPA=" & SSOLAPA & "&OPCION=SI"
'     Agregar = ofv.BotonAgregar(sAccion)
'  end if

  sAccion = "SectorPERFILES.asp"
  retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",agregar,retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>