<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

    Response.buffer = True
  APLCHECK()
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  CantRegMostrar = 12
  CantRegMover = cdbl(request.querystring("CantRegMover"))

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")

  sTest = request.querystring("Test")

  ID      = request.querystring("ID")
  pmk     = request.querystring("pmk")
  OSec    = request.querystring("OSec")
  'response.write id

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if
  xxform=""
  if OSec <> "" then
     session("OSec") = OSec
  else
     OSec = session("OSec")
  end if


  if OSec = "" then
     OSec = "NO"
     session("OSec") = OSec
  end if


  wposi   = 0
  wpossal = 0

     session("comp31") = "TODOS"
     session("comp32") = "JPG"

'  vuelta = ofv.vueltaaspEx("../",CantRegMostrar,CantRegMover) & "_ID=" & ID

  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID
  vuelta2 = ofv.vueltaaspEx("../",CantRegMostrar2,CantRegMover2) & "_ID=" & ID
  vueltaC = ofv.vueltaaspExC("../",CantRegMover,CantRegMover2) & "_ID=" & ID

  vueltaD = ofv.vueltaaspExC("../",CantRegMover,CantRegMover2)


  eststt = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  eststa = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst1 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst2 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst3 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst4 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  estst5 = " STYLE={font-family:tahoma,arial;color:#000045;font-size:10pt;background-color:#cceeff;border-style:solid;border-width:1;border-color:#cccccc;} "
  esttxt = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#f7eedd;} "

  if pmk = "" then
     sfuncion = ""
  else
    sfuncion = " onload=" & chr(34) & "window.scrollTo(document.edestrdet" & pmk & ".offsetParent.offsetLeft,document.edestrdet" & pmk & ".offsetParent.offsetTop);" & chr(34)
  end if

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2' style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=1111>"

'-------------------------------------------------

  dim jsHashPefiles
   jsHashPefiles="hashperfiles"


    dim jsHashsectores
  jsHashsectores="hashsectores"

  set cn = ofv.conectar(ofv.strconn2)
 if session("APLMDCLI") =  "BCR" or session("APLMDCLI") = "BGAL" or session("APLMDCLI") = "BANELCO" then

  shtml=shtml & "<script type='text/javascript' src='hashtable.js'></script>"
  sHTML = sHTML & "<script> var hashperfiles;var hashsectores;" & qjsHashPefiles & "</script>"
  shtml=shtml & "<script language=javascript>" & chr(13)
   shtml = shtml  & armaperfiles(jsHashPefiles,cn,nCodigo) &  armasectores(jsHashsectores,cn,nCodigo)
     shtml = shtml & "function AbrirVentana(datos) { " & Chr(13)
     shtml = shtml & " window.open(datos"
     shtml = shtml & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
     shtml = shtml & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
     shtml = shtml & " ); " & Chr(13)
     shtml = shtml & " } " & Chr(13) & Chr(13)
   shtml=shtml & "</script>" & chr(13)
  end if
  sHTML = replace(sHTML,"<body ","<body " & sfuncion)


  response.write sHTML


  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
 'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    Nestado=rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "Select * From Objestr Where objid = " & ID
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bsecuen = ""
  bDescri = " "
  if not rs.eof then
     bsecuen = rs(4)
     bDescri = rs(3)
  end if
  call ofv.cerrarconsulta(rs)




    Formulario = ""

  celda = ""
  sImageSc1 = "<i class='las la-eye' ></i><i class='las la-angle-right' ></i>"
   sImageSc2 = "<i class='las la-eye' ></i><i class='las la-angle-left' ></i>"
   sImageadd = "<i class='las la-plus' ></i>"

  sImageCap = "<i class='las la-pencil-alt' ></i>"
  sImageTit = "<i class='las la-pencil-alt' ></i>"'titulo'
  sImageSt1 = "<i class='las la-pencil-alt' ></i>"'subtitulo1'
  sImageSt2 = "<i class='las la-pencil-alt' ></i>"'subtitulo2'
 sImageSt3 = "<i class='las la-pencil-alt' ></i>" 'subtitulo3'
  sImageSt4 = "<i class='las la-pencil-alt' ></i>"'subtitulo4'
  sImageSt5 = "<i class='las la-pencil-alt' ></i>"'subtitulo5'
  sImageItG = "<i class='las la-pencil-alt' ></i>" 'texto del contenido'
  sImageItH = "<i class='las la-pencil-alt' ></i>"
  sImageItI = "<i class='las la-pencil-alt' ></i>"
  sImageItJ = "<i class='las la-pencil-alt' ></i>"



  '   sImageCap = "<font class='btiw' >&#57386;</font>"
  '  sImageTit = "<font class='btiw' >&#57360;</font>"
  '  sImageSt1 = "<font class='btiw' >&#57441;</font>"
  '  sImageSt2 = "<font class='btiw' >&#57442;</font>"
  '  sImageSt3 = "<font class='btiw' >&#57443;</font>"
  '  sImageSt4 = "<font class='btiw' >&#57444;</font>"
  '  sImageSt5 = "<font class='btiw' >&#57445;</font>"
  '  sImageItG = "<font class='btic' >&#57433;</font>"
  '  sImageItH = "<font class='btic' >&#57360;</font>"
  '  sImageItI = "<font class='btic' >&#57433;</font>"
  '  sImageItJ = "<font class='btic' >&#57433;</font>"

  redim nivel(12)
  nivel(0)  = "120"
  nivel(1)  = "110"
  nivel(2)  = "100"
  nivel(3)  = "90"
  nivel(4)  = "80"
  nivel(5)  = "70"
  nivel(6)  = "60"
  nivel(7)  = "50"
  nivel(8)  = "40"
  nivel(9)  = "30"
  nivel(10) = "20"
  nivel(11) = "10"
  nivel(12) = "5"

  redim nivel2(2)
  nivel2(0)  = "IZQUIERDA"
  nivel2(1)  = "CENTRADO"
  nivel2(2)  = "DERECHA"


 '' response.write "<BIGGER>"

  response.write "<TABLE class=aaa width=100% marginwidth=0  cellspacing=0  cellpadding=2 >" & CHR(13)

'''''''''''''''''''''''

    'secuencia
      if OSec <> "SI" then

        'Secuencia Oculta
         sAccion = "edEstrDet.asp?ID=" & ID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sAccion = sAccion & "&OSec=SI" & "&cantregmostrar=" & cantregmostrar & "&CantRegMover=" & CantRegMover

       xclas   = "btbn"
       'xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
       sHref   = sAccion
       sInput  = "<div class=" & xclas & " >" & sImageSc1 & "</div>" & amk
     ''  Celda1  = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Mostrar secuencia' ","","","1","","",sInput,"si")
    ''   Celda1  = ofv.GenLink("",xclas & " title='Mostrar secuencia' ",sHref,"","","",Celda1,"si")

              XXX= ofv.GenLink("",xclas & " title='Mostrar secuencia' ",sHref,"","","",sInput,"si")
 Celda1 = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Mostrar secuencia' ","","","1","","",xxx,"si")

      else

        'Secuencia Visible
         sAccion = "edEstrDet.asp?ID=" & ID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sAccion = sAccion & "&OSec=NO" & "&cantregmostrar=" & cantregmostrar & "&CantRegMover=" & CantRegMover

       xclas   = "btbn"
       'xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
       sHref   = sAccion
       sInput  = "<div class=" & xclas & " >" & sImageSc2 & "</div>" & amk
      '' Celda1  = ofv.GenCelda("",xclas & xsffun & " style='filter:shadow(color=#dddddd,direction=125);' title='Ocultar secuencia' ","","","1","","",sInput,"si")
     ''  Celda1  = ofv.GenLink("",xclas & " title='Ocultar secuencia' ",sHref,"","","",Celda1,"si")
             XXX= ofv.GenLink("",xclas & " title='Ocultar secuencia' ",sHref,"","","",sInput,"si")
 Celda1 = Celda1 & ofv.GenCelda("",xclas & xsffun & " style='filter:shadow(color=#dddddd,direction=125);' title='Ocultar secuencia' ","","","1","","",xxx,"si")

      end if

    Formulario = Formulario & Celda1

    'Agregar Titulos
    if bModificar then
       sAccion = "AgregarEstDet.asp?ID=" & ID & "&Tabla=Objestdet&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&IDx=0&IDt=X"
       xclas   = "btbn"
       'xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
       sHref   = sAccion
       sInput  = "<div class=" & xclas & " >" & sImageadd & "</div>"
    ''     Celda1  = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Agregar contenido al Capitulo' ","","","1","","",sInput,"si")
  ''     Celda1  = ofv.GenLink("",xclas & " title='Agregar contenido al Capitulo' ",sHref,"","","",Celda1,"si")

           xxx1= ofv.GenLink("",xclas & " title='Agregar contenido al Capitulo' ",sHref,"","","",sInput,"si")
Celda1 =  ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Agregar contenido al Capitulo' ","","","1","","",xxx1,"si")

    else
       Celda1  = ofv.GenCelda("","ut","","","1","","","&nbsp;","si")
       sAccion = ""
    end if

    Formulario = Formulario & Celda1

    'Propiedades
    if bModificar then
       sAccion = "EdestrPrtyMST.asp?volver=" & vuelta & "&Codobjeto=" & nCodigo & "&Tipobjeto="
       sAccion = sAccion & sTipo & "&ID=" & ID
       xclas   = "btbn"
       'xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
       sHref   = sAccion
       sInput  = "<div class=" & xclas & " >" & sImageCap & "</div>"
        'Celda2  = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Propiedades del Capitulo' ","","","1","","",sInput,"si")
       'Celda2  = ofv.GenLink("",xclas & " title='Propiedades del Capitulo' ",sHref,"","","",Celda2,"si")

           xxx2= ofv.GenLink("",xclas & " title='Propiedades del Capitulo' ",sHref,"","","",sInput,"si")
Celda2 =  ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Propiedades del Capitulo' ","","","1","","",xxx2,"si")

    else
       Celda2 = ofv.GenCelda("","ut","","","1","","","&nbsp;","si")
       sAccion = ""
    end if

    Formulario = Formulario & Celda2

    'Titulo del Capitulo
    Formulario = Formulario & ofv.GenCelda("","ttt colspan=9 " & bstylo,"","90%","","","","&nbsp;" & bdescri,"si")
    Formulario = Formulario & ofv.GenCelda("","vt","","","","","","&nbsp;","si")

  Fila = ofv.GenRow("","","","","","","",Formulario,"si") & CHR(13)
  response.write Fila

  Formulario = ""
  strsql = ""



''''''''''''''''''''''''

  strsql = " SELECT CSESTDET.*, b.codinx2  FROM CSESTDET "
 ' strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (CSESTDET.objid = b.codinx2) "
 strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast ( CSESTDET.objid as nvarchar) = b.codinx2) "

  strsql = strsql & " where detcon = " & ID & "  "
  if bVersion = "NO" then
     strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
  end if
  strsql = strsql & "and not ( "
  strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
  strsql = strsql & ") and stmodi = ''  "
  strsql = strsql & " order by TSEC"
'response.write  strsql & "<br>"

Dim Inx
dim nrs(100)

'''''''''''''''''''''''''


Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorType = 3
Rs.CacheSize = 15
'Rs.Cursorlocation = 3
Rs.Open strsql, cn

  for inx = 0 to (rs.fields.count - 1)
      nrs(inx) = rs(inx).name
  next

 campos = rs.fields.count - 1

 if not rs.eof then

    if cantregmover > 0  then rs.move cantregmover

     if not rs.eof then
       ars = rs.GetRows(cint(CantRegMostrar)+1,0)
     else
       redim  ars(cint(campos),0)
       ars(39,0) = "N"
     end if

 else

    redim  ars(cint(campos),0)

    ars(39,0) = "N"

 end if

xsfii=0
''response.write "aca paso 2"
' if not rs.eof then
'    for xsfii=1 to cantregmostrar
'        rs.movenext
'        if rs.eof then
'           xsfii = xsfii - 1
'           exit for
'        end if
'    next
' end if



rs.Close
set rs = nothing


  xTrColor = "#FFFFFF"

  sant = "X"
  bseci = 0
  bsecix = 0
  redim preserve bsec(8)
  for i=0 to 8
      bsec(i) = 0
  next
  amkeya = 0
  wentro = 0



 inx = 0

 ' if not rs.eof then
 '    rs.move cantregmover
'  end if

 do while wentro = 0

    '  do while not rs.eof

for Inx = inx to UBound(ars,2)

         IF ars(39,inx) = "N" OR ((ars(40,inx) & "") = "" AND clng("0" & ars(6,inx)) > 0 )THEN
           inx = inx + 1
           exit for
         END IF


         SCLAVE    = ars(0,inx)
         Stip      = ars(5,inx)
         smodi     = ars(29,inx)
         stipmodi  = ars(30,inx)
         sForm     = "edestrdet" & SCLAVE



      'Secuencia
      select case stip
             case "A"
                  btip = "TIT"
                  btipi = sImageTit
                  btipp = "TITULO"
                  bseci = 0
                  bsecix = 0
                  bsec(1) = ars(4,inx)

             case "B"
                  btip = "ST1"
                  btipi = sImageSt1
                  btipp = "SUBTITULO 1"
                  bseci = 1
                  bsecix = 1
                  bsec(2) = ars(4,inx)



             case "C"
                  btip = "ST2"
                  btipi = sImageSt2
                  btipp = "SUBTITULO 2"
                  bseci = 2
                  bsecix = 2
                  bsec(3) = ars(4,inx)
             case "D"
                  btip = "ST3"
                  btipi = sImageSt3
                  btipp = "SUBTITULO 3"
                  bseci = 3
                  bsecix = 3
                  bsec(4) = ars(4,inx)
             case "E"
                  btip = "ST4"
                  btipi = sImageSt4
                  btipp = "SUBTITULO 4"
                  bseci = 4
                  bsecix = 4
                  bsec(5) = ars(4,inx)
             case "F"
                  btip = "ST5"
                  btipi = sImageSt5
                  btipp = "SUBTITULO 5"
                  bseci = 5
                  bsecix = 5
                  bsec(6) = ars(4,inx)
             case "G"
                  btip = "TXT"
                  btipi = sImageItG
                  btipp = "TEXTO"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = ars(4,inx)
             case "H"
                  btip = "IMG"
                  btipi = sImageItH
                  btipp = "IMAGEN"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = ars(4,inx)
             case "I"
                  btip = "TAB"
                  btipi = sImageItI
                  btipp = "TABLA"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = ars(4,inx)
             case "J"
                  btip = "CMP"
                  btipi = sImageItJ
                  btipp = "COMPONENTE"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bsecix = bseci + 1
                  bsec(7) = ars(4,inx)

      end select
      Amk = "<a name=" & chr(34) & "X" & ars(0,inx) & chr(34) & "> "
      amkey = ars(0,inx)

      Fila        = ""
      Formulario  = ""
      Formulario1 = ""
      Formulario2 = ""

     'Eliminar
      sHref = ""

  '' response.write "001 " & smodi
  '' response.write "002 " & stipmodi
     sImage = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")

      if smodi And stipmodi = "E" then

''      sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
        sImage = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
        Celda = ofv.GenCelda("","","","","","","",sImage ,"si")
      else
        if bModificar then
           if cstr(amkeya) = cstr(0) then
               vlk = vuelta
           else
               vlk = vuelta & "_pmk=" & amkeya
           end if
          if smodi And stipmodi = "A" then
            sHref = "asp/EliminarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vlk
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          else
            sHref = "asp/ActualizarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vlk
            sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=SI"
          end if
        end if
        if sHref <> "" then

           Celda = BotonEliminar(BtipP ,sHref,SCLAVE)

        else
          Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
        end if
       '''''      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      Formulario1 = Celda

      if bModificar then
       'Agregar Titulos

         if stip <> "G" and stip <> "H" and stip <> "I" and stip <> "J" Then



            vlk = vuelta & "_pmk=" & amkey
            sAccion = "AgregarEstDet.asp?ID=" & ID & "&Tabla=Objestdet&volver=" & vlk
            sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&IDx=" & sclave
            sAccion = sAccion & "&IDt=" & stip

            xclas   = "btbn"
            'xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
            sHref   = sAccion
            sInput  = "<div class=" & xclas & " >" & sImageadd & "</div>"

            xxx  = ofv.GenLink("",xclas & " title='Agregar contenido al " & BtipP & "' ",sHref,"","","",sInput,"si")
                'XXX= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
            Celda1  = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;'  title='Agregar contenido al " & BtipP & "'  ","","","1","","",xxx,"si")
           'Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","center","5%","","","",XXX,"si")


           Formulario1 = Formulario1 & Celda1       ''        se elimino ya que se duplicaba




         else
            Celda = ofv.GenCelda("","ut","","20","","","","&nbsp;","si")
            Formulario1 = Formulario1 & Celda 'agregaba un espacio a la celda!

         end if

''''''''''''''''''
       'Propiedades
         if stip <> "H" and stip <> "I" and stip <> "J" Then

            sImage = "<B>" & BtipI & "</B>"

            sAccion = "EdestrPrtyMST.asp?volver=" & vuelta  & "&Codobjeto=" & nCodigo
            sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Part=SI&TIPO=" & stip  & "&ID=" & sclave  & "&ID2=" & ID

            xclas   = "btbn"

            xclasadd   = ""
            bpespec    = ""

            if  ars(31,inx) = "NO" then

                xclasadd   = " style='color:#FFFFFF;background-color:" & Session("CMFPNBOT") & ";border-radius:15px;' "
                bpespec    = chr(13) & "(propiedades especiales activadas)."

            end if

            ''xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
            sHref   = sAccion
            sInput  = "<div class=" & xclas & xclasadd  & " >" & BtipI & "</div>"  '& stip

     ''       Celda1  = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Propiedades de " & BtipP & bpespec & "' ","","","1","","",sInput,"si")
   ''    Celda1  = ofv.GenLink("",xclas & " title='Propiedades de " & BtipP & bpespec    & "' ",sHref,"","","",Celda1,"si")

      'XXX= ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")
     xxx= ofv.GenLink("",xclas & " title='Propiedades de " & BtipP & bpespec    & "' ",sHref,"","","",sInput,"si")
'Celda = Celda & ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","center","5%","","","",XXX,"si")
      celda1 =  ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Propiedades de " & BtipP & bpespec & "' ","","","1","","",xxx,"si")


      ' estado observacion

            Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")

           'response.write nCodigo & " nCodigo " & "<br>"
           'response.write sTipo   & " sTipo " & "<br>"

      ''''' agregado para color del cuadrito de Observaciones version 4.6
            wComentario = ""
            strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and "
            strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"      ' and "
            strsql = strsql & " order by codcircuito, codusr, objid"
          ''response.write strsql
            set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rss.eof then
               wComentario = rss(13)
            end if
            call ofv.cerrarconsulta(rss)

           if (Nestado=5  or Nestado=6) and (ars(42,inx) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & ars(42,inx)  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion

     ''''''agregado para color del cuadrito de Observaciones version 4.6

               if wComentario = "" then
      	          sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
      	   '''''''sInput  = "<div style=' border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
       	       else
       	          sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
       	       end if
  ''''''

               Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")

               Celdaobs  = ofv.GenCelda("",xclas & xsffun & " title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")
''             response.write BtipI & " 00001 "
           end if


    ''''  Celda = Celda & Celdaobs


            Formulario1 = Formulario1 & Celda1 & Celdaobs


         ELSE

            xclas   = "btbn"
            ''xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
            sHref   = "javascript:void(0);"
            sInput  = "<div class=" & xclas & " >" & BtipI & "</div>"
            lik  = ofv.GenLink("",xclas & " title='Propiedades de " & BtipP & "' ",sHref,"_blank","","",sInput,"si")
            Celda1  = ofv.GenCelda("",xclas & xsffun & " style='cursor:pointer;filter:shadow(color=#dddddd,direction=125);' title='Propiedades de " & BtipP & "' onclick='return false;' ","","","1","","",lik,"si")

''          response.write BtipI & " 00002 "





      ' estado observacion

            Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")


            if (Nestado=5  or Nestado=6) and (ars(42,inx) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & ars(42,inx)  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style=' border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

               Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")
      	       Celdaobs  = ofv.GenCelda("",xclas & xsffun & " title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")
''             response.write BtipI & " 00003 "
            end if


    ''''  Celda = Celda & Celdaobs



            Formulario1 = Formulario1 & Celda1  & Celdaobs

         END IF

''''''''''''''''
      else
        Celda = ofv.GenCelda("","ut colspan=2","","20","","","","&nbsp;","si")
        Celda = Celda & ofv.GenCelda("","ut ","","20","","","","&nbsp;","si")
        Formulario1 = Formulario1 & Celda
      end if


''''''''''''''''''''''''''''''''''''''''''''''
''''  Estructura de Edicion ''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''


      NWsec = ars(38,inx)

      NWsecR = split(NWsec,".")

      BsecD = ""
      for i=0 to ubound(NWsecR)-1
          BsecD = BsecD & (NWsecR(i)-900000) & "."
      next

      if  bsecix = 7 then
          bspan = ""
      else
          Bspan = "ut colspan=" & 7 - bsecix
      end if

      if bModificar then
         vlk = vuelta & "_pmk=" & amkey
         sAccionX = "asp/ActualizarEN.asp?ID=" & SCLAVE & "&Tabla=Objestdet&volver=" & vlk
         sAccionX = sAccionX & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      else
         sAccionX = ""
      end if



      if OSec <> "SI" then

        'Secuencia Oculta


         sInput = ofv.GenerarInput(nrs(4),ars(4,inx),"hidden","","2","dt","")
         sInput = sInput & ofv.GenerarInput(nrs(5),ars(5,inx),"hidden","","2","dt","")
         sInput = sInput & ofv.GenerarInput(nrs(6),ars(6,inx),"hidden","","2","dt","")


         Celda = ""
         Formulario2 = ofv.GenCelda("","","","1","","","",sInput & "&nbsp;","si")

      else

        'Secuencia Visible


         if bModificar then
            sInput = ofv.GenerarInput(nrs(4),ars(4,inx),"text",sForm & "X1","2","dtns","SI")
            sInput = ofv.GenForm(sForm & "X1","",sAccionx,"","post","",sInput,"si")
         else
            sInput = ofv.GenerarInput("",ars(4,inx),"readonly","","2","dtns","SI")
         end if

         sInput = sInput & ofv.GenerarInput(nrs(5),ars(5,inx),"hidden","","2","dt","")
         sInput = sInput & ofv.GenerarInput(nrs(6),ars(6,inx),"hidden","","2","dt","")

         Celda = ofv.GenCelda("","ut","","120","","","",BsecD & sInput,"si")
         Formulario2 = ""

      end if

    ''  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

      for i=1 to bsecix
      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      next



         select case stip
                case "A"
      sant = "A"
      'Descripcion

      sicK = "<i class='las la-angle-down'></i>&nbsp;"
      if bModificar then

        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","190","dts","")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","190","dts","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

           inx = inx + 1



                case "B"
      sant = "B"
      'Descripcion
      sicK = "<i class='las la-angle-down'></i>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","190","dts","")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","190","dts","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

           inx = inx + 1




                case "C"
      sant = "C"
      'Descripcion
      sicK = "<i class='las la-angle-right'></i>&nbsp;"
      if bModificar then

        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","90","dts","")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","90","dts","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

           inx = inx + 1


                case "D"
      sant = "D"
      'Descripcion
      sicK = "<i class='las la-angle-right'></i>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","90","dts","")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","90","dts","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

           inx = inx + 1

  '    set rsd = rs
  '    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "E"
      sant = "E"
      'Descripcion
      sicK = "<i class='las la-angle-right'></i>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","90","dts","")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","90","dts","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

           inx = inx + 1

    '  set rse = rs
    '  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

                case "F"
      sant = "F"
      'Descripcion
      sicK = "<i class='las la-angle-right'></i>&nbsp;"
      if bModificar then
        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","90","dts","")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","90","dts","SI")
      end if
      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sicK & sInput,"si")

           inx = inx + 1



                case "G"

      'Texto



      wstexto = replace(ars(8,inx)&"",chr(34),"'")


      if isnull(wstexto)  then
'        response.write "longv " & len(wstexto) & " " & wstexto & "<br>"
         sInput = space(80)
         sinput = replace(sinput," ","-")
      elseif len(wstexto) > 90 then
 '        response.write "longg " & len(wstexto) & " " & wstexto & "<br>"
         sinput = left(wstexto,90) & "<font size=1 color="& Session("CMFPNBOT") &" ><b>...(+)</b></font>"
      else
         if wstexto = space(len(wstexto)) then wstexto = replace(wstexto," ","-")
 '        response.write "longc " & len(wstexto) & " " & wstexto & "<br>"
         sinput = wstexto
      end if

    '  wstexto = "<font face='tahoma' size=2 color=#000080 >" & sInput  & "</font>"
      wstexto = sinput

      wstextoLNK = ars(32,inx) & ""

       if isnumeric(wstextoLNK) and wstextoLNK <> "0" then
          sicK = "<i class='las la-angle-right' ></i>&nbsp;"
       else
          sicK = "<i class='las la-angle-right'></i>&nbsp;"
       end if



      addqry   = "cantregmostrar=" & cantregmostrar & "&cantregmover=" & cantregmover
      addqry   = addqry & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec

      addqrySP = addqry

      addqry   = addqry & "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo


      sinput = "<a class=nt href=" & chr(34) & "edestrdetTXT.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sinput = sinput & "&IDTXT=" & ars(0,inx) & "&vuelve="
      sinput = sinput & replace(addqry,"&","_") & chr(34) & "   title='Ingresar/Modificar el texto' >" & sicK & wstexto & "</a>"

 '     sID = "edestrdetLnk.asp?"
 '     sID = sID & "IDTXT=" & ars(0,inx)
 '     sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&volver=edestrdet.asp?" & replace(addqrySP,"&","_")

         sID = "EdGENLnk.asp?"
         sID = sID & "ID=" & ars(0,inx) & "&TABLA=OBJESTDET&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaD & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo & "_" & replace(addqrySP,"&","_")



      sHref = sID

       if  isnumeric(wstextoLNK) and wstextoLNK <> "0" then
           sImagenP = "<i class='las la-link'></i>"'"<font face='corflaticon' size=2 color=#005500 >&#0252;</font>"
       else
           sImagenP="<i class='las la-link' ></i>"
         '  xclas  = "btbn"
          ' sImagenP ="<div class=" & xclas & " >" & sImagenP & "</div>"
           ' "<font face='corflaticon' size=2 color=#f7aa22 >&#0252;</font>"
       end if

''''   response.write sImagenP & " 0001 "


       sinputk = "<a class='ft1o'  href='" & sHref
       sinputk = sinputk & "' title='Actualizar link a otro documento' >" & sImagenP & "</a>"
       slink   = "&nbsp;" & sinputk


      sreduc = "<table cellpadding=2 cellspacing=0 ><tr >"
      sreduc = sreduc & "<td class=ut width=95% >" & sInput & "</td>"
      sreduc = sreduc & "<td class=ft >" & slink & "</td>"
      sreduc = sreduc & "</tr></table>"


      Celda = Celda & ofv.GenCelda("",bspan & " valign=middle  ","","","","","",sreduc,"si")

           inx = inx + 1


                case "H"

      'IMAGENES

      sicK = "&nbsp;<u><font face='corflaticon' size=5 color=#993333 >N</font></u>"
      if ars(9,inx) <> "" then
         estia=" style={font-size:7pt;font-family:tahoma,arial;color:#eeeeee;background-color:#000080;font-weight:600;border-style:outset;border-width:2;}"

         'sImagenP = "&#0232;"
         sinput = "<a href='Coedestimg.asp?A=" & ars(16,inx) & "&B=" & ars(15,inx)
         sinput = sinput & "&R=" & ars(12,inx) & "&S=" & server.urlencode(ars(9,inx))
         sinput = sinput & "' target=_blank  title='Ver Imagen'<i class='las la-arrow-right'></i></a>"
         slink  = sinput



      else
         slink = ""
      end if


     SCOMSEP = ""
     scomK   = ""

     For i=0 to 12
         scomK = scomK & SCOMSEP & nivel(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     if not  isnull(ars(12,inx)) then
        sreducD = ars(12,inx)
     else
        sreducD = 100
     end if



      if bModificar then

         sInput = ofv.gencomboFX(sForm & "X2",nrs(12),sreducD,scomK,scomD,"")
         sInputr1 = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
         if  isnull(ars(12,inx))   then
             sreduc = "100"
         else
             sreduc = ars(12,inx)
         end if

         sInputr1 = ofv.gencomboFX(sForm,nrs(12),sreducD,scomK,scomD,"no")

      end if

      scomK = "SI" & chr(9) & "NO"
      scomD = "SI" & chr(9) & "NO"

         if  isnull(ars(15,inx))   then
             sreducd = "NO"
         else
             sreducd = ars(15,inx)
         end if

      if bModificar then

         sInput = ofv.gencomboFX(sForm & "X3",nrs(15),sreducD,scomK,scomD,"")
         sInputr2 = ofv.GenForm(sForm & "X3","",sAccionx,"","post","",sInput,"si")

      else
         if isnull(ars(15,inx))  then
            sbord = "NO"
         else
             sbord = ars(15,inx)
         end if

         sInputr2 = ofv.gencomboFX(sForm,nrs(15),sreducD,scomK,scomD,"no")

      end if


     SCOMSEP = ""
     scomK   = ""

     For i=0 to 2
         scomK = scomK & SCOMSEP & nivel2(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     if not  isnull(ars(16,inx)) then
        sreducD = ars(16,inx)
     else
        sreducD = "IZQUIERDA"
     end if




      if bModificar then

         sInput = ofv.gencomboFX(sForm & "X4",nrs(16),sreducD,scomK,scomD,"")
         sInputr3 = ofv.GenForm(sForm & "X4","",sAccionx,"","post","",sInput,"si")

      else
         if isnull(ars(16,inx))  then
            salng = "IZQUIERDA"
         else
             salng =  ars(16,inx)
         end if

         sInputr3 = ofv.gencomboFX(sForm,nrs(16),sreducD,scomK,scomD,"no")

      end if



      wstexto = ars(9,inx)

      if wstexto <> "" then
         wstexto = chr(34) & wstexto & chr(34)
      else
         wstexto = "<font color='#dddddd'>--Sin&nbsp;Asociar--</font>"
      end if


      wstexto = "<u>" &  wstexto   & "</u>"

      addqry = "cantregmostrar=" & cantregmostrar & "&cantregmover=" & cantregmover
      addqry = addqry & "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec

      sinput = "<a class=nt href=" & chr(34) & "edestrdetIMG.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sinput = sinput & "&IDTXT=" & ars(0,inx) & "&vuelve="
      sinput = sinput & replace(addqry,"&","_") & chr(34) & " title='Asignar/Modificar Imagen' >" & wstexto & "</a>"



      sicK = "<i class='las la-angle-right'></i>&nbsp;"


      sreduc = "<table><tr>"
      sreduc = sreduc & "<td>" & sicK & "</td>"
      sreduc = sreduc & "<td >" & sInput & "&nbsp;</td>"
      sreduc = sreduc & "<td class=ft >" & slink & "</td>"
      sreduc = sreduc & "<td width=20% >&nbsp;</td>"
      sreduc = sreduc & "<td>" & sInputr1 & "</td>"
      sreduc = sreduc & "<td>" & sInputr2 & "</td>"
      sreduc = sreduc & "<td>" & sInputr3 & "</td>"
      sreduc = sreduc & "</tr></table>"


      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sreduc,"si")

           inx = inx + 1


               case "I"

      'Tablas
         estlink = "edTablas.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo
         estlink = estlink & "&OConte=SI&OConteK=" & ars(10,inx) & "&OConteD=" & ofv.convertircar(ars(7,inx)," ","_")
         estia=" style={font-size:7pt;font-family:tahoma,arial;color:#eeeeee;font-weight:600;text-decoration:none;background-color:#000080;border-style:outset;border-width:1pt;}"
         slink = "<a class=nn href=" & chr(34) & estlink & chr(34) & " target=_blank " & estia & " ><img src=imagenes/botedit.jpg border=0></a>"

      if bModificar then
        sInput = ofv.GenerarInput(nrs(7),ars(7,inx),"text",sForm & "X2","80","nn" & esttxt,"")
        sInput = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",ars(7,inx),"readonly","","80","nn" & esttxt,"SI")
      end if

      sicK = "<i class='las la-angle-right' ></i>&nbsp;"

      sCelda = sicK & sInput & "&nbsp;" & slink

      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sreduc,"si")


           inx = inx + 1


               case "J"

      'COMPONENTES





      sicK = "<i class='las la-angle-right' ></i>&nbsp;"


'''''''''''''''''




      'Componentes

        sComponente = "N"

        if ars(9,inx) <> "" then

          Strsql = "Select codobjetol,tipobjetol,* From objlnk Where ltipo = 'KL' and objid = 0" & ars(9,inx)
          set rsCompoM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsCompoM.eof then

              scompocod = rsCompoM(0)
              wtipoL   = rsCompoM(8)
	      wtipoLc  = rsCompoM(9) & ""

	       wNV1    = rsCompoM(13) & ""
	       wNV2    = rsCompoM(14) & ""
	       wNV3    = rsCompoM(15) & ""

	       IF  WNV1 = "" THEN WNV1 = "0"
	       IF  WNV2 = "" THEN WNV2 = "0"
	       IF  WNV3 = "" THEN WNV3 = "0"


              Strsql = "Select nombre,TIPOCOMP From csacomple Where codcomp = " & rsCompoM(0)
              set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              if  not rsComponentes.eof then
                  sComponente = rsComponentes(0)
                  sTIPOCMP    = rsComponentes(1)

              end if
              call ofv.cerrarconsulta(rsComponentes)

          end if


        end if

        sInput = sComponente

        if len(sInput) > 50 then
           sInput = left(sInput,47) & "..."
        end if


      wscompo =   scompocod
      if isnull(wscompo) then wscompo = "000"

      if sComponente = "N" then
         sInput1 = ""
         sInput  = "<U><font color=" & Session("CMFPNBOT") & " >ASOCIAR&nbsp;COMPONENTE&nbsp;</U></font>"
      else
         sInput1 = ""
         sInput  = "<u>" & replace(sInput," ","&nbsp;") & "</U>"
      end if


      if  bModificar then

          vlk   = vuelta & "_pmk=" & amkey
        addqry   = "cantregmostrar=" & cantregmostrar & "&cantregmover=" & cantregmover
        addqry   = addqry & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec

        addqrySP = addqry
      '    sID   = "EdEstrCMP.asp?"
      '    sID   = sID & "ID=" & SCLAVE & "&Codobjeto=" & nCodigo
      '    sID   = sID & "&Tipobjeto=" & sTipo & "&volver=" & vlk
         IF session("APLMDCLI")="BGAL" OR session("CORMODUFORM")="S"  THEN
         sID = "EdGENLnk.asp?base=K&"
         else

         sID = "EdGENLnkCOMP.asp?"
         end if
         sID = sID & "ID=" & ars(0,inx) & "&TABLA=OBJESTDET&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vueltaD & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo & "_" & replace(addqrySP,"&","_")
          IF session("APLMDCLI")="BGAL"  OR session("CORMODUFORM")="S" THEN
          slink = "<a class=nt href=" & chr(34) & sID & chr(34) & "   title='Asociar FORMULARIO' >" & sInput & "</a>"
          ELSE
          slink = "<a class=nt href=" & chr(34) & sID & chr(34) & "   title='Asociar componente' >" & sInput & "</a>"
          END IF
         ' response.write addqrySP
      else

         slink  = ""

     end if



      sreduc = ""

      'Imagen de Complemento
      if  ars(9,inx) <> "" then

          sImagenP = "&#0232;"
          scompocod =scompocod   & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

          sinput   = "<a class='ft1o'  href='EditarImagen2.asp?ID=" & scompocod
          sinput   = sinput & "' target='_new'  title='Visualizar el componente' >" & sImagenP & "</a>"
          slink2   = sinput

     SCOMSEP = ""
     scomK   = ""

     For i=0 to 12
         scomK = scomK & SCOMSEP & nivel(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     if not  isnull(ars(12,inx)) then
        sreducD = ars(12,inx)
     else
        sreducD = 100
     end if



      if bModificar then

         sInput = ofv.gencomboFX(sForm & "X2",nrs(12),sreducD,scomK,scomD,"")
        sInputr1 = ofv.GenForm(sForm & "X2","",sAccionx,"","post","",sInput,"si")

      else
         if  isnull(ars(12,inx))   then
             sreduc = "100"
         else
             sreduc = ars(12,inx)
         end if

         sInputr1 = ofv.gencomboFX(sForm,nrs(12),sreducD,scomK,scomD,"no")

      end if

      scomK = "SI" & chr(9) & "NO"
      scomD = "SI" & chr(9) & "NO"

         if  isnull(ars(15,inx))   then
             sreducd = "NO"
         else
             sreducd = ars(15,inx)
         end if

      if bModificar then

         sInput = ofv.gencomboFX(sForm & "X3",nrs(15),sreducD,scomK,scomD,"")
        sInputr2 = ofv.GenForm(sForm & "X3","",sAccionx,"","post","",sInput,"si")

      else
         if isnull(ars(15,inx))  then
            sbord = "NO"
         else
             sbord = ars(15,inx)
         end if

         sInputr2 = ofv.gencomboFX(sForm,nrs(15),sreducD,scomK,scomD,"no")

      end if


     SCOMSEP = ""
     scomK   = ""

     For i=0 to 2
         scomK = scomK & SCOMSEP & nivel2(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     if not  isnull(ars(16,inx)) then
        sreducD = ars(16,inx)
     else
        sreducD = "IZQUIERDA"
     end if




      if bModificar then

         sInput = ofv.gencomboFX(sForm & "X4",nrs(16),sreducD,scomK,scomD,"")
        sInputr3 = ofv.GenForm(sForm & "X4","",sAccionx,"","post","",sInput,"si")

      else
         if isnull(ars(16,inx))  then
            salng = "IZQUIERDA"
         else
             salng =  ars(16,inx)
         end if

         sInputr3 = ofv.gencomboFX(sForm,nrs(16),sreducD,scomK,scomD,"no")

      end if

      else
         sreduc = "&nbsp;"
      end if

      sreduc = "<table><tr>"
      sreduc = sreduc & "<td>" & sicK & "</td>"
      sreduc = sreduc & "<td>" & slink & "</td>"
      if sComponente = "N" then
         sreduc = sreduc & "<td width=20% colspan=4 >&nbsp;</td>"
      else
         sreduc = sreduc & "<td width=20% class=ft >" & slink2 & "</td>"
         sreduc = sreduc & "<td>" & sInputr1 & "</td>"
         sreduc = sreduc & "<td>" & sInputr2 & "</td>"
         sreduc = sreduc & "<td>" & sInputr3 & "</td>"
      end if
      sreduc = sreduc & "</tr></table>"


      Celda = Celda & ofv.GenCelda("",bspan,"","","","","",sreduc,"si")



''''''''''''''''''


           inx = inx + 1



      end select

  'Version
      sSelect = ""
      if bVersion = "SI" then
         if smodi then sSelect = "selected"
         if sSelect <> "" then
            if bModificar then


               if stipmodi = "A" then
                  sInput1 = "<b>A</b>"
               elseif stipmodi = "E" then
                  sInput1 = "<b>E</b>"
               elseif stipmodi = "M" then
                  sInput1 = "<b>M</b>"
               end if

               sInput = "<font class=at><select name=modi"
               sInput = sInput & " onchange=" & sForm & ".submit(); >"
               sInput = sInput & "<option value=False>No</option>"
               sInput = sInput & "<option value=True " & sSelect & ">" & sInput1
               sInput = sInput & "</option></select></font>"
            else
               sInput = ofv.GenerarInput("modi","Modificada","readonly","","10","dt","")
            end if

'            Celda = Celda & ofv.GenCelda("","dt","","","","","",sInput,"si")
        else
'            Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
        end if
      else
'        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

'      Celda = Celda & ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")

'''''''''

      sSelect = ""
    if bVersion = "SI" and smodi then


              if stipmodi = "A" then
                 stitle = "Agregado."
                 xclas  = "btbn"
                 'sInput = "<font class='btic'  >6</font>"
				 sInput = "<i class='las la-user-plus' ></i>"
              elseif stipmodi = "E" then
                 stitle = "Eliminado."
                 xclas  = "btbn"
                 'sInput = "<font class='btic'  >5</font>"
				 sInput = "<i class='las la-user-minus' ></i>"
              elseif stipmodi = "M" then
                 stitle = "Modificado."
                 xclas  = "btbn"
                 'sInput = "<font class='btic'  >4</font>"
				 sInput = "<i class='las la-user-edit' ></i>"
              end if

         sHref  = "javascript:void(0);"
         'xsffun = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
          ''   xsffun = xsffun & " onclick=""alert(document." & sform & ".all.modi);"" "
          '              xsffun =xsffun & document.getElementById('sform').modi;
         sInput = "<div class=" & xclas & " >" & sinput & "</div>"
         sInput = ofv.GenCelda("",xclas & " title='" & stitle & "'  " & xsffun & " ","","1%","","","",sInput,"si")

         Celda = Celda & ofv.GenLink("",xclas & " title='" & stitle & "' ",sHref,"","","",sInput,"si")''

         if bModificar then
            sInput = ofv.combosino(sForm,"1","modi")'form,valor,nombre'
            'sInput = REPLACE(sInput,"all.modi",sform & ".all.modi")
            sInput = REPLACE(sInput,"getElementById('modi')",  sform & ".modi")
            sInput = sInput & "<input type='hidden' value='"&ars(4,inx-1)&"' class='bt' name='"&nrs(4) &"' > "
            'sInput = sInput & "<input type='hidden' value='"&"ars(5,inx)"&"' class='bt' name='"&nrs(5) &"' > "
            'sInput = sInput & "<input type='hidden' value='"&"ars(6,inx)"&"' class='bt' name='"&nrs(6) &"' > "
            'sInput = sInput & "<input type='hidden' value='"&"ars(7,inx)"&"' class='bt' name='"&nrs(7) &"' > "
           'sInput = REPLACE(sInput,"getElementById(modi)",sform & ".getElementById(modi)")
           ' scomK  = "TRUE" & chr(9) & "FALSE"
           ' scomD  = "Si" & chr(9) & "No"
           ' sInput = ofv.gencomboFX(sForm,"modi","TRUE",scomK,scomD,"sb")
         else
            sInput = "&nbsp;"
         end if

         '' Fila = ofv.GenForm(sForm,"",sAccion,"","post","",sInput,"si") & CHR(13)
         '' Celda = Celda & ofv.GenCelda("","ut ","","","","","",Fila,"si")

          formx = ofv.GenForm(sForm,"",sAccionx,"","post","", sInput,"si") & CHR(13)

            Celda = Celda & ofv.GenCelda("","ut ","","","","","",formx,"si")

      else

          Celda = Celda & ofv.GenCelda("","ut colspan=2","","1%","","","","&nbsp;","si")

      end if

'''''''''

      Formulario = celda




      if wposi = cantregmostrar then
         Fila = ""
         exit for
      else

         Fila = ofv.GenRow("","aaa height=30px bgcolor=" & xTrColor & "","","","","","",Formulario1 & Formulario & Formulario2,"si")

         if  xTrColor = "#FFFFFF" then
             xTrColor = "#F9F9F9"
         ELSE
             xTrColor = "#FFFFFF"
         END IF

     '''    Fila = ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si") & CHR(13)
         response.write Fila
         wposi = wposi + 1
      end if




      amkeya = amkey


      inx = inx - 1


    next

         if inx > UBound(ars,2) then
            exit do

         end if

'    loop

         if wposi = cantregmostrar then
            exit do
         end if

      ssale = "no"
      do until ssale = "si"

         if     sant = "A" then
            sant = "X"
         elseif sant = "B" then
            sant = "A"
         elseif sant = "C" then
            sant = "B"
         elseif sant = "D" then
            sant = "C"
         elseif sant = "E" then
            sant = "D"
         elseif sant = "F" then
            sant = "E"
         end if

      if  sant = "X" then
         ssale = "si"
      end if

      loop


 loop

      response.flush


  celda = ""



      for i=1 to 10
         Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
      next

      if OSec = "SI" then
         Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")

      end if

''    Nuevo 10/02/2014
    '  Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
''    Nuevo 10/02/2014 Fin

      Celda = Celda & ofv.GenCelda("","","","79%","","","","&nbsp;","si")

      Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")


      if OSec <> "SI" then
         Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
      end if


      Fila = ofv.GenRow("","","","","","","",celda,"si")
      response.write Fila

response.write ofv.GenFTabla()



addqry = "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo & "&ID=" & ID & "&pmk=" & pmk & "&OSec=" & OSec

response.write  "<TABLE align=center  name=btnas2 id=btnas2 width=2%  marginwidth=0  marginheight=0  cellspacing=1  cellpadding=2 ><tr>"

if cantregmover > 0 then
wcantmov = cantregmover - cantregmostrar
response.write  "<TD ><FORM  name=Bretro ID=Bretro action="
response.write  chr(34) & "edestrdet.asp?cantregmostrar=" & cantregmostrar & "&CantRegMover=" & wcantmov & addqry & chr(34) & "  method=post >"
response.write  "<input  type=SUBMIT  size=10%  value='Anterior' class=bt >"
response.write  "</FORM></TD>"
end if

wposi2 = (UBound(ars,2) + 1) - wposi

if (not (wposi < cantregmostrar)) and wposi2 > 0  then
wcantmov = cantregmover + cantregmostrar
response.write  "<TD ><FORM  name=Bavan ID=Bavan action="
response.write  chr(34) & "edestrdet.asp?cantregmostrar=" & cantregmostrar & "&CantRegMover=" & wcantmov & addqry & chr(34) & "  method=post >"
response.write  "<input  type=SUBMIT  size=10%  value='Siguiente' class=bt >"
response.write  "</FORM></TD>"
end if

response.write  "</tr></table>"


  Agregar = ""
  Buscar = ""

  'Boton Volver
  sAccion   = "edestr.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sRetornar = ofv.BotonVolver(sAccion)

  Buscar    = ""

 if  ars(39,0) <> "N" then

  'Boton test
    'sAccion   = "CoestrCons.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&sdocum=" & ID & "&origen=ED"
     sAccion= " Consultar.asp?Codigo=" & nCodigo & "&Tipo=" & STIPO &"&niv1=11&niv2=" & ID & "&test=" & sTest
     Buscar    = ofv.BotonTest(sAccion)
 end if

  xsfadd    = "&ID=" & ID & "&OSec=" & OSec




  sHTML = ""
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmover,cantregmostrar,"no","no",xsfii,nCodigo,sTipo & xsfadd,"",Agregar,Buscar,sRetornar)

response.write sHTML

 call ofv.cerrarconn(cn)

response.write GTCOBS()
response.write xxform
 sHTML =  "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
    'sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }</script>"
    response.write shtml
    response.write  "</body></html>"
if  session("APLMDCLI") = "BGAL" OR session("APLMDCLI") = "BANELCO" OR session("APLMDCLI") = "BCR"  then

response.write  "</body><script>change_span('',"& jsHashPefiles &",'spanpf');change_span('',"& jsHashsectores &",'spansec');</script></html>"
else
response.write  "</body></html>"
end if

      response.flush

end if






Public Function BotonEliminar(sNombre, sHref,CLAVE)

             wktitdelete="Confirma la eliminacion "
           '' sImageELS = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
           '' sImageels = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
             sImageELS  = "<input  type=BUTTON  size=10%  value='&#xf2ed;' class=bte     >"
             sformobj   = "FRMOBJ" & CLAVE
             sAccion    = sHref
             formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
             sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

            BotonEliminar=ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs ,"si")
            xxform=xxform & formobj


End Function

%>