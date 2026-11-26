<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")



  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaaspEx("../","0","0")'  &"_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
 ' vuelta = ofv.vueltaaspEx("../","0","0")
  'vuelta=ofv.vueltaasp("",0,0)  &"&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
'response.write vuelta
  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

%>
<!-- #INCLUDE FILE="IncFile/CorVinculadorPL.inc" -->
<%


  IF WKSIGUE = "S" THEN

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  if session("Nomodificar") = "S" then bModificar = False

   if session("Nomodificar") = "S" then bSuprimir = False



  'Encabezado sector

 dim STab (4,5)


 for yyy=1 to 4

      for zzz=1 to 5

           STab (yyy,zzz) = "N"

      next

 next


  Strsql = "SELECT OBJportada.detcon, OBJportada.detsecuen, OBJportada.dettipo,OBJID from OBJportada "
  Strsql = Strsql & " where OBJportada.codobjeto=" & nCodigo & " and OBJportada.detcon > 0 "
  Strsql = Strsql & " order by OBJportada.detcon, OBJportada.detsecuen "
 '  response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

           yyy = rs(0)
           zzz = rs(1)
           STab (yyy,zzz) = rs(2) & ""
           ID=RS(3)
   rs.movenext
  loop

  Formulario = ""
  Fila= ""

  for yyy=1 to 4

      Celda = ""

      for zzz=1 to 5

       sclase = "bt"
       sclaseC = "Ut style='border-style:solid;border-width:1;border-color:dddddd;' "

       if    STab (yyy,zzz) = "S" then
           sclaseC = "UT bgcolor='#f7eedd' STYLE='border-style:outset;border-width:3;border-color:#DDCCAA;' "
       end if

        sInput = ofv.GenerarInput("",yyy & "-" & zzz ,"submit","","90",sclase,"")


        ''  Celdax = ofv.GenCelda("",sclaseC & " height=40px","center ","10%","","","",sInput,"si")




          if bModificar then
             sAccion = "edPortadaDet.asp?F=" & yyy & "&C=" & zzz
             sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
              formu= ofv.GenForm(sID,"",sAccion,"","post","",sinput,"si")
           ''  Celda = Celda & ofv.GenForm(sID,"",sAccion,"","post","",Celdax,"si")
          end if

          Celda = celda & ofv.GenCelda("",sclaseC & " height=40px","center ","10%","","","",formu,"si")

      next

        'Celdax = ofv.GenCelda("",sclaseC & " height=40px","center ","10%","","","",formu,"si")
      fila = fila & ofv.GenRow("","","","","","","",Celda,"si")





  next


   if bSuprimir then
        sHref = "asp/Eliminaren.asp?ID="&ID&"&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sHref = sHref & "&Tabla=objportada&volver=" & vuelta & "&Eliminar=NSNC"
        'Celdas = ofv.BotonEliminar("eliminar",sHref)
        sInput = ofv.GenerarInput("","Eliminar Links","submit","","10","bt","")
	'CeldaELIM = ofv.GenCelda("","","center COLSPAN=5","40%","","","",sInput,"si")
        FormularioELIM = ofv.GenForm("Mensajes","",sHref,"","post","",sInput,"si")

      else
        FormularioELIM =  "&nbsp;"'' ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    CeldaELIM = ofv.GenCelda("","","center COLSPAN=5","40%","","","",FormularioELIM,"si")

    fila1= ofv.GenRow("","","","","","","",CeldaELIM,"si")
    'Formulario =  ofv.GenRow("","","","","","","",FormularioELIM,"si")
   ' Formulario =  ofv.GenRow("","","","","","","",FormularioELIM,"si") & FORMULARIO

  sHTML = sHTML & "<br><center>" & ofv.GenTabla("","aaa","","50%","","0","","0","0","","",fila1 & fila,"si") & "</center><br>"





  strsqlx1 = "SELECT * from Objportada where  OBJportada.codobjeto=" & nCodigo & " and  OBJportada.detcon = 0 and OBJportada.detsecuen = 99"
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSqlx1,cn,1,parametros)

  if  rs.eof then

             strsql = "insert into Objportada  ("
             strsql = strsql & "codobjeto,tipobjeto,"
             strsql = strsql & "detcon,"
             strsql = strsql & "detsecuen,"
             strsql = strsql & "dettipo,"
             strsql = strsql & "detdepen,"
             strsql = strsql & "dettitulo,"
             strsql = strsql & "dettexto,"
             strsql = strsql & "detimagen,"
             strsql = strsql & "dettabla,"
             strsql = strsql & "detfont,"
             strsql = strsql & "detfsize,"
             strsql = strsql & "detfcolor,"
             strsql = strsql & "detbgcolor,"
             strsql = strsql & "detborde,"
             strsql = strsql & "detftipo,"
             strsql = strsql & "detsang,"
             strsql = strsql & "detindent1,"
             strsql = strsql & "detindent2,"
             strsql = strsql & "detvin,"
             strsql = strsql & "detvintipo,"
             strsql = strsql & "detvinindent,"
             strsql = strsql & "detnumer,"
             strsql = strsql & "detnumertipo,"
             strsql = strsql & "detnumerclase,"
             strsql = strsql & "detvinfont,"
             strsql = strsql & "detalign,"
             strsql = strsql & "dettextalign,"
             strsql = strsql & "modi,"
             strsql = strsql & "tipmodi,"
             strsql = strsql & "detdefault,"
             strsql = strsql & "rellink,"
             strsql = strsql & "reflink,"
             strsql = strsql & "relant,"
             strsql = strsql & "modautor,"
             strsql = strsql & "modfecha,"
             strsql = strsql & "modhora"

             strsql = strsql & ") values ("
             strsql = strsql & nCodigo & ",'" & sTIPo & "',"
             strsql = strsql & "'0',"
             strsql = strsql & "'99',"
             strsql = strsql & "'N',"
             strsql = strsql & "'0',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'1',"
             strsql = strsql & "'',"
             strsql = strsql & "'SI',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'" & SESSION("USUARIO") & "',"
             strsql = strsql & "'" & WKFECHA & "',"
             strsql = strsql & "'" & WKHORA & "'"

             strsql = strsql & ")"
             'call ofv.crearconsultaEx(StrSql,cn,1,volver)
             xxx=SqlExLog("INSERT",strsql,CN,"Objportada","","","")


  strsqlx1 = "SELECT * from Objportada where  OBJportada.codobjeto=" & nCodigo & " and  OBJportada.detcon = 0 and OBJportada.detsecuen = 99"
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSqlx1,cn,1,parametros)


  end if



      sEst     = "TEXTO"
      sDettipo = "G"
      sPart    = "SI"



dim atit(10)
titc = 0
atit(1)  = "TITULO"
atit(2)  = "SUBTITULO 1"
atit(3)  = "SUBTITULO 2"
atit(4)  = "SUBTITULO 3"
atit(5)  = "SUBTITULO 4"
atit(6)  = "SUBTITULO 5"
atit(7)  = "TEXTO"
atit(8)  = "NO"
atit(9)  = "NO"
atit(10) = "NO"

  dim aFName(18)
  dim aFValue(18)
  dim aFValueD(18)
  dim aFTitle(18)
  dim aInput(18)
  dim aInputD(18)

  aFname(0) = "detfont"
  aFname(1) = "detfsize"
  aFname(2) = "detfcolor"
  aFname(3) = "detbgcolor"
  aFname(4) = "detborde"
  aFname(5) = "detftipo"
  aFname(6) = "detsang"
  aFname(7) = "detindent1"
  aFname(8) = "detindent2"
  aFname(9) = "detvin"
  aFname(10) = "detvinfont"
  aFname(11) = "detvintipo"
  aFname(12) = "detvinindent"
  aFname(13) = "detnumer"
  aFname(14) = "detnumertipo"
  aFname(15) = "detnumerclase"
  aFname(16) = "detalign"
  aFname(17) = "dettextalign"
  aFname(18) = "detdefault"


  aFtitle(0) = "Fuente"
  aFtitle(1) = "Tamaño Fuente"
  aFtitle(2) = "Color Fuente"
  aFtitle(3) = "Color Fondo"
  aFtitle(4) = "Borde"
  aFtitle(5) = "Tipo"
  aFtitle(6) = "Sangria"
  aFtitle(7) = "Indentación"
  aFtitle(8) = "Espaciado después del bloque"
  aFtitle(9) = "Viñeta"
  aFtitle(10) = "Fuente de Viñeta"
  aFtitle(11) = "Tipo de Viñeta"
  aFtitle(12) = "Espaciado de Viñeta"
  aFtitle(13) = "Numeración"
  aFtitle(14) = "Tipo de Numeración"
  aFtitle(15) = "Estilo de Numeración"
  aFtitle(16) = "Alineación"
  aFtitle(17) = "Alineación del Texto"
  aFtitle(18) = "Default"


  sEsquema = "S/D"
  scodiden = "S/D"
  sesqnom = "S/D"

  sform = "Propiedades"




    Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then   sEsquema = rs(1)

    Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
       scodiden = rs(1)
       sesqnom = rs(3)
    end if


    strsqlx = "Select dettipo as a0,dettipo as a0,dettipo as a0,dettipo as a0, SYSESTPRTY.* From SYSESTPRTY Where "
    strsqlx = strsqlx & " Dettipo = '" & sDettipo & "'"
    nTop = 17

 ' REPONSE.WRITE STRSQLX




     set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
     if  not rs.eof then
         for n = 0 to nTop
             aFValueD(n) = rs(aFName(n))
         next
         aFValueD(18) = "SI"
     end if
     call ofv.cerrarconsulta(rs)


  '   strsql = "Select * From Objestprty Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
  '   strsql = strsql & sTipo & "' And Detcon = " & ID2 & " And Dettipo = '" & sDettipo & "'"

     strsql = strsqlx1
   '  response.write strsql
   '  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   '  if  not rs.eof then
   '      for n = 0 to nTop
   '          if not isnull(rs(aFName(n))) then aFValueD(n) = rs(aFName(n))
   '      next
   '  end if
   '  call ofv.cerrarconsulta(rs)



   ' strsql = "Select * From Objestdet Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
  '  strsql = strsql & sTipo & "' And Objid = " & ID


     strsql = strsqlx1
    nTop = 18






  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      nID = rs(0)
      for n = 0 to nTop
          if not isnull(rs(aFName(n))) then aFValue(n) = rs(aFName(n))
      next
      if  aFValue(18) = "" Then aFValue(18) = "SI"


  end if
  call ofv.cerrarconsulta(rs)



  dim aSiNo(1)
  aSiNo(0) = "NO"
  aSiNo(1) = "SI"




'Default





    sDef = aFValue(18)

     SCOMSEP = ""
     scomK   = ""

     For n=0 to Ubound(aSiNo)
         scomK = scomK & SCOMSEP & aSiNo(n)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     aInput(18) = ofv.gencomboFX(sForm & "18",aFname(18),aFValue(18),scomK,scomD,"")





''''''' Armado de estructuras de seleccion de opciones '''''''''

  dim  aSpacing(10)
  for  i = 0 to 10
       aSpacing(i) = CStr(i)
  next

  dim  aVSpacing(4)
  for  i = 0 to 4
       aVSpacing(i) = CStr(i)
  next

'Font Size
  dim aFontSize(13)
  aFontSize(0) = ""
  aFontSize(1) = "7"
  aFontSize(2) = "8"
  aFontSize(3) = "9"
  aFontSize(4) = "10"
  aFontSize(5) = "11"
  aFontSize(6) = "12"
  aFontSize(7) = "14"
  aFontSize(8) = "16"
  aFontSize(9) = "18"
  aFontSize(10) = "20"
  aFontSize(11) = "22"
  aFontSize(12) = "24"
  aFontSize(13) = "36"

'Border
  dim aBorder(5)
  aBorder(0) = "Ninguno"
  aBorder(1) = "Recuadro"
  aBorder(2) = "Elevado"
  aBorder(3) = "Bajorrelieve"
  aBorder(4) = "Marco 3D (1)"
  aBorder(5) = "Marco 3D (2)"
  dim aBorder2(5)
  aBorder2(0) = ""
  aBorder2(1) = "Solid"
  aBorder2(2) = "Outset"
  aBorder2(3) = "Inset"
  aBorder2(4) = "Ridge"
  aBorder2(5) = "Groove"

'Font de Viñeta
  dim aFontName2(3)
  aFontName2(0) = ""
  aFontName2(1) = "corflaticon"
  aFontName2(2) = "Symbol"
  aFontName2(3) = "corflaticon"

'Tipo de Viñeta
  dim aTipoVin(222)
  j = 0
  for i = 33 to 255
      aTipoVin(j) = Chr(i)
      j = j + 1
  next

'Tipo de Numeracion
  dim aTipoNum(5)
  aTipoNum(0) = "1"
  aTipoNum(1) = "I"
  aTipoNum(2) = "i"
  aTipoNum(3) = "A"
  aTipoNum(4) = "a"
  aTipoNum(5) = "NO"

'Clase de Numeracion
  dim aClaseNum(1)
  aClaseNum(0) = "Simple"
  aClaseNum(1) = "Compuesta"

'Alineacion
  dim aAlign(4)
  aAlign(0)  = ""
  aAlign(1)  = "IZQUIERDA"
  aAlign(2)  = "CENTRADA"
  aAlign(3)  = "DERECHA"
  aAlign(4)  = "JUSTIFICADA"

  dim aAlign2(4)
  aAlign2(0) = ""
  aAlign2(1) = "left"
  aAlign2(2) = "center"
  aAlign2(3) = "right"
  aAlign2(4) = "justify"


set ofr = server.createobject("corscr.Corrender")

ofr.verfonts


'Font Name
  redim preserve aFontName(0)
  aFontName(0) = ""

for ii=0 to ofr.wscreen.fontcount -1
    redim preserve aFontName(ii + 1)
    aFontName(ii + 1) = ofr.wscreen.fonts(ii)
next




     SCOMSEP = ""
     scomK   = ""

     For n=0 to Ubound(aFontName)
         scomK = scomK & SCOMSEP & aFontName(n)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     nn = 0

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(0) = ofv.gencomboFX(sForm & nn,aFname(0),xvalor,scomK,scomD,"sb")


'' Font-Size

     SCOMSEP = ""
     scomK   = ""

     For n=0 to Ubound(aFontSize)
         scomK = scomK & SCOMSEP & aFontSize(n)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     nn = 1

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(1) = ofv.gencomboFX(sForm & nn,aFname(1),xvalor,scomK,scomD,"sb")




'Font Color
  strsql = "Select * From Syscolor"


     nn = 2

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

  aInput(2) = ofv.GenerarCombo(sForm & nn,aFName(2),xvalor,Strsql,cn,2,1,"sb")

'Background Color

     nn = 3

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

  aInput(3) = ofv.GenerarCombo(sForm & nn,aFName(3),xvalor,Strsql,cn,2,1,"sb")

'' Border

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aBorder)
         scomK = scomK & SCOMSEP & aBorder2(n)
         scomD = scomD & SCOMSEP & aBorder(n)
         SCOMSEP = chr(9)
     next

   ''  scomD = scomK

     nn = 4

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(4) = ofv.gencomboFX(sForm & nn,aFname(4),xvalor,scomK,scomD,"sb")

'Tipo

     nn = 5

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if


  sValue = xvalor
  n = instr(sValue,"1")
  sChecked = ""
'  response.write n & "---" & sValue & "<br>"
  if n > 0 then sChecked = " checked"
 ' sChecked = " checked"
  aInput(5) = "<input type=checkbox value=1 name='" & aFName(5) & "a'" & sChecked
  aInput(5) = aInput(5) & " onclick='document." & sForm & nn & ".submit();'>Bold"
  sChecked = ""
  n = instr(sValue,"2")
 ' response.write n & "---" & sValue & "<br>"
  if n > 0 then sChecked = " checked"
  aInput(5) = aInput(5) & "<input type=checkbox value=2 name='" & aFName(5) & "b'" & sChecked
  aInput(5) = aInput(5) & " onclick=document." & sForm & nn & ".submit();>Italic"
  sChecked = ""
  n = instr(sValue,"4")
'  response.write n & "---" & sValue & "<br>"
  if n > 0 then sChecked = " checked"
  aInput(5) = aInput(5) & "<input type=checkbox value=4 name='" & aFName(5) & "c'" & sChecked
  aInput(5) = aInput(5) & " onclick=document." & sForm & nn & ".submit();>Underline"
  sChecked = ""
  n = instr(sValue,"8")
'  response.write n & "---" & sValue & "<br>"
  if n > 0 then sChecked = " checked"
  aInput(5) = aInput(5) & "<input type=checkbox value=8 name='" & aFName(5) & "d'" & sChecked
  aInput(5) = aInput(5) & " onclick=document." & sForm & nn & ".submit();>Strikeout"





'Sangria
  if sEst = "TEXTOx" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aSpacing)
         scomK = scomK & SCOMSEP & aSpacing(n)
         scomD = scomD & SCOMSEP & aSpacing(n)
         SCOMSEP = chr(9)
     next

     nn = 6

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(6) = ofv.gencomboFX(sForm & nn,aFname(6),xvalor,scomK,scomD,"sb")



  else
    aInput(6) = "NO"
  end if



'Indent 1
  if sEst = "TITULO"  And sPart <> "SI" then
     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aSiNO)
         scomK = scomK & SCOMSEP & aSiNO(n)
         scomD = scomD & SCOMSEP & aSiNO(n)
         SCOMSEP = chr(9)
     next

     nn = 7

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(7) = ofv.gencomboFX(sForm & nn,aFname(7),xvalor,scomK,scomD,"sb")


  else
    aInput(7) = "NO"
  end if

'Espaciado vertical
  if sPart <> "SI" then
     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aVSpacing)
         scomK = scomK & SCOMSEP & aVSpacing(n)
         scomD = scomD & SCOMSEP & aVSpacing(n)
         SCOMSEP = chr(9)
     next

     nn = 8

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(8) = ofv.gencomboFX(sForm & nn,aFname(8),xvalor,scomK,scomD,"sb")


  else
    aInput(8) = "NO"
  end if

'Viñeta
  if sEst <> "TEXTO" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aSiNO)
         scomK = scomK & SCOMSEP & aSiNO(n)
         scomD = scomD & SCOMSEP & aSiNO(n)
         SCOMSEP = chr(9)
     next

     nn = 9

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(9) = ofv.gencomboFX(sForm & nn,aFname(9),xvalor,scomK,scomD,"sb")


  elseif sEst = "TEXTOx" then




     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aSiNO)
         scomK = scomK & SCOMSEP & aSiNO(n)
         scomD = scomD & SCOMSEP & aSiNO(n)
         SCOMSEP = chr(9)
     next

     nn = 9

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(9) = ofv.gencomboFX(sForm & nn,aFname(9),xvalor,scomK,scomD,"sb")


  else
    aInput(9)  = "NO"
    aFValue(9) = "NO"
    xvalor     = "NO"
  end if

'' Font de Viñetas

  if sEst <> "TEXTO" and xvalor = "SI" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aFontName2)
         scomK = scomK & SCOMSEP & aFontName2(n)
         scomD = scomD & SCOMSEP & aFontName2(n)
         SCOMSEP = chr(9)
     next

     nn = 10

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(10) = ofv.gencomboFX(sForm & nn,aFname(10),xvalor,scomK,scomD,"sb")


'' Tipo de Viñetas

    if xvalor <> "" then

       xvalor10 = xvalor

     nn = 11

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

      aInput(11) = "<font face='" & xvalor10 & "' size=2><select "
      aInput(11) = aInput(11) & " style={font-family:'"
      aInput(11) = aInput(11) & xvalor10 & "';font-size:14pt;}"
      aInput(11) = aInput(11) & " name=" & aFname(11)
      aInput(11) = aInput(11) & " onchange=document." & sForm & nn & ".submit();>"
      For n = 0 To Ubound(aTipoVin)
        sSelect = ""
        If aTipoVin(n) = xvalor Then sSelect = " selected"
        aInput(11) = aInput(11) & "<option value='" & aTipoVin(n) & "' " & sSelect & ">"
        aInput(11) = aInput(11) & aTipoVin(n) & "</option>"
      Next
      aInput(11) = aInput(11) & "</select></font>"
    else
      aInput(11) = "NO"
    end if


    'Espaciado de Viñeta

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aSpacing)
         scomK = scomK & SCOMSEP & aSpacing(n)
         scomD = scomD & SCOMSEP & aSpacing(n)
         SCOMSEP = chr(9)
     next

     nn = 12

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(12) = ofv.gencomboFX(sForm & nn,aFname(12),xvalor,scomK,scomD,"sb")

  else
     aInput(10) = "NO"
     aInput(11) = "NO"
     aInput(12) = "NO"
  end if

    xvalor     = "NO"

'Numeracion
  if sEst = "TITULO"  And sPart <> "SI" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aSiNo)
         scomK = scomK & SCOMSEP & aSiNo(n)
         scomD = scomD & SCOMSEP & aSiNo(n)
         SCOMSEP = chr(9)
     next

     nn = 13

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(13) = ofv.gencomboFX(sForm & nn,aFname(13),xvalor,scomK,scomD,"sb")

  else
    aInput(13) = "NO"
    xvalor     = "NO"

  end if

''Tipo de Numeracion
'  if xvalor <> "NO" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aTipoNum)
         scomK = scomK & SCOMSEP & aTipoNum(n)
         scomD = scomD & SCOMSEP & aTipoNum(n)
         SCOMSEP = chr(9)
     next

     nn = 14

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(14) = ofv.gencomboFX(sForm & nn,aFname(14),xvalor,scomK,scomD,"sb")

     aInput(14) = "NO"


 '   if sEst = "TITULO" then
''Clase de Numeracion
  if xvalor <> "NO" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aClaseNum)
         scomK = scomK & SCOMSEP & aClaseNum(n)
         scomD = scomD & SCOMSEP & aClaseNum(n)
         SCOMSEP = chr(9)
     next

     nn = 15

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(15) = ofv.gencomboFX(sForm & nn,aFname(15),xvalor,scomK,scomD,"sb")

    else
        aInput(15) = "NO"
    end if

'  else
'      aInput(14) = "NO"
'      aInput(15) = "NO"
'  end if

     aInput(14) = "NO"

'Alineacion

  if sEst <> "TEXTO" then

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aAlign)
         scomK = scomK & SCOMSEP & aAlign2(n)
         scomD = scomD & SCOMSEP & aAlign(n)
         SCOMSEP = chr(9)
     next

     nn = 16

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(16) = ofv.gencomboFX(sForm & nn,aFname(16),xvalor,scomK,scomD,"sb")

  else
      aInput(16) = "NO"
  end if

'Alineacion Texto

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For n=0 to Ubound(aAlign)
         scomK = scomK & SCOMSEP & aAlign2(n)
         scomD = scomD & SCOMSEP & aAlign(n)
         SCOMSEP = chr(9)
     next

     nn = 17

     if stipo = "SYSTEM" then
        aInputD(nn) = "N"
        xvalor     =  aFValue(nn)
     else
        if aFValue(nn) <> "" then
           xvalor     =  aFValue(nn)
           aInputD(nn) = "N"
        else
           xvalor     =  aFValueD(nn)
           aInputD(nn) = "S"
        end if
     end if

     aInput(17) = ofv.gencomboFX(sForm & nn,aFname(17),xvalor,scomK,scomD,"sb")






  if  sPart = "SI" then
      sInput = "Particulares"
  else
      sInput = "Generales"
  end if


'''''''''''''''''''''''' lista de opciones ''''''''''''''''''


     sHTMLa   = ""




''''''''''''' Panel de Propiedades

  sImagenDF = "<FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT>"
  sImagenNM = "<FONT FACE=corflaticon COLOR=#dddddd SIZE=2>4</FONT>"


  xTrColor = "#FFFFFF"

  bstitulo = "<h5>Propiedades </h5>"

  Celda    = ofv.GenCelda("","vt colspan=3 ","","","","","",bstitulo,"si")
  fila=""
  Fila     = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


  Celda    = ofv.GenCelda("","vt","","40%","","","","---------------------------","si")
  sInput   = ofv.GenerarInput("","Actualizar","submit","","10","bt","")

  sInput   = sHTMLa

  Celda    = Celda & ofv.GenCelda("","ut colspan=2 ","right","60%","","","",sInput,"si")
  Fila     = Fila & ofv.GenRow("","","","","","","",Celda,"si")



  sAccion = ""
 ' response.write sPart
  if bModificar then

        if sPart = "SI" then
          sAccion = "asp/ActualizarPrtyPortada.asp?ID=" & nID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        end if

  end if





  if  sPart  = "SI" then
      Celda  = ofv.GenCelda("","ut","","","","","",aFtitle(18),"si")
      sInput = aInput(18)

      sInput = ofv.GenForm("Propiedades18","",sAccion,"","post","",sInput,"si")

      Celda  = Celda & ofv.GenCelda("","uti colspan=2 ","","","","","",sInput,"si")
      Fila   = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

      if  xTrColor = "#FFFFFF" then
          xTrColor = "#F9F9F9"
      ELSE
          xTrColor = "#FFFFFF"
      END IF

  end if



  if  sDef = "NO" then
      for n = 0 to UBound(aFName) - 1
          if  aInput(n) <> "NO" and n <> 5 then
              Celda = ofv.GenCelda("","uti","","","","","",aFtitle(n) ,"si")
              if  aInput(n) = "" then
                  sInput = ofv.GenerarInput(aFname(n),aFvalue(n),"text","","80","dt","")
              else
                  sInput = aInput(n)
              end if

              sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",sInput,"si")

              if  aInputD(n) = "N" then
                  sclasx  = " style='border-style:solid;border-width:1;border-color:#993333;' "
                  sInputf = sImagenDF

              else
                  sclasx  = ""
                  sInputf = sImagenNM
              end if

              Celda = Celda & ofv.GenCelda("","ut","center","1%","","","",sInputf,"si")

              Celda = Celda & ofv.GenCelda("","uti " & sclasx & " ","","","","","",sInput,"si")

            '''''''''''''''  sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",Celda,"si")

              Fila  = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

              if  xTrColor = "#FFFFFF" then
                  xTrColor = "#F9F9F9"
              ELSE
                  xTrColor = "#FFFFFF"
              END IF


          end if
      next

          if  aInput(5) <> "NO" then
              n=5
              Celda = ofv.GenCelda("","uti","","","","","",aFtitle(n),"si")
              if  aInput(n) = "" then
                  sInput = ofv.GenerarInput(aFname(n),aFvalue(n),"text","","80","dt","")
              else
                  sInput = aInput(n)
              end if

              sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",sInput,"si")

              if  aInputD(n) = "N" then
                  sclasx  = " style='border-style:solid;border-width:1;border-color:#993333;' "
                  sInputf = sImagenDF

              else
                  sclasx  = ""
                  sInputf = sImagenNM
              end if

              Celda = Celda & ofv.GenCelda("","ut","center","1%","","","",sInputf,"si")

              Celda = Celda & ofv.GenCelda("","uti " & sclasx & " ","","","","","",sInput,"si")

       ''''''''''''''       sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",Celda,"si")

              Fila  = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

              if  xTrColor = "#FFFFFF" then
                  xTrColor = "#F9F9F9"
              ELSE
                  xTrColor = "#FFFFFF"
              END IF


          end if

  end if


  Formulario = Fila

  sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","0","0","","",Formulario,"si")




 end if




'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
 sAccion = "Coportada.asp?skind=1&tipobjeto=" & stipo &"&codobjeto=" & ncodigo'&"&genlnk=NO"
 listar =  "<td><input  type=BUTTON  size=10%  value='Test' class=bietx  onclick=""document.getElementById('TestFrm').submit();""   ></td>"
 listar= listar & ofv.genform("TestFrm","AAA STYLE='display:none;' ",sAccion,"","post target=_self","","","si")

sHTML = sHTML & ofv.botoneraESSX("0","",CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")

''  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</body></html>"
end if



response.write sHTML

%>