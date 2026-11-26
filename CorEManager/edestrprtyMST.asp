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

  volver  = request.querystring("volver")
  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sEst    = request.querystring("Tipo")
  sPart   = request.querystring("Part")
  ID      = request.querystring("ID")
  ID2      = request.querystring("ID2")

  if  sEst = "" then sEst = "TITULO"

  sEst    = UCase(sEst)

  select case sEst
    case "TITULO"
      sDettipo = "A"
    case "SUBTITULO1"
      sDettipo = "B"
    case "SUBTITULO2"
      sDettipo = "C"
    case "SUBTITULO3"
      sDettipo = "D"
    case "SUBTITULO4"
      sDettipo = "E"
    case "SUBTITULO5"
      sDettipo = "F"
    case "TEXTO"
      sDettipo = "G"

    case "A"
      sEst     = "TITULO"
      sDettipo = "A"
    case "B"
      sEst     = "SUBTITULO1"
      sDettipo = "B"
    case "C"
      sEst     = "SUBTITULO2"
      sDettipo = "C"
    case "D"
      sEst     = "SUBTITULO3"
      sDettipo = "D"
    case "E"
      sEst     = "SUBTITULO4"
      sDettipo = "E"
    case "F"
      sEst     = "SUBTITULO5"
      sDettipo = "F"
    case "G"
      sEst     = "TEXTO"
      sDettipo = "G"

  end select


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

  sHTML = ofv.FormHeader(session("FormN"))



  set cn = ofv.conectar(ofv.strconn2)
  sEsquema = "S/D"
  scodiden = "S/D"
  sesqnom = "S/D"

  sform = "Propiedades"

  if stipo = "SYSTEM" then

    sEsquema = "DEFAULT"
    scodiden = ""
    sesqnom  = "SISTEMA"

    strsql = "Select dettipo as a0,dettipo as a0,dettipo as a0,dettipo as a0, SYSESTPRTY.* From SYSESTPRTY Where "
    strsql = strsql & " Dettipo = '" & sDettipo & "'"
    nTop = 17

  else

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

    Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then   sEsquema = rs(1)

    Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then
       scodiden = rs(1)
       sesqnom = rs(3)
    end if


    strsql = "Select dettipo as a0,dettipo as a0,dettipo as a0,dettipo as a0, SYSESTPRTY.* From SYSESTPRTY Where "
    strsql = strsql & " Dettipo = '" & sDettipo & "'"
    nTop = 17




  if sPart = "SI" then

     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if  not rs.eof then
         for n = 0 to nTop
             aFValueD(n) = rs(aFName(n))
         next
         aFValueD(18) = "SI"
     end if
     call ofv.cerrarconsulta(rs)


     strsql = "Select * From Objestprty Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
     strsql = strsql & sTipo & "' And Detcon = " & ID2 & " And Dettipo = '" & sDettipo & "'"
   '  response.write strsql
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if  not rs.eof then
         for n = 0 to nTop
             if not isnull(rs(aFName(n))) then aFValueD(n) = rs(aFName(n))
         next
     end if
     call ofv.cerrarconsulta(rs)



    strsql = "Select * From Objestdet Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
    strsql = strsql & sTipo & "' And Objid = " & ID
    nTop = 18
  else


     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if  not rs.eof then
         for n = 0 to nTop
             aFValueD(n) = rs(aFName(n))
         next
         aFValueD(18) = "SI"

     end if
     call ofv.cerrarconsulta(rs)



    strsql = "Select * From Objestprty Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
    strsql = strsql & sTipo & "' And Detcon = " & ID & " And Dettipo = '" & sDettipo & "'"
    nTop = 17
  end if


  end if


'response.write StrSql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      nID = rs(0)
      for n = 0 to nTop
          if not isnull(rs(aFName(n))) then aFValue(n) = rs(aFName(n))
      next
      if  aFValue(18) = "" Then aFValue(18) = "SI"

  else

      if stipo <> "SYSTEM" and sPart <> "SI" then

         strsql2 = "Insert Into Objestprty (Codobjeto,Tipobjeto,detcon,dettipo"
         strsql2 = strsql2 & ") Values (" & nCodigo & ",'" & sTipo & "'," & id & ",'"
         strsql2 = strsql2 & sDettipo & "')"
         call ofv.crearconsultaEx(StrSql2,cn,1,volver)

         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if  not rs.eof then
             nID = rs(0)
         end if

      end if



  end if
  call ofv.cerrarconsulta(rs)



  dim aSiNo(1)
  aSiNo(0) = "NO"
  aSiNo(1) = "SI"




'Default
  if sPart = "SI" then




    sDef = aFValue(18)

     SCOMSEP = ""
     scomK   = ""

     For n=0 to Ubound(aSiNo)
         scomK = scomK & SCOMSEP & aSiNo(n)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     aInput(18) = ofv.gencomboFX(sForm & "18",aFname(18),aFValue(18),scomK,scomD,"")


  else
    sDef = "NO"
  end if


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
  dim aBorder(8)
  aBorder(0) = "Ninguno"
  aBorder(1) = "Recuadro"
  aBorder(2) = "Elevado"
  aBorder(3) = "Bajorrelieve"
  aBorder(4) = "Marco 3D (1)"
  aBorder(5) = "Marco 3D (2)"
  aBorder(6) = "Punteado"
  aBorder(7) = "Guiones"
  aBorder(8) = "Doble"

  dim aBorder2(8)
  aBorder2(0) = ""
  aBorder2(1) = "Solid"
  aBorder2(2) = "Outset"
  aBorder2(3) = "Inset"
  aBorder2(4) = "Ridge"
  aBorder2(5) = "Groove"
  aBorder2(6) = "dotted"
  aBorder2(7) = "dashed"
  aBorder2(8) = "double"

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
  if sEst = "TEXTO" then

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


  elseif sEst = "TEXTO" then




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
      aInput(11) = aInput(11) & " style=width:31%;font-family:'"
      aInput(11) = aInput(11) & xvalor10 & "';"
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

  if  sPart  <> "SI" then

             sHTMLa = sHTMLa &  "<table  cellspacing=1 cellpadding=3  >"

             for ii=1 to 10
                 if  atit(ii) <> "NO" then

                    if  REPLACE(atit(ii)," ","") <> sEst then
                        xmDato  = REPLACE(atit(ii)," ","&nbsp;")
                        sHTMLa  = sHTMLa &  "<td class=bietx align=center style={border-style:solid;border-width:1;"
                        sHTMLa  = sHTMLa &  "border-color:#eeeeee} >"

                        href = "edestrprtymst.asp?codobjeto=" & nCodigo & "&tipobjeto=" & sTipo
                        href = href & "&Part=" & sPart & "&ID=" & ID
                        href = href & "&tipo=" & REPLACE(atit(ii)," ","")
                        href = href & "&volver=" & volver

                        sHTMLa  = sHTMLa &  "<a href='" & href & "' >" & xmDato & "</a></TD>"
                    end if

                 end if
            next

            sHTMLa = sHTMLa &  "</table>"

  end if



''''''''''''' Panel de Propiedades

  'sImagenDF = "<FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT>"
  sImagenDF = "<i class='las la-angle-right' STYLE=COLOR:" & Session("CMFPNTIT") & "></i>"

  sImagenNM = "<FONT FACE=corflaticon COLOR=#dddddd SIZE=2>4</FONT>"


  xTrColor = "#FFFFFF"

  bstitulo = "<font>Propiedades " & sInput & "&nbsp;" & sEsquema & ": " & scodiden & " - " & sesqnom & "</font>"

  Celda    = ofv.GenCelda("","ttt colspan=3 ","","","","","",bstitulo,"si")
  Fila     = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


  Celda    = ofv.GenCelda("","cc","","40%","","","","<big>" & sEst & "</big>","si")
  sInput   = ofv.GenerarInput("","Actualizar","submit","","10","bt","")

  sInput   = sHTMLa

  Celda    = Celda & ofv.GenCelda("","ut colspan=2 ","right","60%","","","",sInput,"si")
  Fila     = Fila & ofv.GenRow("","","","","","","",Celda,"si")


  ftab = ofv.GenTabla("","","","50% ","","0","","0","0","","",Fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","90%","","0","","0","0","","",filas,"si")

'  sHTML = sHTML & ofv.GenTabla("","","","60%","","0","","0","0","","",Fila,"si")
  Fila     = ""
    FilaS = ""

  sAccion = ""
 ' response.write sPart
  if bModificar then
     if stipo = "SYSTEM" then
          sAccion = "asp/ActualizarPrty.asp?ID=SYSTEM&detcon=NO&volver=" & volver
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst
     else
        if sPart = "SI" then
          sAccion = "asp/ActualizarPrty2.asp?ID=" & nID & "&volver=" & volver
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst & "&ID2=" & id2
        else
          sAccion = "asp/ActualizarPrty.asp?ID=" & nID & "&detcon=" & ID & "&volver=" & volver
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst  '& "&ID2=" & id2
        end if
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
              Celda = ofv.GenCelda("","uti","","","","","",aFtitle(n),"si")
              if  aInput(n) = "" then
                  sInput = ofv.GenerarInput(aFname(n),aFvalue(n),"text","","80","dt","")
              else
                  sInput = aInput(n)
              end if

              if  aInputD(n) = "N" then
                  sclasx  = " style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' "
                  sInputf = sImagenDF

              else
                  sclasx  = ""
                  sInputf = sImagenNM
              end if

              Celda = Celda & ofv.GenCelda("","ut","center","1%","","","",sInputf,"si")
              sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",sInput,"si")
              Celda = Celda & ofv.GenCelda("","uti " & sclasx & " ","","","","","",sInput,"si")

              'sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",Celda,"si")

              Fila  = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", celda, "si")

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

              if  aInputD(n) = "N" then
                  sclasx  = " style='border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";' "
                  sInputf = sImagenDF

              else
                  sclasx  = ""
                  sInputf = sImagenNM
              end if

              Celda = Celda & ofv.GenCelda("","ut","center","1%","","","",sInputf,"si")

            '  Celda = Celda & ofv.GenCelda("","uti " & sclasx & " ","","","","","",sInput,"si")

             ' sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",Celda,"si")
              sInput = ofv.GenForm(sform & n,"",sAccion,"","post","",sinput,"si")
              Celda = Celda & ofv.GenCelda("","uti " & sclasx & " ","","","","","",sInput,"si")

              Fila  = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", celda, "si")

              if  xTrColor = "#FFFFFF" then
                  xTrColor = "#F9F9F9"
              ELSE
                  xTrColor = "#FFFFFF"
              END IF


          end if

  end if

'  sAccion = ""
'  if bModificar then
'     if stipo = "SYSTEM" then
'          sAccion = "asp/ActualizarPrty.asp?ID=SYSTEM&detcon=NO&volver=" & volver
'          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst
'     else
'        if sPart = "SI" then
'          sAccion = "asp/ActualizarPrty2.asp?ID=" & nID & "&volver=" & volver
'          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst & "&ID2=" & id2
'        else
'          sAccion = "asp/ActualizarPrty.asp?ID=" & nID & "&detcon=" & ID & "&volver=" & volver
'          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst
'        end if
'     end if
'  end if

  Formulario = Fila


  ftab = ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;'","","90%","","0","","0","0","","",filas,"si")


 ' sHTML = sHTML & ofv.GenTabla("","","","60%","","0","","0","0","","",Formulario,"si")


  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  if  volver <> "" then
      sAccion = replace(replace(volver,"../",""),"_","&") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Celda = ofv.BotonVolver(sAccion)
  end if


  xxxxx = ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,Celda)

  if stipo = "SYSTEM" then

  else

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","10%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  end if

  sHTML = sHTML & xxxxx

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>