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
  sEst    = UCase(request.querystring("Tipo"))
  sPart   = request.querystring("Part")
  ID      = request.querystring("ID")

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
  end select

  dim aFName(18)
  dim aFValue(18)
  dim aFTitle(18)
  dim aInput(18)

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

  sHTML = ofv.Menuheader("","")

  set cn = ofv.conectar(ofv.strconn2)
  sEsquema = "S/D"
  scodiden = "S/D"
  sesqnom = "S/D"

  if stipo = "SYSTEM" then

    sEsquema = "DEFAULT"
    scodiden = ""
    sesqnom  = "SISTEMA"

    strsql = "Select dettipo as a0,dettipo as a0,dettipo as a0,dettipo as a0, * From SYSESTPRTY Where "
    strsql = strsql & " Dettipo = '" & sDettipo & "'"
    nTop = 17

  else

  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then   sEsquema = rs(1)

  Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     scodiden = rs(1)
     sesqnom = rs(3)
  end if

  if sPart = "SI" then
    strsql = "Select * From Objestdet Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
    strsql = strsql & sTipo & "' And Objid = " & ID
    nTop = 18
  else
    strsql = "Select * From Objestprty Where Codobjeto = " & nCodigo & " And Tipobjeto = '"
    strsql = strsql & sTipo & "' And Detcon = " & ID & " And Dettipo = '" & sDettipo & "'"
    nTop = 17
  end if


  end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    nID = rs(0)
    for n = 0 to nTop
      if not isnull(rs(aFName(n))) then aFValue(n) = rs(aFName(n))
    next
    if aFValue(18) = "" Then aFValue(18) = "SI"
  end if
  call ofv.cerrarconsulta(rs)

  aFtitle(0) = "Font Name"
  aFtitle(1) = "Font Size"
  aFtitle(2) = "Font Color"
  aFtitle(3) = "Background Color"
  aFtitle(4) = "Border"
  aFtitle(5) = "Tipo"
  aFtitle(6) = "Sangria"
  aFtitle(7) = "Indentacion"
  aFtitle(8) = "Espaciado despues del bloque"
  aFtitle(9) = "Viñeta"
  aFtitle(10) = "Font de Viñeta"
  aFtitle(11) = "Tipo de Viñeta"
  aFtitle(12) = "Espaciado de Viñeta"
  aFtitle(13) = "Numeracion"
  aFtitle(14) = "Tipo de Numeracion"
  aFtitle(15) = "Estilo de Numeracion"
  aFtitle(16) = "Alineacion"
  aFtitle(17) = "Alineacion del Texto"
  aFtitle(18) = "Default"

  dim aSiNo(1)
  aSiNo(0) = "NO"
  aSiNo(1) = "SI"

  dim aSpacing(10)
  for i = 0 to 10
    aSpacing(i) = CStr(i)
  next

  dim aVSpacing(4)
  for i = 0 to 4
    aVSpacing(i) = CStr(i)
  next

'Default
  if sPart = "SI" then
    aInput(18) = "<font class=at><select name=" & aFname(18)
    aInput(18) = aInput(18) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSiNo)
      sSelect = ""
      If aSiNo(n) = aFValue(18) Then sSelect = " selected"
      aInput(18) = aInput(18) & "<option value='" & aSiNo(n) & "' " & sSelect & ">"
      aInput(18) = aInput(18) & aSiNo(n) & "</option>"
    Next
    aInput(18) = aInput(18) & "</select></font>"
    sDef = aFValue(18)
  else
    sDef = "NO"
  end if

set ofr = server.createobject("corscr.Corrender")

ofr.verfonts

'wdimen = ofr.wscreen.fontcount
'wdimen = cint(wdimen)


'Font Name
  redim preserve aFontName(0)
  aFontName(0) = ""
'  aFontName(1)  = "Arial"
'  aFontName(2)  = "Courier"
'  aFontName(3)  = "Tahoma"
'  aFontName(4)  = "Times New Roman"
'  aFontName(5)  = "Verdana"
'  aFontName(6)  = "Comic Sans MS"
'  aFontName(7)  = "Helvetica"
'  aFontName(8)  = "Garamond"
'  aFontName(9)  = "Lucida"
'  aFontName(10) = "Impact"

for ii=0 to ofr.wscreen.fontcount -1
    redim preserve aFontName(ii + 1)
    aFontName(ii + 1) = ofr.wscreen.fonts(ii)
next
 
 aInput(0) = "<font class=at><select name=" & aFname(0)
  aInput(0) = aInput(0) & " onchange=document.Propiedades.submit();>"
  For n = 0 To Ubound(aFontName)
    sSelect = ""
    If aFontName(n) = aFValue(0) Then sSelect = " selected"
    aInput(0) = aInput(0) & "<option value='" & aFontName(n) & "' " & sSelect & ">"
    aInput(0) = aInput(0) & aFontName(n) & "</option>"
  Next
  aInput(0) = aInput(0) & "</select></font>"

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
  aInput(1) = "<font class=at><select name=" & aFname(1)
  aInput(1) = aInput(1) & " onchange=document.Propiedades.submit();>"
  For n = 0 To Ubound(aFontSize)
    sSelect = ""
    If aFontSize(n) = aFValue(1) Then sSelect = " selected"
    aInput(1) = aInput(1) & "<option value='" & aFontSize(n) & "' " & sSelect & ">"
    aInput(1) = aInput(1) & aFontSize(n) & "</option>"
  Next
  aInput(1) = aInput(1) & "</select></font>"

'Font Color
  strsql = "Select * From Syscolor"
  aInput(2) = ofv.GenerarCombo("Propiedades",aFName(2),aFValue(2),Strsql,cn,2,1,"sb")

'Background Color
  aInput(3) = ofv.GenerarCombo("Propiedades",aFName(3),aFValue(3),Strsql,cn,2,1,"sb")

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
  aInput(4) = "<font class=at><select name=" & aFname(4)
  aInput(4) = aInput(4) & " onchange=document.Propiedades.submit();>"
  For n = 0 To Ubound(aBorder)
    sSelect = ""
    If aBorder2(n) = aFValue(4) Then sSelect = " selected"
    aInput(4) = aInput(4) & "<option value='" & aBorder2(n) & "' " & sSelect & ">"
    aInput(4) = aInput(4) & aBorder(n) & "</option>"
  Next
  aInput(4) = aInput(4) & "</select></font>"

'Tipo
  sValue = aFValue(5)
  n = instr(sValue,"1")
  sChecked = ""
  if n > 0 then sChecked = " checked"
  aInput(5) = "<input type=checkbox value=1 name=" & aFName(5) & "a" & sChecked
  aInput(5) = aInput(5) & " onclick=document.Propiedades.submit();>Bold"
  sChecked = ""
  n = instr(sValue,"2")
  if n > 0 then sChecked = " checked"
  aInput(5) = aInput(5) & "<input type=checkbox value=2 name=" & aFName(5) & "b" & sChecked
  aInput(5) = aInput(5) & " onclick=document.Propiedades.submit();>Italic"
  sChecked = ""
  n = instr(sValue,"4")
  if n > 0 then sChecked = " checked"
  aInput(5) = aInput(5) & "<input type=checkbox value=4 name=" & aFName(5) & "c" & sChecked
  aInput(5) = aInput(5) & " onclick=document.Propiedades.submit();>Underline"
  sChecked = ""
  n = instr(sValue,"8")
  if n > 0 then sChecked = " checked"
  aInput(5) = aInput(5) & "<input type=checkbox value=8 name=" & aFName(5) & "d" & sChecked
  aInput(5) = aInput(5) & " onclick=document.Propiedades.submit();>Strikeout"

'Sangria
  if sEst = "TEXTO" then
    aInput(6) = "<font class=at><select name=" & aFname(6)
    aInput(6) = aInput(6) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSpacing)
      sSelect = ""
      If aSpacing(n) = aFValue(6) Then sSelect = " selected"
      aInput(6) = aInput(6) & "<option value='" & aSpacing(n) & "' " & sSelect & ">"
      aInput(6) = aInput(6) & aSpacing(n) & "</option>"
    Next
    aInput(6) = aInput(6) & "</select></font>"
  else
    aInput(6) = "NO"
  end if

'Indent 1
  if sEst = "TITULO"  And sPart <> "SI" then
    aInput(7) = "<font class=at><select name=" & aFname(7)
    aInput(7) = aInput(7) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSiNO)
      sSelect = ""
      If aSiNO(n) = aFValue(7) Then sSelect = " selected"
      aInput(7) = aInput(7) & "<option value='" & aSiNO(n) & "' " & sSelect & ">"
      aInput(7) = aInput(7) & aSiNO(n) & "</option>"
    Next
    aInput(7) = aInput(7) & "</select></font>"
  else
    aInput(7) = "NO"
  end if

'Espaciado vertical
  if sPart <> "SI" then
    aInput(8) = "<font class=at><select name=" & aFname(8)
    aInput(8) = aInput(8) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aVSpacing)
      sSelect = ""
      If aVSpacing(n) = aFValue(8) Then sSelect = " selected"
      aInput(8) = aInput(8) & "<option value='" & aVSpacing(n) & "' " & sSelect & ">"
      aInput(8) = aInput(8) & aVSpacing(n) & "</option>"
    Next
    aInput(8) = aInput(8) & "</select></font>"
  else
    aInput(8) = "NO"
  end if

'Viñeta
  if sEst <> "TEXTO" then
    aInput(9) = "<font class=at><select name=" & aFname(9)
    aInput(9) = aInput(9) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSiNo)
      sSelect = ""
      If aSiNo(n) = aFValue(9) Then sSelect = " selected"
      aInput(9) = aInput(9) & "<option value='" & aSiNo(n) & "' " & sSelect & ">"
      aInput(9) = aInput(9) & aSiNo(n) & "</option>"
    Next
    aInput(9) = aInput(9) & "</select></font>"
  elseif sEst = "TEXTO" then
   aInput(9) = "<font class=at><select name=" & aFname(9)
    aInput(9) = aInput(9) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSiNo)
      sSelect = ""
      If aSiNo(n) = aFValue(9) Then
         sSelect = " selected"
      elseif  aSiNo(n) = "SI" and aFValue(9) = "" Then
         sSelect = " selected"
      end if
      aInput(9) = aInput(9) & "<option value='" & aSiNo(n) & "' " & sSelect & ">"
      aInput(9) = aInput(9) & aSiNo(n) & "</option>"
    Next
    aInput(9) = aInput(9) & "</select></font>"
  else
    aInput(9) = "NO"
    aFValue(9) = "NO"
  end if

  if sEst <> "TEXTO" and aFValue(9) = "SI" then
    'Font de Viñeta
    dim aFontName2(3)
    aFontName2(0) = ""
    aFontName2(1) = "corflaticon"
    aFontName2(2) = "Symbol"
    aFontName2(3) = "corflaticon"
    aInput(10) = "<font class=at><select name=" & aFname(10)
    aInput(10) = aInput(10) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aFontName2)
      sSelect = ""
      If aFontName2(n) = aFValue(10) Then sSelect = " selected"
      aInput(10) = aInput(10) & "<option value='" & aFontName2(n) & "' " & sSelect & ">"
      aInput(10) = aInput(10) & aFontName2(n) & "</option>"
    Next
    aInput(10) = aInput(10) & "</select></font>"
    if aFValue(10) <> "" then
      'Tipo de Viñeta
      dim aTipoVin(222)
      j = 0
      for i = 33 to 255
        aTipoVin(j) = Chr(i)
        j = j + 1
      next
      aInput(11) = "<font face='" & aFValue(10) & "' size=2><select "
      aInput(11) = aInput(11) & " style={font-family:'"
      aInput(11) = aInput(11) & aFValue(10) & "';font-size:14pt;}"
      aInput(11) = aInput(11) & " name=" & aFname(11)
      aInput(11) = aInput(11) & " onchange=document.Propiedades.submit();>"
      For n = 0 To Ubound(aTipoVin)
        sSelect = ""
        If aTipoVin(n) = aFValue(11) Then sSelect = " selected"
        aInput(11) = aInput(11) & "<option value='" & aTipoVin(n) & "' " & sSelect & ">"
        aInput(11) = aInput(11) & aTipoVin(n) & "</option>"
      Next
      aInput(11) = aInput(11) & "</select></font>"
    else
      aInput(11) = "NO"
    end if
    'Espaciado de Viñeta
    aInput(12) = "<font class=at><select name=" & aFname(12)
    aInput(12) = aInput(12) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSpacing)
      sSelect = ""
      If aSpacing(n) = aFValue(12) Then sSelect = " selected"
      aInput(12) = aInput(12) & "<option value='" & aSpacing(n) & "' " & sSelect & ">"
      aInput(12) = aInput(12) & aSpacing(n) & "</option>"
    Next
    aInput(12) = aInput(12) & "</select></font>"
  else
    aInput(10) = "NO"
    aInput(11) = "NO"
    aInput(12) = "NO"
  end if

'Numeracion
  if sEst = "TITULO"  And sPart <> "SI" then
    aInput(13) = "<font class=at><select name=" & aFname(13)
    aInput(13) = aInput(13) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aSiNo)
      sSelect = ""
      If aSiNo(n) = aFValue(13) Then sSelect = " selected"
      aInput(13) = aInput(13) & "<option value='" & aSiNo(n) & "' " & sSelect & ">"
      aInput(13) = aInput(13) & aSiNo(n) & "</option>"
    Next
    aInput(13) = aInput(13) & "</select></font>"
  else
    aInput(13) = "NO"
 '   if sEst <> "TEXTO" then
      aFValue(13) = "SI"
 '   else
  '    aFValue(13) = "NO"
'    end if
  end if

  if aFValue(13) = "SI" then
    'Tipo de Numeracion
    dim aTipoNum(5)
    aTipoNum(0) = "1"
    aTipoNum(1) = "I"
    aTipoNum(2) = "i"
    aTipoNum(3) = "A"
    aTipoNum(4) = "a"
    aTipoNum(5) = "NO"
    aInput(14) = "<font class=at><select name=" & aFname(14)
    aInput(14) = aInput(14) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aTipoNum)
      sSelect = ""
      If aTipoNum(n) = aFValue(14) Then sSelect = " selected"
      aInput(14) = aInput(14) & "<option value='" & aTipoNum(n) & "' " & sSelect & ">"
      aInput(14) = aInput(14) & aTipoNum(n) & "</option>"
    Next
    aInput(14) = aInput(14) & "</select></font>"
 '   if sEst = "TITULO" then
     if aFValue(14) <> "NO" then
      'Clase de Numeracion
      dim aClaseNum(1)
      aClaseNum(0) = "Simple"
      aClaseNum(1) = "Compuesta"
      aInput(15) = "<font class=at><select name=" & aFname(15)
      aInput(15) = aInput(15) & " onchange=document.Propiedades.submit();>"
      For n = 0 To Ubound(aClaseNum)
        sSelect = ""
        If aClaseNum(n) = aFValue(15) Then sSelect = " selected"
        aInput(15) = aInput(15) & "<option value='" & aClaseNum(n) & "' " & sSelect & ">"
        aInput(15) = aInput(15) & aClaseNum(n) & "</option>"
      Next
      aInput(15) = aInput(15) & "</select></font>"
    else
      aInput(15) = "NO"
    end if
  else
    aInput(14) = "NO"
    aInput(15) = "NO"
  end if

'Alineacion
  dim aAlign(3)
  aAlign(0) = ""
  aAlign(1) = "Izquierda"
  aAlign(2) = "Centro"
  aAlign(3) = "Derecha"
  dim aAlign2(3)
  aAlign2(0) = ""
  aAlign2(1) = "left"
  aAlign2(2) = "center"
  aAlign2(3) = "right"
  if sEst <> "TEXTO" then
    aInput(16) = "<font class=at><select name=" & aFname(16)
    aInput(16) = aInput(16) & " onchange=document.Propiedades.submit();>"
    For n = 0 To Ubound(aAlign)
      sSelect = ""
      If aAlign2(n) = aFValue(16) Then sSelect = " selected"
      aInput(16) = aInput(16) & "<option value='" & aAlign2(n) & "' " & sSelect & ">" & aAlign(n)
      aInput(16) = aInput(16) & "</option>"
    Next
    aInput(16) = aInput(16) & "</select></font>"
  else
    aInput(16) = "NO"
  end if

'Alineacion Texto
  aInput(17) = "<font class=at><select name=" & aFname(17)
  aInput(17) = aInput(17) & " onchange=document.Propiedades.submit();>"
  For n = 0 To Ubound(aAlign)
    sSelect = ""
    If aAlign2(n) = aFValue(17) Then sSelect = " selected"
    aInput(17) = aInput(17) & "<option value='" & aAlign2(n) & "' " & sSelect & ">" & aAlign(n)
    aInput(17) = aInput(17) & "</option>"
  Next
  aInput(17) = aInput(17) & "</select></font>"
  if sPart = "SI" then
     sInput = "Particulares"
  else
     sInput = "Generales"
  end if
  Celda = ofv.GenCelda("","tt","","1%","","","","Propiedades " & sInput,"si")
  bstitulo = " " & sEsquema & ": <b>" & scodiden & " - " & sesqnom & "</b>"
  Celda = Celda & ofv.GenCelda("","cc","","1%","","","",bstitulo,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  sInput = ofv.GenerarInput("","Actualizar","submit","","10","bt","")
  Celda = ofv.GenCelda("","tt","","1%","","","","Propiedades de " & sEst,"si")
  Celda = Celda & ofv.GenCelda("","cc","right","1%","","","",sInput,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  if sPart = "SI" then
    Celda = ofv.GenCelda("","cc","","1%","","","",aFtitle(18),"si")
    sInput = aInput(18)
    Celda = Celda & ofv.GenCelda("","cc","","1%","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  end if
  if sDef = "NO" then
    for n = 0 to UBound(aFName) - 1
      if aInput(n) <> "NO" then
        Celda = ofv.GenCelda("","cc","","1%","","","",aFtitle(n),"si")
        if aInput(n) = "" then
          sInput = ofv.GenerarInput(aFname(n),aFvalue(n),"text","","80","dt","")
        else
          sInput = aInput(n)
        end if
        Celda = Celda & ofv.GenCelda("","cc","","1%","","","",sInput,"si")
        Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      end if
    next
  end if

  sAccion = ""
  if bModificar then
     if stipo = "SYSTEM" then
          sAccion = "asp/ActualizarPrty.asp?ID=SYSTEM&detcon=NO&volver=" & volver
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst
     else 
        if sPart = "SI" then
          sAccion = "asp/ActualizarPrty2.asp?ID=" & nID & "&volver=" & volver
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst
        else
          sAccion = "asp/ActualizarPrty.asp?ID=" & nID & "&detcon=" & ID & "&volver=" & volver
          sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Tipo=" & sEst
        end if
     end if
  end if
  Formulario = ofv.GenForm("Propiedades","",sAccion,"","post","",Fila,"si")

  sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","#993333","",Formulario,"si")
  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>