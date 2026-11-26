<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else


    

    
    nCodigo = Request.querystring("Codproced")
    sTipo = Request.querystring("Tipobjeto")
    vuelta = Request.querystring("Volver")

    Set cn = OFV.conectar(OFV.strconn2)

    bstilo = ""
    borderl = ""
    borderr = ""
    bordert = ""
    borderb = ""
    Cstilo = "border-style:outset;border-width:2;border-color:#dddddd;"
    Cbordert = "border-top-style:ridge;border-top-width:2;border-top-color:#dddddd;"
    Cborderl = "border-left-style:ridge;border-left-width:2;border-left-color:#dddddd;"
    Cborderr = "border-right-style:ridge;border-right-width:2;border-right-color:#dddddd;"
    Cborderb = "border-bottom-style:ridge;border-bottom-width:2;border-bottom-color:#dddddd;"

    Cborderls = "border-left-style:solid;border-left-width:1;border-left-color:#336699;"

    bstilo = Cstilo
    borderl = Cborderl
    borderr = Cborderr
    bordert = Cbordert
    borderb = Cborderb

    StrSql = "Select codobjeto,tipobjeto From OBJPROCED Where codproced = '" & nCodigo & "'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then
       sOCesque = rs(0)
       sOTesque = rs(1)
    End If
    Call OFV.cerrarconsulta(rs)

  strsqlx = "Select Version From Objprty Where codobjeto = " & sOCesque
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

    StrSql = "Select descri From tipobjeto Where tipobjeto = '" & sOTesque & "'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then sTesque = rs(0)
    Call OFV.cerrarconsulta(rs)

    StrSql = "Select codiden,descri From objetos Where codobjeto = " & sOCesque
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then
       sCesque = rs(0)
       sDesque = rs(1)
    End If
    Call OFV.cerrarconsulta(rs)

    StrSql = "Select Archivocur From Modo Where tipomodo = 'I'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then sInicio = rs(0)
    Call OFV.cerrarconsulta(rs)

    StrSql = "Select Archivocur From Modo Where tipomodo = 'F'"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    If Not rs.EOF Then sFin = rs(0)
    Call OFV.cerrarconsulta(rs)


    '' response.write "aa"

      sCaption = sTesque & "&nbsp;" & sCesque & " - " & sDesque
'    sHTML = OFV.Repbody("CURSOGRAMA:&nbsp;" & sCaption) & OFV.RepHeader2("CURSOGRAMA:&nbsp;" & sCaption, "100% ")

  sHTML = ofv.FormBody
    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=300,width=500,scrollbars=yes,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)
    xs = xs & "function pointcon(xcon) {" & Chr(13)
'    xs = xs &  " alert(xcon + ': ' + document.all(xcon).offsetParent.offsetParent.offsetParent.tagName  + document.all(xcon).offsetParent.offsetParent.offsetParent.offsetTop + ' ' + "
'    xs = xs &  " document.all(xcon).offsetParent.offsetParent.tagName  + document.all(xcon).offsetParent.offsetParent.offsetTop + ' ' + "
'    xs = xs &  " document.all(xcon).offsetParent.offsetLeft + ' ' + document.all(xcon).offsetParent.offsetTop);"
'    xs = xs &  " alert(xcon + ': ' + document.all(xcon).offsetParent.offsetParent.offsetLeft + ' ' + document.all(xcon).offsetParent.offsetParent.offsetTop);"
'    xs = xs &  " alert(document.all(xcon).scrollLeft + ' ' + document.all(xcon).scrollTop);"


'    xs = xs &  " window.scrollTo(document.all(xcon)"
'    xs = xs &  ".offsetParent.offsetLeft,document.all(xcon)"
'    xs = xs &  ".offsetParent.offsetTop);"

    xs = xs &  " window.scrollTo(document.all(xcon)"
    xs = xs &  ".offsetParent.offsetLeft,document.all(xcon)"
    xs = xs &  ".offsetParent.offsetParent.offsetParent.offsetTop + document.all(xcon).offsetParent.offsetTop);"

 '   xs = xs &  " alert(xcon);"
 '   xs = xs &  " document.all(xcon).scrollIntoView(true);"



    xs = xs &  Chr(13) & " } " & Chr(13)
    xs = xs &  "</script>" & Chr(13)

    sHTML = sHTML & xs



   redim asecto(0)

   asecto(0) = 0  
   isec = 0
   isecb = 0

'''' loop sectores

    StrSql = "SELECT * from cursmaster "
    StrSql = StrSql & " Where codobjeto = " & sOCesque 
    StrSql = StrSql & " order by secuen, secuencia " 
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    do until rs.eof
       if isec = 0 then
          isec = isec + 1
          asecto(0) = asecto(0) + 1
          redim preserve asecto(isec) 
          asecto(isec) = rs(1)
       else
          for isecb=1 to isec
              if asecto(isecb) = rs(1) then 
                 isecb = 9999
                 exit for
              end if
          next
          if isecb <> 9999 then
             isec = isec + 1
             asecto(0) = asecto(0) + 1
             redim preserve asecto(isec) 
             asecto(isec) = rs(1)    
          end if     
       end if   
    rs.movenext
    loop
    Call OFV.cerrarconsulta(rs)



   redim alnkv(0)
   redim alnkl(0)
   redim alnkvD(0)
   redim alnklD(0)


   ilnkv = 0
   ilnkvb = 0
   ilnkl = 0
   ilnklb = 0

    StrSql = "SELECT * from cursmasterlnk "
    StrSql = StrSql & " Where codobjeto = " & sOCesque & " and ltipo = 'R' "
    StrSql = StrSql & " order by sece, secp, seca" 
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    do until rs.eof
          ilnkv = ilnkv + 1
          redim preserve alnkv(ilnkv) 
          redim preserve alnkvD(ilnkv) 
          alnkv(ilnkv) = rs(1)
          daccD = "&nbsp;"
          daccP = " "
          daccE = "&nbsp;"

          StrSql = "Select codproced,secuen,descri From Cursograma Where codaccion = '" & rs(2) & "'"
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          if not rs2.eof then
             daccD = rs2(1) & "." & rs2(2)
             daccP = rs2(0)
          end if
          call ofv.cerrarconsulta(rs2)
 

          Strsql = "SELECT  nombre "
          Strsql = Strsql & " from CSCEventoSectorPerfil Where codproced = '" & daccP & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rs2.eof then
             daccE = rs2(0)
             daccE = replace(daccE," ","&nbsp;")
          end if
          call ofv.cerrarconsulta(rs2)

          alnkvD(ilnkv) = "&nbsp;A:&nbsp;" & daccE & chr(13) & daccD 

          ilnkl = ilnkl + 1
          redim preserve alnkl(ilnkl) 
          redim preserve alnklD(ilnkl) 
          alnkl(ilnkl) = rs(2)

          daccD = "&nbsp;"
          daccP = " "
          daccE = "&nbsp;"

          StrSql = "Select codproced,secuen,descri From Cursograma Where codaccion = '" & rs(1) & "'"
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          if not rs2.eof then
             daccD = rs2(1) & "." & rs2(2)
             daccP = rs2(0)
          end if
          call ofv.cerrarconsulta(rs2)
 

          Strsql = "SELECT  nombre "
          Strsql = Strsql & " from CSCEventoSectorPerfil Where codproced = '" & daccP & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rs2.eof then
             daccE = rs2(0)
             daccE = replace(daccE," ","&nbsp;")
          end if
          call ofv.cerrarconsulta(rs2)

          alnklD(ilnkl) = "&nbsp;DE:&nbsp;" & daccE & chr(13) & daccD

    rs.movenext
    loop
    Call OFV.cerrarconsulta(rs)



'''' loop proced
     Tablon = ""
     filon = "" 
     Celda = ""
     tsize = 100

     if isec > 2 then tsize = 40 * isec

  sCaption = "<font color=#336699 face=verdana size=1><b>CURSOGRAMA:&nbsp;" & sCaption & "</b></font>"

  sHTML = sHTML & "<table width=" & tsize & "% cellspacing=1 cellpadding=3 border ><tr><td width=100% "
  sHTML = sHTML & "  style={border-style:outset;border-width:2;} >" & sCaption & "</td></tr></table>"
     
            for isecb = 1 to isec 
                sInput = "<p align=center><font color=#ffffff face=verdana size=1><b>" & asecto(isecb)
                sInput = sInput & "</b></font></p>"
                Celda = Celda & OFV.GenCelda("", "aaa style='" & borderr & borderl & "'", "", "", "", "#6699cc", "", sInput, "si")

            next 
    
            Filon = Filon & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


 aeven = "xxx"

'    StrSql = "SELECT * from cursmaster "
'    StrSql = StrSql & " Where codproced = '" & nCodigo & "'"

    StrSql = "SELECT * from cursmaster "
    StrSql = StrSql & " Where codobjeto = " & sOCesque 
    StrSql = StrSql & " order by secuen, secuencia " 

    Set rsx = OFV.crearconsultaEx(StrSql, cn, 1, "")
    do until rsx.eof 
    sEvento = "-"
    sSector = "-"
    sPerfil = "-"
    sSecu = "-"
    sESecu = "-"
    If Not IsNull(rsx(0)) Then sEvento = rsx(0)
    If Not IsNull(rsx(1)) Then sSector = rsx(1)
    If Not IsNull(rsx(2)) Then sPerfil = rsx(2)
    If Not IsNull(rsx(3)) Then sSecu = rsx(3)
    If Not IsNull(rsx(4)) Then sESecu = rsx(4)
 '   Call OFV.cerrarconsulta(rs)


    if aeven <> sEvento then
       aeven = sEvento
       sInput = "<p align=center><font color=#000080 face=verdana size=1><b>EVENTO:&nbsp;" & sESecu & "&nbsp;" & sEvento
       sInput = sInput & "</b></font></p>"
       if isec = 1 then
          Celda = OFV.GenCelda("", "cc", "", "", "", "#336699", "", sInput, "si")
       else
          Celda = OFV.GenCelda("", "cc colspan=" & isec & " ", "", "", "", "#336699", "", sInput, "si")
       end if

       Filon = Filon & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
    end if




' ''''''''''''''

    StrSql = "Select * From Cursograma Where codproced = '" & rsx(5) & "'"
    StrSql = StrSql & " and (modi = 0 or (modi = -1 and tipmodi <> 'E')) "
    StrSql = StrSql & " order by secuen, codaccion"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")




      'Encabezado

 '     sInput = "<p align=center><font color=#000077 face=verdana size=1><b>" & sCaption
 '     sInput = sInput & "</b></font></p>"
 '     Celda = OFV.GenCelda("", "", "", "", "15", "#eeddcc", "", sInput, "si")

      sCaption = ""
  '    sCaption = " EVENTO:&nbsp;" & sESecu & "." & sSecu & " - " & sEvento & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  '    sCaption = sCaption & "<BR>"
      sCaption = sCaption & "PERFIL/FUNCION:&nbsp;&nbsp;" & sPerfil
      sInput = "<p align=center><font color=#000080 face=verdana size=1><b><u>" & sCaption
      sInput = sInput & "</u></b></font></p>"
      Celda = OFV.GenCelda("", "", "", "", "15", "#336699", "", sInput, "si")
      Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
  '    Tabla = OFV.GenTabla("", "", "", "100% border", "", "", "", "0", "0", "", "", Fila, "si")
      Tabla = ""
    If Not rs.EOF Then

      'Titulos


      Fila = "" 
      Celda = ""
      Celda = Celda & OFV.GenCelda("", "", "", "10%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "3%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "3%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "6%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "3%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "3%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "6%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "3%", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "63%", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'perfil
      Celda = OFV.GenCelda("", "aaa colspan=9", "", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


      'Icono de Inicio
      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4 colspan=6", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      sImagen = OFV.GenImage("", "", "", "", "Iconos/" & sInicio, "0", "")
      Celda = Celda & OFV.GenCelda("", "abc rowspan=2  colspan=2 style='" & bstilo & "'", "center", "", "", "", "", sImagen, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f3
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f4
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



      nCont = 0
      If Not rs.EOF Then
        StrSql = "Select Continuador From AccionesAlt "
        StrSql = StrSql & "Where codproced = '" & rs(2) & "' and Continuador <> '0'"
        Set rsAlt = OFV.crearconsultaEx(StrSql, cn, 1, "")
        z = 1
        ReDim Continuadores(z)
        Do Until rsAlt.EOF
          ReDim Preserve Continuadores(z)
          Continuadores(z) = rsAlt(0)
          z = z + 1
          rsAlt.movenext
        Loop
        Call OFV.cerrarconsulta(rsAlt)

      End If

      Do Until rs.EOF

     slnkv = ""
     slnkl = ""
     ilnkvt = 0
     ilnklt = 0




        SArma2 = "&nbsp;"
        sComponente = "&nbsp;"
        stcomp = ""
        sacomp = "ICorAbierto.jpg"
        if rs(18) <> "" and rs(18) > "0"  then
          'Imagen de Complemento

          Strsql = "Select * From tipocomp Where tipo = '" & rs(19) & "' "
          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsComponentes.eof then 
             stcomp = rsComponentes(1) & ":&nbsp;" 
             sacomp = rsComponentes(2) 
          end if
          call ofv.cerrarconsulta(rsComponentes)

          sImage = ofv.GenImage("","","","","Iconos/" & sacomp,"0","")
          sHref = "EditarImagen2.asp?ID=" & rs(18)
          sLink = ofv.GenLink("","",sHref,"_blank","","",sImage,"si")

          Strsql = "Select descri From Complementos Where codcomp = " & rs(18)
          set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsComponentes.eof then sComponente = rsComponentes(0) 
          call ofv.cerrarconsulta(rsComponentes)

          SArma2 = "<font color=#000080 face=arial-narrow size=1><b>" & stcomp & sComponente & "</b></font>"
          SArma2 = sLink & "&nbsp;&nbsp;" & SArma2
        end if





        'Icono de la Accion


        'Secuencia
 ''''       sinputs = "<font color=#336699 face=verdana size=1><b>" & rs(3) & "&nbsp;" & slnkv & slnkl 
        sinputs = "<font color=#000000 face=verdana size=1><b>" & rs(3) & "&nbsp;" & slnkv & slnkl 
        sinputs = sinputs & "</b></font>"

 
       'Referencia
       sVerbo = rs(4)
        sAcciont = rs(4)
        i = InStr(sVerbo, " ")
        If i > 0 Then
          sVerbo = Left(rs(4), i - 1)
          sAcciont = Mid(rs(4), i + 1)
        End If

 ''''       sInputr = "<font color=#336699 face=verdana size=1><b>" & rs(3) & ": "

        sInputr = "<font color=#000000 face=verdana size=1><b>" & rs(3) & ": "
        sInputr = sInputr & "<i><b>" & sVerbo & "</b></i>" & "&nbsp;&nbsp;&nbsp;" & sAcciont & "</b></font>"


        If rs(16) > "0" Then
        ' Sector Continuador
        StrSql = "select * from eco where eco.eco = " & rs(16)
        Set rsc = OFV.crearconsultaEx(StrSql, cn, 1, "")
        If Not rsc.EOF Then sconti = rsc(1)
        End If
        If rs(17) > "0" Then
        ' Perfil Continuador
        StrSql = "select * from tipoperfil where tipoperfil = '" & rs(17) & "'"
        Set rsd = OFV.crearconsultaEx(StrSql, cn, 1, "")
        If Not rsd.EOF Then sperfi = rsd(1)
        End If
        If rs(16) > "0" Or rs(17) > "0" Then
           sArma = "<font color=#000080 face=arial-narrow size=1><b>Continuador:"
           sArma = sArma & sconti & " - " & sperfi
        Else
           sArma = "&nbsp;"
        End If



' '''''''''''''

     for ilnklb=1 to ilnkl
         if alnkl(ilnklb) = rs(12) then
'            ilnklt = 99999
'            if ilnklt = 99999 then
               slnkl = "<font color=#ffffff face=arial-narrow size=1><b>" & ilnklb & "</b></font>"
               pmkl = "<form name="  & chr(34) & "D" & ilnklb & "L"  & chr(34) &  " method=post >"
               pmkl = pmkl & "<input type=hidden value=a ></form>"
'            end if






  '   if ilnklt = 99999 then 
 '       slnkl = "<" & ilnklb 

          dtitu = " title='" & alnklD(ilnklb) & "' "

          sImagen = OFV.GenImage("", "", "", "", "Iconos/icorcconecva.jpg", "0", "")
      slnklt = "<A " & dtitu & " onclick=" & chr(34) & "pointcon('D" & ilnklb & "V');" & chr(34) & "  STYLE=" & chr(34) & "cursor:pointer;" & chr(34) & ">" & sImagen & "</A>"
      slnklt2 = "<A " & dtitu & " onclick=" & chr(34) & "pointcon('D" & ilnklb & "V');" & chr(34) & "  STYLE=" & chr(34) & "cursor:pointer;position:relative;top:-12px;left:-20px" & chr(34) & ">" & slnkl & "</A>"
      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "abc rowspan=2  colspan=2  ", "center valign=middle ", "", "", "", "", slnklt & slnklt2, "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   ", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;" & pmkl, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   style='" & cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")





'     end if

'            exit for
         end if  
     next

' '''''''''''''''


      sImage = "&nbsp;"
      if bVersion = "SI" and rs(14) then
      if rs(15) = "M" then
        sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
      elseif rs(15) = "A" then
        sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
      end if
      end if



              sActcon = "Coprocaccion.asp?codobjeto=" & rs(0) & "&Tipobjeto=" & rs(1)
              sActcon = sActcon & "&bversion=" & bVersion & "&baccion=" & rs(12) & "&nID=" & rs(2)
              sHref = "&nbsp;"

              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & rs(1) & "'); return false;" & Chr(34)


           sImagen1 = sInputr
           sImagen2 = OFV.GenImage("", "", "", "", "Iconos/" & rs(5), "0", "")

           sLink1 = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen1, "si")
           sLink2 = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen2, "si")



      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", sImage, "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4 colspan=4", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", sLink1 & "<br><br>" & sarma2, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      sImagen = OFV.GenImage("", "", "", "", "Iconos/" & rs(5), "0", "")
      Celda = Celda & OFV.GenCelda("", "abc rowspan=2  colspan=2 style='" & bstilo & "'", "center", "", "", "", "", sLink2, "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;" & sinputs, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f3
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", sArma, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f4
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


' '''''''''''''

     for ilnkvb=1 to ilnkv
         if alnkv(ilnkvb) = rs(12) then
 '           ilnkvt = 99999


  '   if ilnkvt = 99999 then 
        slnkv = "<font color=#ffffff face=arial-narrow size=1><b>" & ilnkvb & "</b></font>"
        pmkv = "<form name="  & chr(34) & "D" & ilnkvb & "V"  & chr(34) &  " method=post >"
        pmkv = pmkv & "<input type=hidden value=a ></form>"
 '    end if

 '    if ilnkvt = 99999 then
  '       slnkv = ">" & ilnkvb 

          sImagen = OFV.GenImage("", "", "", "", "Iconos/icorcconecllega.jpg", "0", "")

          dtitu = " title='" & alnkvD(ilnkvb) & "' "

      slnkvt = "<A " & dtitu & " onclick=" & chr(34) & "pointcon('D" & ilnkvb & "L');" & chr(34) & "  STYLE=" & chr(34) & "cursor:pointer;" & chr(34) & ">" & sImagen & "</A>"
      slnkvt2 = "<A " & dtitu & " onclick=" & chr(34) & "pointcon('D" & ilnkvb & "L');" & chr(34) & "  STYLE=" & chr(34) & "cursor:pointer;position:relative;top:-12px;left:-22px" & chr(34) & ">" & slnkv & "</A>"
      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "abc rowspan=2  colspan=2 ", "center valign=middle ", "", "", "", "", slnkvt & slnkvt2, "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   ", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;" & pmkv, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   style='" & cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")






 '     end if

 '           exit for
         end if  
     next


' '''''''''''''''



        StrSql = "Select * From AccionesAlt Where codaccion = '" & rs(12) & "'"
        StrSql = StrSql & " and (modi = 0 or (modi = -1 and tipmodi <> 'E')) "
           StrSql = StrSql & " Order By secuencia "

        Set rsAlt = OFV.crearconsultaEx(StrSql, cn, 1, "")

           nAlt = ""
           bordert = Cbordert
           borderr = Cborderr


        Do Until rsAlt.EOF

          'Icono de la Accion Alt
          sImagen = OFV.GenImage("", "", "", "", "Iconos/" & rsAlt(7), "0", "")


          'Secuencia Alt
          sinputs = "<font color=#993333 face=verdana size=1><b>" & rs(3) & "."
          sinputs = sinputs & rsAlt(1) & "</b></font>"


          'Referencia Alt
          sVerbo = rsAlt(0)
          sAcciont = rsAlt(0)
          i = InStr(sVerbo, " ")
          If i > 0 Then
             sVerbo = Left(rsAlt(0), i - 1)
             sAcciont = Mid(rsAlt(0), i + 1)
          End If
          sInputr = "<font color=#993333 face=verdana size=1><b>" & rs(3) & "."
          sInputr = sInputr & rsAlt(1) & ": "
          sInputr = sInputr & "<i><b>" & sVerbo & "</b></i>" & "&nbsp;&nbsp;&nbsp;" & sAcciont & "</b></font>"


          If nAlt = "" Or nAlt = "0" Then
                If Not IsNull(rsAlt(5)) Then nAlt = rsAlt(5)
          End If

          naccion = rsAlt(10) 

      sImage = "&nbsp;"
      if bVersion = "SI" and rsAlt(8) then
      if rsAlt(9) = "M" then
        sImage = ofv.GenImage("","","","","Imagenes/Botmodx.jpg","0","")
      elseif rsAlt(9) = "A" then
        sImage = ofv.GenImage("","","","","Imagenes/Botagrx.jpg","0","")
      end if
      end if


          rsAlt.movenext

          'Lineas
          If Not rsAlt.EOF Then

            If nAlt = "" Or nAlt = "0" Then
                  If Not IsNull(rsAlt(5)) Then nAlt = rsAlt(5)
            End If
            sArma = "&nbsp;"
            nSec = "NO"
          Else
            StrSql = "Select secuen From Cursograma Where codaccion = '" & nAlt & "'"
            StrSql = StrSql & " And codproced = '" & nCodigo & "'"

            Set rsSec = OFV.crearconsultaEx(StrSql, cn, 1, "")
            nSec = "Fin"
            If Not rsSec.EOF Then
              If Not IsNull(rsSec(0)) Then nSec = rsSec(0)
            End If
            Call OFV.cerrarconsulta(rsSec)
 
            'Conector
            sArma = "<font color=#000077 face=verdana size=1><b>" & nSec
            sArma = sArma & "</b></font>"

          End If



              sActcon = "Coprocaccionalt.asp?codobjeto=" & rs(0) & "&Tipobjeto=" & rs(1) & "&bsec=" & rs(3)
              sActcon = sActcon & "&bversion=" & bVersion & "&naccion=" & rs(12) & "&nID=" & rs(2) & "&baccion=" & naccion
              sHref = "&nbsp;"

              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & rs(1) & "'); return false;" & Chr(34)


           sImagen1 = sInputr
           sImagen2 = sImagen

           sLink1 = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen1, "si")
           sLink2 = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen2, "si")




      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", sImage, "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & bordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4  style='" & bordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderr & bordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4  ", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", sLink1, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "abc rowspan=2  colspan=2 style='" & bstilo & "'", "center", "", "", "", "", sLink2, "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;" & sinputs, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f3
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;" & sArma, "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f4
      If nSec <> "NO" Then borderr = ""

      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderr & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

        bordert = ""


        Loop
        Call OFV.cerrarconsulta(rsAlt)


 
        rs.movenext
        Loop




      'Icono de fin
      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4 colspan=6", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      sImagen = OFV.GenImage("", "", "", "", "Iconos/" & sFin, "0", "")
      Celda = Celda & OFV.GenCelda("", "abc rowspan=2  colspan=2 style='" & bstilo & "'", "center", "", "", "", "", sImagen, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f3
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f4
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa colspan=2", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

 


      Tabla = Tabla & "<center>"
      Tabla = Tabla & OFV.GenTabla("", "aaa ", "", "100%", "", "0", "0", -1, -2, "", "", Fila, "si")
      Tabla = Tabla & "</center>"

    End If


 '   sHTML = sHTML & Tabla
    Call OFV.cerrarconsulta(rs)

          isecc = 1

          for isecb=1 to isec
              if asecto(isecb) = rsx(1) then 
                 isecc = isecb
                 isecb = 9999
                 exit for
              end if
          next

       if isec = 1 then
            Celda = OFV.GenCelda("", "aaa style='" & borderr & borderl & "'", "", "", "", "", "", Tabla, "si")
       else
            celda = ""
            cbl = borderl
            cbr = borderr
            for isecb = 1 to isec 
                if isecb = 1 then
                   cbl = borderl
                   cbr = ""
                elseif isecb = isec then
                       cbl = cborderls
                       cbr = borderr
                else
                       cbl = cborderls
                       cbr = ""
                end if 
                if isecb = isecc then
                   Celda = Celda & OFV.GenCelda("", "aaa style='" & cbr & cbl & "'", "", "", "", "", "", Tabla, "si")
                else
                   Celda = Celda & OFV.GenCelda("", "aaa style='" & cbr & cbl & "'", "", "", "", "", "", "&nbsp;", "si")
                end if
            next 
       end if
            Filon = Filon & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

'''' fin loop proced
    rsx.movenext
    loop 
    Call OFV.cerrarconsulta(rsx)


'''' fin loop sectores

      Celda = OFV.GenCelda("", "aaa style='" & borderr & borderl & "' colspan=" & isec , "", "", "", "#6699cc", "", "&nbsp;", "si")
      Filon = Filon & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      Tablon = Tablon & "<center>"
      Tablon = Tablon & OFV.GenTabla("", "aaa rules=columns  ", "", tsize & "%", "", "0", "0", -1, -2, "", "", Filon, "si")
      Tablon = Tablon & "</center>"


    Call OFV.cerrarconn(cn)


    sHTML = sHTML & Tablon
    sHTML = sHTML & "</body></html>"




end if

      response.write  sHTML
    
%>