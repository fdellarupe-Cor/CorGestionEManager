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
    Cstilo   = "border-style:solid;border-width:1;border-color:#dddddd;"
    Cbordert = "border-top-style:solid;border-top-width:1;border-top-color:#dddddd;"
    Cborderl = "border-left-style:solid;border-left-width:1;border-left-color:#dddddd;"
    Cborderr = "border-right-style:solid;border-right-width:1;border-right-color:#dddddd;"
    Cborderb = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;"

    cborderls = "border-left-style:solid;border-left-width:2;border-left-color:#888888;"

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

  StrSqlx = "Select Version From Objprty Where codobjeto = " & sOCesque
  Set rs = OFV.crearconsultaEx(StrSqlx, cn, 1, "")
  bVersion = "NO"
  If Not rs.EOF Then
    If rs(0) = "V" Then bVersion = "SI"
  End If
  Call OFV.cerrarconsulta(rs)

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


  'sHTML = ofv.FormHeader(session("FormN"))

     shtml = ofv.FormHeader(session("FormN"))


 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"


 '******************************************************************
'* COMIENZO CONFIGURACION DEL ENCABEZADO DE LA PAGINA EL 04/05/2010
'******************************************************************


'Response.Write corstyler2()

'Response.Write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
'Response.Write "<tr>"
'Response.Write "<td width=100% >"
'Response.Write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"), "_", " "), "600")
'Response.Write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
'Response.Write "<tr>"
'Response.Write "<td width=10% >&nbsp;</td>"
'Response.Write "<td width=90% >&nbsp;</td>"
'Response.Write "</tr></table>"

'Encabezados

Formulario = Chr(13) & " <TR ><TD  class=tt width=2% >Tipo</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & Chr(13)


'******************************************************************
'* COMIENZO CONFIGURACION DEL ENCABEZADO DE LA PAGINA EL 04/05/2010
'******************************************************************

    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " myWin= open(Datos"
    xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," & Chr(34)
    xs = xs & "menubar=no,toolbar=no,height=300,width=1500,scrollbars=yes,resizable=yes" & Chr(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13)
    xs = xs & "function pointcon(xcon) {" & Chr(13)

    xs = xs & " window.scrollTo(document.getElementById(xcon)"
    xs = xs & ".offsetParent.offsetLeft,document.getElementById(xcon)"
    xs = xs & ".offsetParent.offsetParent.offsetParent.offsetTop + document.getElementById(xcon).offsetParent.offsetTop);"




    xs = xs & Chr(13) & " } " & Chr(13)
    xs = xs & "</script>" & Chr(13)

    sHTML = sHTML & xs



   ReDim asecto(0)

   asecto(0) = 0
   isec = 0
   isecb = 0

'''' loop sectores

    StrSql = "SELECT * from cursmaster "
    StrSql = StrSql & " Where codobjeto = " & sOCesque
    StrSql = StrSql & " order by secuen, secuencia "
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    Do Until rs.EOF
       If isec = 0 Then
          isec = isec + 1
          asecto(0) = asecto(0) + 1
          ReDim Preserve asecto(isec)
          asecto(isec) = rs(1)
       Else
          For isecb = 1 To isec
              If asecto(isecb) = rs(1) Then
                 isecb = 9999
                 Exit For
              End If
          Next
          If isecb <> 9999 Then
             isec = isec + 1
             asecto(0) = asecto(0) + 1
             ReDim Preserve asecto(isec)
             asecto(isec) = rs(1)
          End If
       End If
    rs.movenext
    Loop
    Call OFV.cerrarconsulta(rs)



   ReDim alnkv(0)
   ReDim alnkl(0)
   ReDim alnkvD(0)
   ReDim alnklD(0)


   ilnkv = 0
   ilnkvb = 0
   ilnkl = 0
   ilnklb = 0

    StrSql = "SELECT * from cursmasterlnk "
    StrSql = StrSql & " Where codobjeto = " & sOCesque & " and ltipo = 'R' "
    StrSql = StrSql & " order by sece, secp, seca"
    Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
    Do Until rs.EOF
          ilnkv = ilnkv + 1
          ReDim Preserve alnkv(ilnkv)
          ReDim Preserve alnkvD(ilnkv)
          alnkv(ilnkv) = rs(1)
          daccD = "&nbsp;"
          daccP = " "
          daccE = "&nbsp;"

          StrSql = "Select codproced,secuen,descri From Cursograma Where codaccion = '" & rs(2) & "'"
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccD = rs2(1) & "." & rs2(2)
             daccP = rs2(0)
          End If
          Call OFV.cerrarconsulta(rs2)


          StrSql = "SELECT  nombre "
          StrSql = StrSql & " from CSCEventoSectorPerfil Where codproced = '" & daccP & "' "
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccE = rs2(0)
             daccE = Replace(daccE, " ", "&nbsp;")
          End If
          Call OFV.cerrarconsulta(rs2)

          alnkvD(ilnkv) = "&nbsp;A:&nbsp;" & daccE & Chr(13) & daccD

          ilnkl = ilnkl + 1
          ReDim Preserve alnkl(ilnkl)
          ReDim Preserve alnklD(ilnkl)
          alnkl(ilnkl) = rs(2)

          daccD = "&nbsp;"
          daccP = " "
          daccE = "&nbsp;"

          StrSql = "Select codproced,secuen,descri From Cursograma Where codaccion = '" & rs(1) & "'"
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccD = rs2(1) & "." & rs2(2)
             daccP = rs2(0)
          End If
          Call OFV.cerrarconsulta(rs2)


          StrSql = "SELECT  nombre "
          StrSql = StrSql & " from CSCEventoSectorPerfil Where codproced = '" & daccP & "' "
          Set rs2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
          If Not rs2.EOF Then
             daccE = rs2(0)
             daccE = Replace(daccE, " ", "&nbsp;")
          End If
          Call OFV.cerrarconsulta(rs2)

          alnklD(ilnkl) = "&nbsp;DE:&nbsp;" & daccE & Chr(13) & daccD

    rs.movenext
    Loop
    Call OFV.cerrarconsulta(rs)



'''' loop proced
     Tablon = ""
     filon = ""
     Celda = ""
     tsize = 100

     If isec > 2 Then tsize = 40 * isec
     if tsize >100 then tsize=100

  sCaption = " CURSOGRAMA:&nbsp;" & sCaption & " "


            For isecb = 1 To isec
                sInput = "&nbsp;&nbsp;" & asecto(isecb)
               ' Celda  = Celda & OFV.GenCelda("", "ttt style='background-color:#aaaaff;color:#000000;font-size:8pt;'  ", "", "", "", "", "", sInput, "si")
                 Celda  = Celda & OFV.GenCelda("", "ttr ", "", "", "", "", "", sInput, "si")
            Next

            filon = filon & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


 aeven = "xxx"


    StrSql = "SELECT * from cursmaster "
    StrSql = StrSql & " Where codobjeto = " & sOCesque
    StrSql = StrSql & " order by secuen, secuencia "

    Set rsx = OFV.crearconsultaEx(StrSql, cn, 1, "")
    Do Until rsx.EOF
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


    If aeven <> sEvento Then
       aeven = sEvento
       sInput = "<big><b>EVENTO:&nbsp;" & sESecu & "&nbsp;" & sEvento & "</b></big>"

       If isec = 1 Then
          Celda = OFV.GenCelda("", "ttr style='" & Cstilo & "' ", "center", "", "", "", "", sInput, "si")
       Else
          Celda = OFV.GenCelda("", "ttr style='" & Cstilo & "' colspan=" & isec & " ", "center", "", "", "", "", sInput, "si")
       End If

       filon = filon & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")
    End If

      Fila     = ""
      Tabla    = ""


' ''''''''''''''
    Set cnn = OFV.conectar(OFV.strconn2)
     Set cnx = OFV.conectar(OFV.strconn2)
    StrSql = "Select * From Cursograma Where codproced = '" & rsx(5) & "'"
    StrSql = StrSql & " and (modi = 0 or ((modi = -1 or modi = 1) and tipmodi <> 'E')) "
    StrSql = StrSql & " order by secuen, codaccion"
    Set rs = OFV.crearconsultaEx(StrSql, cnn, 1, "")




      'Encabezado



      sCaption = ""
      sCaption = sCaption & "PERFIL/FUNCION:&nbsp;&nbsp;" & sPerfil
      sInput   = " " & sCaption & " "
     ''' Celda    = OFV.GenCelda("", "ttt", "center", "", "", "", "", sInput, "si")
     ''' Fila     = OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

      Tabla = ""
    If Not rs.EOF Then

      'Titulos


      Fila  = ""
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
      Celda = OFV.GenCelda("", "ttt colspan=9 style='border-style:solid;border-width:1;border-color:#000000;' ", "center", "", "", "", "", sInput, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


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
        Set rsAlt = OFV.crearconsultaEx(StrSql, cnx, 1, "")
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




        sarma2 = "&nbsp;"
        sComponente = "&nbsp;"
        stcomp = ""
        sacomp = "ICorAbierto.jpg"
        If  rs(18) <> "" And rs(18) > "0" Then
          'Imagen de Complemento




          Strsql = "Select * From objlnk Where ltipo = 'KL' and objid = 0" & rs(18)
           set rsx1 = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
           if not rsx1.eof then

              wtipoL   = rsx1(6)
              wtipoLc  = rsx1(7) & ""

              wNV1    = rsX1(11) & ""
              wNV2    = rsX1(12) & ""
              wNV3    = rsX1(13) & ""

              IF  WNV1 = "" THEN WNV1 = "0"
              IF  WNV2 = "" THEN WNV2 = "0"
              IF  WNV3 = "" THEN WNV3 = "0"

          StrSql = "Select * From tipocomp Where tipo = '" & wtipoLc & "' "
          Set rsComponentes = OFV.crearconsultaEx(StrSql, cnx, 1, "")
          If  Not rsComponentes.EOF Then
              stcomp = rsComponentes(1) & ":&nbsp;"
              sacomp = rsComponentes(2)
          End If
          Call OFV.cerrarconsulta(rsComponentes)


          Strsql = "Select nombre From csacomple Where codcomp = " & wtipoL
          set rsComponentes = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
          if  not rsComponentes.eof then
              sComponente = rsComponentes(0)
              SArma       = replace(sComponente," ","&nbsp;")  & nCopias & ""
              sHref       = "EditarImagen2.asp?ID=" & wtipoL
              sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

              sImage = OFV.GenImage("", "", "", "", "Iconos/" & sacomp, "0", "")

              slink  = OFV.GenLink("", "", sHref, "_blank", "", "", sImage, "si")

              sarma2 = sComponente
              sarma2 = slink & "&nbsp;&nbsp;" & sarma2

          end if
          call ofv.cerrarconsulta(rsComponentes)

          end if
          call ofv.cerrarconsulta(rsx1)






        End If





        'Icono de la Accion


        'Secuencia
        sinputs = rs(3) & "&nbsp;" & slnkv & slnkl


       'Referencia
       sVerbo = rs(4)
        sAcciont = rs(4)
        i = InStr(sVerbo, " ")
        If  i > 0 Then
            sVerbo = Left(rs(4), i - 1)
            sAcciont = Mid(rs(4), i + 1)
        End If



        sInputr = "<font color='#000000' >" & rs(3) & ": "
        sInputr = sInputr & "<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;" & sAcciont & "</FONT>"

        If  rs(16) > "0" Then
        ' Sector Continuador
            StrSql = "select * from eco where eco.eco = " & rs(16)
            Set rsc = OFV.crearconsultaEx(StrSql, cnn, 1, "")
            If Not rsc.EOF Then sconti = rsc(1)
        End If
        If  rs(17) > "0" Then
        ' Perfil Continuador
        StrSql = "select * from tipoperfil where tipoperfil = '" & rs(17) & "'"
            Set rsd = OFV.crearconsultaEx(StrSql, cnn, 1, "")
            If Not rsd.EOF Then sperfi = rsd(1)
        End If
        If  rs(16) > "0" Or rs(17) > "0" Then
            sArma = "Continuador:"
            sArma = sArma & sconti & " - " & sperfi
        Else
            sArma = "&nbsp;"
        End If

        'Call OFV.cerrarconn(cnn)

' '''''''''''''

     For ilnklb = 1 To ilnkl
         If alnkl(ilnklb) = rs(12) Then

               slnkl = "<font color='#FFFF00' face=arial-narrow size=1><b>" & ilnklb & "</b></font>"
               pmkl = "<form name=" & Chr(34) & "D" & ilnklb & "L" & Chr(34) & " method=post >"
               pmkl = pmkl & "<input type=hidden value=a ></form>"









          dtitu = " title='" & alnklD(ilnklb) & "' "

          sImagen = OFV.GenImage("", "", "", "", "Iconos/icorcconecva.jpg", "0", "")
      slnklt = "<A " & dtitu & " onclick=" & Chr(34) & "pointcon('D" & ilnklb & "V');" & Chr(34) & "  STYLE=" & Chr(34) & "cursor:pointer;" & Chr(34) & ">" & sImagen & "</A>"
      slnklt2 = "<A " & dtitu & " onclick=" & Chr(34) & "pointcon('D" & ilnklb & "V');" & Chr(34) & "  STYLE=" & Chr(34) & "cursor:pointer;position:relative;top:-12px;left:-20px" & Chr(34) & ">" & slnkl & "</A>"
      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut rowspan=2  colspan=2  ", "center valign=middle ", "", "", "", "", slnklt & slnklt2, "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   ", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;" & pmkl, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & Cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   style='" & Cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")





'     end if

'            exit for
         End If
     Next

' '''''''''''''''


      sImage = "&nbsp;"
      If bVersion = "SI" And rs(14) Then
      If rs(15) = "M" Then
        sImage = OFV.GenImage("", "", "", "", "Imagenes/Botmodx.jpg", "0", "")
      ElseIf rs(15) = "A" Then
        sImage = OFV.GenImage("", "", "", "", "Imagenes/Botagrx.jpg", "0", "")
      End If
      End If



              sActcon = "Coprocaccion.asp?codobjeto=" & rs(0) & "&Tipobjeto=" & rs(1)
              sActcon = sActcon & "&bversion=" & bVersion & "&baccion=" & rs(12) & "&nID=" & rs(2)
              sHref = "&nbsp;"

              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & trim(rs(1) &"") & "'); return false;" & Chr(34)


           sImagen1 = sInputr
           sImagen2 = OFV.GenImage("", "", "", "", "Iconos/" & rs(5), "0", "")

           sLink1 = OFV.GenLink("", "cc", sHref, "", "", sFuncion, sImagen1, "si")
           sLink2 = OFV.GenLink("", "ttt", sHref, "", "", sFuncion, sImagen2, "si")



      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", sImage, "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4 colspan=4", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut rowspan=2", "", "", "", "", "", "<b>" & sLink1 & "</b><br><br>" & sarma2, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      sImagen = OFV.GenImage("", "", "", "", "Iconos/" & rs(5), "0", "")
      Celda = Celda & OFV.GenCelda("", "ut rowspan=2  colspan=2 style='" & bstilo & "'", "center", "", "", "", "", sLink2, "si")
      Celda = Celda & OFV.GenCelda("", "ut", "left", "", "", "", "", "&nbsp;" & sinputs, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f3
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=2", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut rowspan=2", "", "", "", "", "", sArma, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f4
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")


' '''''''''''''

     For ilnkvb = 1 To ilnkv
         If alnkv(ilnkvb) = rs(12) Then
 '           ilnkvt = 99999


  '   if ilnkvt = 99999 then
        slnkv = "<font color=#ffffff face=arial-narrow size=1><b>" & ilnkvb & "</b></font>"
        pmkv = "<form name=" & Chr(34) & "D" & ilnkvb & "V" & Chr(34) & " method=post >"
        pmkv = pmkv & "<input type=hidden value=a ></form>"
 '    end if

 '    if ilnkvt = 99999 then
  '       slnkv = ">" & ilnkvb

          sImagen = OFV.GenImage("", "", "", "", "Iconos/icorcconecllega.jpg", "0", "")

          dtitu = " title='" & alnkvD(ilnkvb) & "' "

      slnkvt = "<A " & dtitu & " onclick=" & Chr(34) & "pointcon('D" & ilnkvb & "L');" & Chr(34) & "  STYLE=" & Chr(34) & "cursor:pointer;" & Chr(34) & ">" & sImagen & "</A>"
      slnkvt2 = "<A " & dtitu & " onclick=" & Chr(34) & "pointcon('D" & ilnkvb & "L');" & Chr(34) & "  STYLE=" & Chr(34) & "cursor:pointer;position:relative;top:-12px;left:-22px" & Chr(34) & ">" & slnkv & "</A>"
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
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & Cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa   style='" & Cbordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "left", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")






 '     end if

 '           exit for
         End If
     Next


' '''''''''''''''


        'Set cnn = OFV.conectar(OFV.strconn2)
        StrSql = "Select * From AccionesAlt Where codaccion = '" & rs(12) & "'"
        StrSql = StrSql & " and (modi = 0 or ((modi = -1 or modi = 1) and tipmodi <> 'E')) "
           StrSql = StrSql & " Order By secuencia "

        Set rsAlt = OFV.crearconsultaEx(StrSql, cn, 1, "")

           nAlt = ""
           bordert = Cbordert
           borderr = Cborderr

        xsensi = "<div style='border-width:1;border-style:SOLID;BORDER-COLOR:#DDDDDD;' ><font COLOR=#336699 size=1 >SI</font>" & "<font face=corflaticon COLOR=#336699 >6</font></div>"
        xsenno = "<div style='border-width:1;border-style:SOLID;BORDER-COLOR:#DDDDDD;' ><font COLOR=#aa0000 size=1 >NO</font>" & "<font face=corflaticon COLOR=#aa0000 >4</font></div>"

        Do Until rsAlt.EOF

          'Icono de la Accion Alt
          sImagen = OFV.GenImage("", "", "", "", "Iconos/" & rsAlt(7), "0", "")


          'Secuencia Alt
          sinputs = rs(3) & "."
          sinputs = sinputs & rsAlt(1)


          'Referencia Alt
          sVerbo = rsAlt(0)
          sAcciont = rsAlt(0)
          i = InStr(sVerbo, " ")
          If i > 0 Then
             sVerbo = Left(rsAlt(0), i - 1)
             sAcciont = Mid(rsAlt(0), i + 1)
          End If
          sInputr = "<font color='#000000' >" & rs(3) & "."
          sInputr = sInputr & rsAlt(1) & ": "
          sInputr = sInputr & "<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;" & sAcciont & "</FONT>"


          If nAlt = "" Or nAlt = "0" Then
                If Not IsNull(rsAlt(5)) Then nAlt = rsAlt(5)
          End If

          naccion = rsAlt(10)

      sImage = "&nbsp;"
      If bVersion = "SI" And rsAlt(8) Then
      If rsAlt(9) = "M" Then
        sImage = OFV.GenImage("", "", "", "", "Imagenes/Botmodx.jpg", "0", "")
      ElseIf rsAlt(9) = "A" Then
        sImage = OFV.GenImage("", "", "", "", "Imagenes/Botagrx.jpg", "0", "")
      End If
      End If


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
            sArma = nSec


          End If



              sActcon = "Coprocaccionalt.asp?codobjeto=" & rs(0) & "&Tipobjeto=" & rs(1) & "&bsec=" & rs(3)
              sActcon = sActcon & "&bversion=" & bVersion & "&naccion=" & rs(12) & "&nID=" & rs(2) & "&baccion=" & naccion
              sHref = "&nbsp;"

              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & trim(rs(1) &"")& "'); return false;" & Chr(34)


           sImagen1 = sInputr
           sImagen2 = sImagen

           sLink1 = OFV.GenLink("", "cc", sHref, "", "", sFuncion, sImagen1, "si")
           sLink2 = OFV.GenLink("", "nt", sHref, "", "", sFuncion, sImagen2, "si")

''        xsensi = ""
  ''      xsenno = ""


      'f1
      Celda = OFV.GenCelda("", "", "", "", "", "", "", sImage, "si")
      Celda = Celda & OFV.GenCelda("", "ut", "", "", "", "", "", xsensi & "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & bordert & "'", "center", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut valign=top rowspan=4  style='" & bordert & "'", "", "", "", "", "", xsenno & "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderr & bordert & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa rowspan=4  ", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut rowspan=2", "", "", "", "", "", sLink1, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f2
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut rowspan=2  colspan=2 style='" & bstilo & "'", "center", "", "", "", "", sLink2, "si")
      Celda = Celda & OFV.GenCelda("", "ut", "left", "", "", "", "", "&nbsp;" & sinputs, "si")
      Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
      'f3
      Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "aaa style='" & borderl & "'", "", "", "", "", "", "&nbsp;", "si")
      Celda = Celda & OFV.GenCelda("", "ut", "left", "", "", "", "", "&nbsp;" & sArma, "si")
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

        xsensi = ""
        xsenno = ""


        Loop
        Call OFV.cerrarconsulta(rsAlt)



        rs.movenext
        Loop


        Call OFV.cerrarconn(cnn)
        Call OFV.cerrarconn(cnx)

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
      Tabla = Tabla & OFV.GenTabla("", "aaa ", "", "98%", "", "0", "0", -2, 1, "", "", Fila, "si")
      Tabla = Tabla & "</center>"

    End If


 '   sHTML = sHTML & Tabla
    Call OFV.cerrarconsulta(rs)

          isecc = 1

          For isecb = 1 To isec
              If asecto(isecb) = rsx(1) Then
                 isecc = isecb
                 isecb = 9999
                 Exit For
              End If
          Next

       If isec = 1 Then
            Celda = OFV.GenCelda("", "aaa style='" & borderr & borderl & "'", "", "", "", "", "", Tabla, "si")
       Else
            Celda = ""
            cbl = borderl
            cbr = borderr
            For isecb = 1 To isec
                If isecb = 1 Then
                   cbl = borderl
                   cbr = ""
                ElseIf isecb = isec Then
                       cbl = cborderls
                       cbr = borderr
                Else
                       cbl = cborderls
                       cbr = ""
                End If
                If isecb = isecc Then
                   Celda = Celda & OFV.GenCelda("", "aaa style='" & cbr & cbl & "'", "", "", "", "", "", Tabla, "si")
                Else
                   Celda = Celda & OFV.GenCelda("", "aaa style='" & cbr & cbl & "'", "", "", "", "", "", "&nbsp;", "si")
                End If
            Next
       End If
            filon = filon & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      Fila  = ""
      Tabla = ""

'''' fin loop proced
    rsx.movenext
    Loop
    Call OFV.cerrarconsulta(rsx)


'''' fin loop sectores

      Celda = OFV.GenCelda("", "aaa style='" & Cstilo & "' colspan=" & isec, "", "", "", "", "", "&nbsp;", "si")
   '   filon = filon & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

      Tablon = Tablon & "<center><br>"
      Tablon = Tablon & OFV.GenTabla("", "aaa rules=columns style='" & Cstilo & "' ", "", tsize & "%", "", "0", "0", 1, 1, "", "", filon, "si")
      Tablon = Tablon & "</center><br>"


    Call OFV.cerrarconn(cn)


    sHTML = sHTML & Tablon






'''''''''''''''''''''''''''''
'''''''
  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


    sHTML = sHTML & "</body></html>"








end if

response.write sHTML

%>