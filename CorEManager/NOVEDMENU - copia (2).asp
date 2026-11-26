
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%


ok = ofv.CheckUsuario()
If ok <> "Y" Then
    uv = ofv.uv
    If uv = "NO" Then ok = ofv.ValidarUsuario
    sHTML = ok
Else
    response.expires = 0
   xxform=""

    session("volver2")=""
    stab=request.querystring("tab")

    if stab="" then stab="view1"
    session("ModMain") = "CL"
    wmnnoved = "NOVED" & wkaplic

    session(wmnnoved) = "SI"
    aplcheck()

    wspviewopc = 2

    stilo1 = " style={background-color:" & session("CMFMENU") & ";border-style:outset;border-width:2;}"
    stilo2 = " style={background-color:" & session("CMFMNBOT") & ";}"


    ReDim Preserve bConsultar(10)

    ''''response.write "<br><br><br><br>" & Session("Form") & "<br>"

    Novedform = "CLIMN16317206150930885546227"

    If (session("UN") & "") = "" Then


        xxxx = NEWEMPRE("NO")


        SEMPRE = ATREMP(0, 0)
        session("UN") = SEMPRE & ""

        If (session("UN") & "") = "" Then session("UN") = 1


    End If


    Set cn = ofv.conectar(ofv.strconn2)
    Set cn0 = ofv.conectar(ofv.strconn0)

    wmnfuncion = ""

    atmensajes = 0
    Dim atitsecciones(50, 5)
    ttitsecc = 0
    acctitblks = 0
    acctitlin = 0

    Dim atit(50)
    Dim awmnasp(50)
    wmntope = 10
    wmntopeX = 10
    WMNINITX = 2

    For tt = 1 To wmntope
        atit(tt) = "NO"
    Next

    If wkaplic = "OR" Or wkaplic = "PL" Then
        '   wmntopeX = 5
        '   WMNINITX = 4
    End If


    titc = 0
    StrSql = "SELECT * FROM novedtit "
    Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
    If Not rsx.EOF Then

        For tt = WMNINITX To wmntopeX
            atit(tt) = rsx(tt)
            If rsx(tt) <> "NO" Then
                ofv.ObtenerAtributosNOdo Novedform, "", "TIPO", CStr(tt)
                bConsultar(tt) = ofv.AConsultar

                If tt = 10 Then

                    If bConsultar(tt) Then
                        session("SMPLMENU") = "N"
                        atit(tt) = "NO"

                    Else
                        If session("SMPLMENU") <> "S" Then
                            session("SMPLMENU") = "S"
                            wmnfuncion = "<script language='javascript' > hdemenu(); </script> "
                        End If
                        titc = titc + 1



                    End If


                Else

                    If bConsultar(tt) Then
                        titc = titc + 1

                    Else

                        atit(tt) = "NO"

                    End If

                End If

            End If
        Next

        If atit(10) <> "NO" Then

            ii = 10
            xxxx = wmnbuscamenu("CLIS150172171122331051096140")
            wmntope = ii

        End If

    End If
    Call ofv.cerrarconsulta(rsx)

    menuSel = request.querystring("menu")
    Menu = menuSel

    If Menu = "" Then

        If atit(3) <> "NO" Then
            'menu = 3
            If wkaplic = "OR" Then
                Menu = 4
                menuSel = 4
            Else
                Menu = 3
                menuSel = 3
            End If


        Else
            Menu = 0
            menuSel = 0

            For iix = 1 To 10
                If atit(iix) <> "NO" Then
                    Menu = iix
                    menuSel = iix
                    Exit For
                End If
            Next
        End If

    End If

    Menu = CInt(Menu)

    If session("FormN") = "" Then session("FormN") = "NOVEDADES"

    If session("SMPLMENU") = "S" Then session("FormN") = "MENU PRINCIPAL"

    sHTML = ofv.FormHeader(session("FormN"))
    shtml=replace(shtml,"</HEAD><Title>","<link href=""template6/tabcontent.css"" rel=""stylesheet"" type=""text/css"" /></HEAD><Title>")

    '' sHTML = ofv.menuheader("","novedades ")

    xs = ""

    xs = "<script language=JavaScript>" & Chr(13)
    xs = xs & "function AbrirCons(Datos,nom) {" & Chr(13)
    xs = xs & " var h=570;w=790 ;"
    if  session("APLMDCLI")="BC" THEN
	xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	END IF
    xs = xs & " myWin= open(Datos"

    xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," '& Chr(34)
' xs = xs & "'menubar=yes,toolbar=yes,height=h,width=w,resizable=yes'" '& Chr(34)
    xs = xs &"  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'" '& CHR(34)
    xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

    s = "<script language=JavaScript>" & Chr(13)
    s = s & "function AbrirVentana(Datosx) {" & Chr(13)
    s = s & " var h=570;w=790 ;"
	 if  session("APLMDCLI")="BC" THEN
	  s = s &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	 END IF

    s = s & " myWin= open(Datosx"
    s = s & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
  '  s = s & "menubar=no,toolbar=no,height=570,width=790,status=yes,resizable=yes" & Chr(34)
  	s = s &"  'status=yes,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'" '& CHR(34)
    s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
    s = s & " } " & Chr(13) & "</script>" & Chr(13)

    sHTML = sHTML & xs & s & wmnfuncion



	sHTML = sHTML & "<script src=""tabcontent.js"" type=""text/javascript""></script>"



    tabla1 = ""
    tablab = ""
    tablan = ""
    tablab1 = ""




    nFecini = 0

    If nFecini < CLng(CDate("01-01-1900")) Then
        nFecini = CLng(CDate("01-01-1900"))
    Else
        nFecini = 0
    End If

    dFechai = CLng(Date) - nFecini
    xTrColor = "#FFFFFF"

    sglmar = "<font class=btiW >*</font>"
    sglman = "<font class=btiW >1</font>"
    sglprf = "<font class=btic >&#0128;</font>"
    sgllnk = "<font class=btic >&#0126;</font>"
    sgladj = "<font class=btic >&#0204;</font>"


    Dim ANTabla(500, 25)


    If session("WEBDESK") = "S" Then


        '''''''''''''' buscador por texto     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        tablab = tablab & "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"

        '''  sInputT  = ofv.GenerarInput("textot","","agregar","","80","dt style='border-style:solid;border-width:1;border-color:#666666;' ","")

        If  session("APLMDCLI")="COR"  OR session("APLMDCLI")="BGAL"   Then
            sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' size='40' style='border-style:solid;border-width:1;border-color:#666666;' >"
        Else
            sInputT = "<input type='text' class='dt' name='textot' id='textot' size='40' style='border-style:solid;border-width:1;border-color:#666666;' >"
        End If

        celda = ""
        celda = celda & ofv.GenCelda("", "UT  ", "right", "30%", "", "", "", sInputT, "si")

        sInput = ofv.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")
        celda = celda & ofv.GenCelda("", "vt", "", "70%", "", "", "", sInput, "si")

        formx = ofv.GenRow("", "", "", "", "", "", "", celda, "si")

        sTarget = ""


        'sAccion  = "consBuscadorSet.asp"
        'sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
        If session("APLMDCLI") = "COR" OR session("APLMDCLI")="BC"   THEN
        sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX"
          'sTarget = "_new"
        ELSEIf session("APLMDCLI") = "BGAL" Then
            sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOR"
            sTarget = "_new"
        Else
            sAccion = "consBuscadorSet.asp"
        End If

        Fila = ofv.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", formx, "si")

        tablab = tablab & ofv.GenTabla("", "aaa STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", Fila, "si")

        tablab = tablab & "</DIV>" & Chr(13)
        'tablab =tablab &" <script language='javascript'>document.Subconsulta.submit();</script>"



        '''''''''''''' indice buscador normativo '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        tabla1 = tabla1 & "<div id='DTTEMA' style='float: left;width:98%; " & Chr(13)
        tabla1 = tabla1 & "border-style:solid;border-width:2;border-color:#ffffff;' >"

        tabla1 = tabla1 & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
        tabla1 = tabla1 & "<tr HEIGHT=25  >"
        xsclase = " CLASS=iett "
        tabla1 = tabla1 & "<td " & xsclase & "  >" & Capitalize("INDICE NORMATIVO") & "</TD>"
        tabla1 = tabla1 & "</TR>"
        tabla1 = tabla1 & "</table>" & Chr(13)

        ''''height:400px;

        tabla1 = tabla1 & "<div id='DTTEMAINT' style='overflow:auto;" & wstyl & ";"
        tabla1 = tabla1 & "border-style:solid;border-width:1;border-color:#dddddd;' >" & Chr(13)

        WSPFRMN = "DTTEMAIT"
        'WSPFRMA = "consdocumindex.asp"
        WSPFRMA = "CONSTEMAIND.asp?FORMATO=ABR"

        WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"

        WSPDTTSCP = ""
        WSPDTTSCP = WSPDTTSCP & "<script FOR=window EVENT=onload  >"
        '' WSPDTTSCP = WSPDTTSCP &  " alert('ENTRO: ');"
        '  WSPDTTSCP = WSPDTTSCP &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
        WSPDTTSCP = WSPDTTSCP & "document." & WSPFRMN & ".submit();"
        '''  WSPDTTSCP = WSPDTTSCP &  "document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").style.display="  & chr(34) & "block" & chr(34) & ";"
        WSPDTTSCP = WSPDTTSCP & "</script>"


        'TABLA1 = Tabla1 & "<script src='../coremanager/consdocumindex.asp?sid=1&target=_self' ></script>"


        tabla1 = tabla1 & "</DIV>"
        tabla1 = tabla1 & "</DIV>" & WSPFORME & Chr(13)


        XTablaN = ""




        xTrColor = "#FFFFFF"

        XCANT = 0



        For ii = 2 To wmntope
            If atit(ii) <> "NO" Then
                XCANT = XCANT + 1

                TNTabla = 0
                INTabla = 0
                xnohay1 = "S"
                xnohay2 = "S"
                saleSel = "N"

                If ii = 10 Then
                    stilo = stilo2
                Else

                    xxxx = noveselec(ii)
                    'response.write "AJA"

                    TNTabla = INTabla

                    XSOPTABN = noveformat(ii, Menu)

                    '  TABLAN = TablaN &  XSOPTABN

                End If

            End If
        Next


        If ttitsecc > 0 Then


            XTablaN = ""

            XTABLAN2 = ""
            ' xsclase = " CLASS=VTI STYLE='background-color:#6666ff;color:#ffffff;font-size:9pt;' "

            xsclase = " class=iett "
            XTablaN = XTablaN & "<div id='DTNOVSEC' style='float: left; "
            XTablaN = XTablaN & "border-style:solid;border-width:2;border-color:#ffffff;' >"


            XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1  style='border-style:solid;border-width:1;border-color:#666666;' >"
            XTablaN = XTablaN & "<tr HEIGHT=25  >"
            XTablaN = XTablaN & "<td " & xsclase & "   >" & Capitalize("secciones") & "</TD>"
            XTablaN = XTablaN & "</TR>"
            XTablaN = XTablaN & "</table>"

            XTablaN = XTablaN & "<div id='DTNOVINTSEC' style='"
            XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#dddddd;' >"

            XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=0  >"

            xxfondo = ""

            colsumd = 0
            colsumi = 0
            ctotspand = 0
            ctotspani = 0
            ctotspanall = 0

            For ii = 1 To ttitsecc



                cworka = CInt(ii / 2)
                cworka = cworka * 2

                If cworka = ii Then
                    colsumd = colsumd + atitsecciones(ii, 2)
                Else
                    colsumi = colsumi + atitsecciones(ii, 2)
                End If



                XTablaN = XTablaN & "<tr HEIGHT='25'  " & xxfondo & " >"
                XTablaN = XTablaN & "<td class=ut width=1% >&nbsp;<B><font color='#f7aa22' size=2 >&#0149;</font></B></TD>"
                XTablaN = XTablaN & "<td class=ut STYLE='color:#000000;font-size:8pt;font-weight:500;' >&nbsp;"
                XTablaN = XTablaN & ""
                XTablaN = XTablaN & atitsecciones(ii, 0) & "&nbsp;(" & atitsecciones(ii, 1) & ")" & ""
                XTablaN = XTablaN & "</TD>"
                XTablaN = XTablaN & "</TR>"

                If xxfondo = "" Then
                    xxfondo = " bgcolor='#ddccff' "
                Else
                    xxfondo = ""
                End If



            Next







            XTablaN = XTablaN & "<tr HEIGHT='20' >"
            XTablaN = XTablaN & "<td class=UT colspan=2  >&nbsp;</TD>"
            XTablaN = XTablaN & "</TR>"

            If atmensajes > 99990 Then

                XTablaN = XTablaN & "<tr HEIGHT='20' bgcolor='#f0eedd' >"
                XTablaN = XTablaN & "<td class=vTi colspan=2 STYLE='color:#000000;font-size:8pt;"
                XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#666666;' >"
                XTablaN = XTablaN & Capitalize("mensajes") & "</TD>"
                XTablaN = XTablaN & "</TR>"

                XTablaN = XTablaN & "<tr HEIGHT='25' >"
                XTablaN = XTablaN & "<td class=ut width=1% >&nbsp;<B><font color='#990000' size=3 >!</font></B></TD>"
                XTablaN = XTablaN & "<td class=ut STYLE='color:#000000;' >&nbsp;"
                XTablaN = XTablaN & "<font color=#666666 style='font-size:8pt;font-weight:700;' >"
                XTablaN = XTablaN & "Ud. tiene&nbsp;" & atmensajes & "&nbsp;mensajes&nbsp;pendientes." & "</font>"
                XTablaN = XTablaN & "</TD>"
                XTablaN = XTablaN & "</TR>"

            End If


            XTablaN = XTablaN & "</table>"
            XTablaN = XTablaN & "</DIV>"
            XTablaN = XTablaN & "</DIV>"

            tablab1 = XTablaN

        End If





    Else

        TNTabla = 0
        INTabla = 0
        xnohay1 = "S"
        xnohay2 = "S"
        saleSel = "N"


        xxxx = noveselec(Menu)

        TNTabla = INTabla


        tablan = ""
        tablan = tablan & "<DIV id='DTNOVE' STYLE='display:BLOCK;z-index:1;margin-top:1;margin-left:1;overflow:auto;'>"


        tablan = tablan & "<table  width=100% cellspacing=0 cellpadding=2 class='HX1CL'  style='border-style:solid;border-width:1;border-color:#ffffff;"
        tablan = tablan & "background-color:#cccccc;"
        tablan = tablan & "' >"

        xnewmenu = ""
        xoldmenu = ""

        tablan = tablan & "<tr >"
        tablan = tablan & "<td class='UT'  >"

        tablan = tablan & "<table   cellspacing=1 cellpadding=3  >"

        For ii = 2 To wmntope
            If atit(ii) <> "NO" Then

                If ii = 10 Then
                    stilo = stilo2
                Else

                    If ii = Menu Then
                        xsclase = "ttr STYLE='border-style:inset;border-color:" & session("CMFPNTITFT") & ";border-width:1;padding-left:5px;padding-right:5px;" & XspEx("#f7aa22", "#dddddd", 0) & ";' "
                    Else
                        xsclase = "ttr STYLE='border-right-style:solid;border-right-color:" & session("CMFPNTITFT") & ";border-right-width:1;padding-left:5px;padding-right:5px;' "
                    End If

                    xmDato = Replace(atit(ii), " ", "&nbsp;")

                    If ii < 11 Then
                        xmhref = "<a class='xts' href='NOVEDMENU.asp?menu=" & ii & "' >" & xmDato & "</a>"
                        xoldmenu = xoldmenu & "<td class=" & xsclase & " align=center  >"
                        xoldmenu = xoldmenu & xmhref & "</TD>"
                    Else
                        xmhref = "<a class='xts' href='" & awmnasp(ii) & "' >" & xmDato & "</a>"
                        xnewmenu = xnewmenu & "<td class=" & xsclase & " align=center  >"
                        xnewmenu = xnewmenu & xmhref & "</TD>"

                    End If



                End If

            End If
        Next

        tablan = tablan & xnewmenu & xoldmenu
        tablan = tablan & "</table>"

        tablan = tablan & "</TD>"
        tablan = tablan & "</TR>"
        tablan = tablan & "</table>"

        '''''''''''''''''''''''''''''
        tablan = tablan & "<center>"

        tablan = tablan & "<div id='WMN0' style='overflow:auto;height:50%;" & wstyl & "' >"

        If xnohay1 = "S" And xnohay2 = "S" Then

            tablan = tablan & "<table  margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"
            tablan = tablan & "<tr  height=200 >"
            tablan = tablan & "<td class=ttt width=100% align=center valign=middle ><big>SIN&nbsp;NOVEDADES&nbsp;DE&nbsp;" & atit(Menu) & "&nbsp;</TD>"
            tablan = tablan & "</TR>"

        Else
            tablan = tablan & "<table bgcolor=#eeeeee margin=3 width=95% cellspacing=0 cellpadding=0  style='border-style:solid;border-width:1;border-color:#eeeeee;'  >"

            For INTabla = 1 To TNTabla

                xtmcod = ANTabla(INTabla, 0)
                xtipom = ANTabla(INTabla, 1)
                xtpri = ANTabla(INTabla, 2)
                xtfvig = ANTabla(INTabla, 3)
                xthvig = ANTabla(INTabla, 4)
                xtsadj = ANTabla(INTabla, 5)
                xtslnk = ANTabla(INTabla, 6)
                xfromusr = ANTabla(INTabla, 7)
                xtmsg = ANTabla(INTabla, 8)
                xtmsgimg = ANTabla(INTabla, 9)
                xtmsglnk = ANTabla(INTabla, 10)

                xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsg)



                Select Case xtipom
                    Case 1
                        xtipomes = "GENERAL"
                    Case 2
                        xtipomes = "PARTICULAR"
                    Case 3
                        xtipomes = "PUBLIC"
                    Case 4
                        xtipomes = "AUTORIZ"
                    Case 5
                        xtipomes = "WORKFLOW"


                End Select

                xfclas = "dt"
                sImage = "&nbsp;"

                Select Case xtpri
                    Case 0

                        xfclas = "btbn"
                        sglmod = 5
                        sgltit = "Prioridad Alta."

                    Case 1

                        xfclas = "btba"
                        sglmod = 4
                        sgltit = "Prioridad Normal."

                    Case 2

                        xfclas = "btbs"
                        sglmod = 6
                        sgltit = "Prioridad baja."


                End Select

                sglmod = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
                sImage = ofv.GenLink("", xfclas & " title='" & sgltit & "' ", "", "", "", "", sglmod, "si")

                If xtfvig = dFechai Then
                    XMARQ = sglmar
                Else
                    XMARQ = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
                End If

                If xtipom = 5 Then
                    xtfechavig = xthvig
                Else
                    xtfechavig = fechaArg(CDate(xtfvig + CLng(CDate("01-01-1900")))) & "&nbsp;&nbsp;" & xthvig
                End If

                Xnlnk = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
                Xnadj = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

                If xtsadj = "S" Then Xnadj = sgladj
                If xtslnk = "S" Then Xnlnk = sgllnk

                Celdak = ""
                Celdak = Celdak & ofv.GenCelda("", "btbs", "", "", "", "", "", sglprf, "si")
                Celdak = Celdak & ofv.GenCelda("", "ut", "", "", "", "", "", "&nbsp;<font color=#999999 ><u>" & xfromusr & "</u></font>&nbsp;&nbsp;", "si")

                Filak = ofv.GenRow("", "", "", "", "", "", "", Celdak, "si")
                sArma = ofv.GenTabla("", "aaa  ", "", "", "", "0", "0", "-1", "-2", "", "", Filak, "si")

                If xtmcod <> "0" Then
                    tablan = tablan & "<tr >"
                    tablan = tablan & "<td class=AAA colspan=8 style='line-height:2pt;' >&nbsp;</TD>"
                    tablan = tablan & "</TR>"



                    tablan = tablan & "<tr bgcolor=" & xTrColor & " >"
                    tablan = tablan & "<td class=" & xfclas & " width=1% bgcolor=#FFFFFF >" & sImage & "</TD>"
                    tablan = tablan & "<td class=vt width=55%  STYLE='text-align:justify;background-color:#DDDDDD;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & session("CMFPNTIT") & ", endColorStr=" & xTrColor & ", gradientType=1);' >&nbsp;&nbsp;<a class=tt >" & xtfechavig & "</a></TD>"
                    tablan = tablan & "<td class=ut width=13% >&nbsp;</TD>"


                    If session("NOVFULL") = "S" Then

                        tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                        tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                        tablan = tablan & "<td class=btba width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnadj & "</TD>"
                        tablan = tablan & "<td class=btbn width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & Xnlnk & "</TD>"
                        tablan = tablan & "<td class=qt width=3% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' >" & XMARQ & "</TD>"
                    Else
                        tablan = tablan & "<td class=UT COLSPAN=3 style='border-style:solid;border-width:1;border-color:#eeeeee;' >&nbsp;</TD>"
                        tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtipomes & "</font></TD>"
                        tablan = tablan & "<td class=ut width=6% align=center style='border-style:solid;border-width:1;border-color:#eeeeee;' ><font color=#999999 SIZE=1 >" & xtmcod & "</font></TD>"
                    End If


                    tablan = tablan & "</TR>"

                End If

                tablan = tablan & "<tr bgcolor=" & xTrColor & " >"
                tablan = tablan & "<td class=uti colspan=8 STYLE='text-align:justify;line-height:30px;' ><font color=#999999 >" & xtmsgimg & "&nbsp;" & xtmsglnk & "</font></TD>"
                tablan = tablan & "</TR>"



                If session("NOVFULL") = "S" Then

                    If xtmcod <> "0" Then


                        tablan = tablan & "<tr bgcolor=" & xTrColor & " height=20 >"
                        tablan = tablan & "<td class=ut colspan=8 valign=top align=right style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#eeeeee;' >" & sArma & "</TD>"
                        tablan = tablan & "</TR>"

                    End If

                Else

                    If xtmcod <> "0" Then


                        tablan = tablan & "<tr bgcolor=" & xTrColor & " height=20 >"
                        tablan = tablan & "<td class=ut colspan=8 valign=top align=right  >&nbsp;</TD>"
                        tablan = tablan & "</TR>"

                    End If

                End If


                If xTrColor = "#FFFFFF" Then
                    xTrColor = "#F9F9F9"
                Else
                    xTrColor = "#FFFFFF"
                End If

            Next


        End If



        tablan = tablan & "</table>"
        tablan = tablan & "</DIV>"
        tablan = tablan & "</div>"



    End If


    'SFAVORITOS=Favoritos()
    'SSUSCRIPCION=SUSCRIPCION()
    Call ofv.cerrarconn(cn)
    Call ofv.cerrarconn(cn0)




   sHTML =shtml & "<div style=""  margin: 0 auto; padding: 0px 0 0px; font: 0.85em arial;"">"

      sHTML =shtml &" <ul class=""tabs"" >"



   if stab="view1" then

    sHTML =shtml &" <li  class=""selected"" ><a class='bietx'  href=""#"" rel=""view1"">NOVEDADES</a></li>"
   else
    sHTML =shtml &" <li   ><a class='bietx'  href=""#"" rel=""view1"">NOVEDADES</a></li>"
   end if




if stab="view2" then

     sHTML =shtml &" <li class=""selected"" ><a class='bietx' href=""#"" rel=""view2"">MIS FAVORITOS</a></li>"
 else
   sHTML =shtml &" <li ><a class='bietx' href=""#"" rel=""view2"">MIS FAVORITOS</a></li>"
 end if



if stab="view3" then
sHTML =shtml &" <li  class=""selected"" ><a class='bietx'  href=""#"" rel=""view3"">MIS SUSCRIPCIONES</a></li>"
else

     sHTML =shtml &" <li   ><a class='bietx'  href=""#"" rel=""view3"">MIS SUSCRIPCIONES</a></li>"
end if
   sHTML =shtml & "</ul>"


    sHTML =shtml &  "<div  class=""tabcontents"" >"
    sHTML =shtml&   "<div id=""view1"" class=""tabcontent"">"
    sHTML = sHTML & "<DIV id='DTGENERAL' STYLE='display: table;' >"

    ' sHTML = sHTML  &  "<DIV id='DTheader' STYLE='float: left; z-index:1;width:22%;' >"

    '  sHTML = sHTML  &  tablab &  tablab1  &  tabla1


    '  sHTML = sHTML  &  "</DIV >"

    ' sHTML = sHTML  &  tablan

    ' sHTML = sHTML  &  "</DIV >"


    If colsumi > colsumd Then
        ctotspanall = colsumi
    Else

        ctotspanall = colsumd
    End If

    cfilltot = 0
    IF session("APLMDCLI") ="BC" THEN

    XTABLAN2 = "<tr ><td class=ut width=35% valign='top'  style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
    XTABLAN21 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
	XTABLAN22 = "<td class=ut width=30% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"

    ELSE
    XTABLAN2 = "<tr ><td class=ut width=30% valign='top'  style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
    XTABLAN21 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
    XTABLAN22 = "<td class=ut width=35% valign='top'       style='border-left-style:solid;border-left-width:1;border-color:#dddddd;' >"
    END IF

    For ii = 1 To ttitsecc

        cworka = CInt(ii / 2)
        cworka = cworka * 2

        If cworka = ii Then
            colsumd = colsumd + atitsecciones(ii, 2)
        Else
            colsumi = colsumi + atitsecciones(ii, 2)
        End If

        xrowspan1 = " rowspan=" & atitsecciones(ii, 2)
        xrowspan1 = " "

        ctotspani = ctotspani + atitsecciones(ii, 2)


        XTABLAN21 = XTABLAN21 & atitsecciones(ii, 3)


        ii = ii + 1
        If ii > ttitsecc Then
            xcolworka = "&nbsp;"
            xrowspan2 = ""
        Else
            xcolworka = atitsecciones(ii, 3)
            xrowspan2 = " rowspan=" & atitsecciones(ii, 2)
            ctotspand = ctotspand + atitsecciones(ii, 2)
            xrowspan2 = ""
        End If

        XTABLAN22 = XTABLAN22 & xcolworka

    Next

   ' if session("APLMDCLI") ="BGAL" THEN
   ' XTABLAN2 = XTABLAN2 & tablab & tablab1 & tabla1 & "</Td>" & XTABLAN21 & "</Td>" & XTABLAN22 & "</Td>" & "</TR>"
   XTABLAN2 = XTABLAN2 & tablab &  tabla1 & "</Td>" & XTABLAN21 & "</Td>" & XTABLAN22 & "</Td>" & "</TR>"




    tablan = "<table width='100%' cellspacing=2 cellpadding=2  style='border-style:solid;border-width:1;border-color:#dddddd;' >"
    tablan = tablan & XTABLAN2
    tablan = tablan & "</table>"



    sHTML = sHTML & tablan
    sHTML = sHTML & "</DIV >"
    sHTML = sHTML &"</div><div id=""view2"" class=""tabcontent"">"
    sHTML = sHTML '& SFAVORITOS

     sHTML = sHTML &"      </div>"
    sHTML = sHTML &    " <div id=""view3"" class=""tabcontent"">"
     sHTML = sHTML '& SSUSCRIPCION
       sHTML = sHTML &"     </div></div>"




    Agregar = ""
    Buscar = ""
    celda = ""

    sHTML = sHTML & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", Agregar, Buscar, celda)

    sHTML = sHTML & WSPDTTSCP


    '' *****  JAVASCRIPT PARA VISTA PRELIMINAR DE NOVEDADES ******
    wstyl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"
    wstyl = wstyl & "border-style:solid;border-width:1;border-color:#000000;background-color:#ffffff;"

    sHTML = sHTML & "<div id='mt2view'  style='Z-INDEX:999999;display:none;position:absolute;top:200;left:300;height:70%;width:50%;overflow:auto;" & wstyl & "' >"

    sHTML = sHTML & "<table width=99% >"
    sHTML = sHTML & "<tr height=10 >"
    sHTML = sHTML & "<td id='mt2viewl1c1' class='tt'  width=90% style='border-style:solid;border-width:1;border-color:#000000;padding-left:5px;background-color:" & session("CMFMENU") & "' >Mostrando vista de </td>"
    sHTML = sHTML & "<td id='mt2viewl1c2'  class='uT' align=center onclick=" & Chr(34) & "document.all.mt2view.style.display='none';" & Chr(34) & " style='cursor:hand;border-style:outset;border-width:1;border-color:#000000;background-color:#eeeeee;' ><h5>Cerrar</h5></td>"
    sHTML = sHTML & "</tr>"

    sHTML = sHTML & "<tr height=100% >"
    sHTML = sHTML & "<td  id='DX4' colspan=2  class='uT' style='background-color:#ffffff;' >contenido</td>"
    sHTML = sHTML & "</tr>"
    sHTML = sHTML & "</table>"

    sHTML = sHTML & "</div >"

    sHTML = sHTML & "<script language=javascript > "
    sHTML = sHTML & "function WKlevanta(wkvista) { " & Chr(13)
    sHTML = sHTML & "  document.all.mt2viewl1c1.innerHTML='<h5> </h5>';" & Chr(13)
    'shtml = shtml & "  document.all.mt2viewl1c1.innerHTML='<h5>Mostrando vista de ' + wkvista + '</h5>';" & chr(13)

    sHTML = sHTML & "  document.all.mt2view.style.display='block';   " & Chr(13)
    sHTML = sHTML & " } " & Chr(13)
    sHTML = sHTML & " </script> "
    '' *****  JAVASCRIPT PARA VISTA PRELIMINAR DE NOVEDADES ******


    sHTML = sHTML & "</body></html>"



    response.write sHTML
    response.write xxform
	 sHTML =  "<script type='text/javascript'>"
	  sHTML = sHTML & " var scw= screen.availWidth;" & Chr(13)
	    sHTML = sHTML & " if (scw <= 800){document.Subconsulta.textot.size=20;} else {document.Subconsulta.textot.size=40;}" & Chr(13)
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

End If

Function EsqnomP(xtipo, xcodigo, ofv)

    Set xcn = ofv.conectar(ofv.strconn2)


    If xtipo = "RSK" Then

        StrSql = "SELECT * FROM  ORRSKCIRCUITO "
        StrSql = StrSql & " Where OBJID = 0" & xcodigo
        Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
        If Not xrs.EOF Then
            If xrs("APLIC") = "OR" Then
                xdescri = "CIRCUITO AUTOEVALUACION"
            Else
                xdescri = "PLAN DE CONTINGENCIAS"
            End If

            StrSql = "Select * From objetos where codobjeto = 0" & xrs("CODOBJETO")
            Set xrs = ofv.crearconsultaEx(StrSql, xcn, 1, parametros)
            If Not xrs.EOF Then
                xiden = xrs(1)
                xobjdes = xrs(3)
            End If


        End If
        Call ofv.cerrarconsulta(xrs)

        sTipobjetoX = Capitalize(xdescri) & ": " & Capitalize(xobjdes)

    Else


        sTipobjetoX = xcodigo

        StrSql2 = "Select * From CSOBJETOSRDO Where CODOBJETO = 0" & xcodigo
        Set RSM = ofv.crearconsultaEx(StrSql2, xcn, 1, parametros)
        If Not RSM.EOF Then




            IF session("APLMDCLI") ="BC" THEN
            sTipobjeto1 = (RSM(1)) & ":&nbsp;"
            sTipobjeto2 = RSM(3) & ""
            sTipobjeto3 = " - " & (RSM(2) & "")
			sTipobjeto4 = " (" & (RSM(5) & "") & " "
            sTipobjeto5 = (RSM(6) & "") & ") "
            ELSE
             sTipobjeto1 = Capitalize(RSM(1)) & ":&nbsp;"
            sTipobjeto2 = Capitalize(RSM(3) & "")

            sTipobjeto3 = " - " & Capitalize(RSM(2) & "")
            sTipobjeto4 = " (" & Capitalize(RSM(5) & "")
            if ucase(RSM(5) & "")="O" THEN
             sTipobjeto5 = ")"
            ELSE
            sTipobjeto5 =  " " & Capitalize(RSM(6) & "") & ") "
            END IF
            END IF
            If session("FULLDOCN") = "1" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2
            ElseIf session("FULLDOCN") = "2" Then
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3
            Else
                sTipobjetoX = sTipobjeto1 & sTipobjeto2 & sTipobjeto3 & sTipobjeto4 & sTipobjeto5
            End If



        End If

    End If





    Call ofv.cerrarconn(xcn)


    EsqnomP = sTipobjetoX


End Function


Function fechaArg(f)
    dia = Day(f)
    mes = Month(f)
    ano = Year(f)
    fechaArg = dia & "/" & mes & "/" & ano
End Function

Function wmnbuscamenu(wmnSubmenu)

    Set cx0 = ofv.conectar(ofv.strconn0)
    Set cx1 = ofv.conectar(ofv.strconn1)

    xsel = 1
    StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & session("CorApli") & "' "

    Set rs = ofv.crearconsultaEx(StrSql, cx1, 1, "")
    If Not rs.EOF Then
        If rs.fields.Count > 6 Then
            xsel = 6
            Call ofv.cerrarconsulta(rs)
            StrSql = "Select * From Menux Where mdepen='" & wmnSubmenu & "' and mcodapli = '" & session("CorApli") & "' order by Morden "

            Set rs = ofv.crearconsultaEx(StrSql, cx1, 1, "")
        Else
            xsel = 1
        End If
    End If
    Do Until rs.EOF



        atrok = "N"
        StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'M' "
        StrSql = StrSql & " And alias = '" & rs(xsel) & "'  "
        StrSql = StrSql & " And usuario = '" & session("segperfil") & "' And "
        StrSql = StrSql & "  Aplicacion = '" & session("CorApli") & "'"

        Set rs1 = ofv.crearconsultaEx(StrSql, cx0, 1, "")
        If rs1.EOF Then
            atrok = "O"
        ElseIf rs1(1) Then
            atrok = "S"
        Else
            atrok = "N"
        End If
        Call ofv.cerrarconsulta(rs1)

        xcount = 0

        If atrok = "O" Then

            StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'M' "
            StrSql = StrSql & " And alias = '" & rs(xsel) & "' And  "

            If session("DELLOG") = "S" Then

                XMFTGRP = Split(session("SegGrpLst"), Chr(9))
                xcount = UBound(XMFTGRP) + 1
                XMFSEP = ""
                StrSql = StrSql & " grupo in ("

                For XMFI = 0 To UBound(XMFTGRP)
                    StrSql = StrSql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                    XMFSEP = ","
                Next

                StrSql = StrSql & ") "

            Else
                StrSql = StrSql & " grupo = '" & session("seggrp") & "' "
            End If


            StrSql = StrSql & " And Aplicacion = '" & session("CorApli") & "'"
            Set rs2 = ofv.crearconsultaEx(StrSql, cx0, 1, "")
            If rs2.EOF Then
                atrok = "O"
            Else
                If session("DELLOG") = "S" Then
                    XMFSUM = 0
                    XMFSUMN = 0
                    Do Until rs2.EOF
                        XMFSUM = XMFSUM + 1
                        If rs2(1) Then
                            atrok = "S"
                            Exit Do
                        Else
                            atrok = "N"
                            XMFSUMN = XMFSUMN + 1
                        End If
                        rs2.Movenext
                    Loop

                    If atrok = "N" Then
                        If xcount > XMFSUMN Then atrok = "S"
                    End If

                Else
                    If rs2(1) Then
                        atrok = "S"
                    Else
                        atrok = "N"
                    End If
                End If
            End If
            Call ofv.cerrarconsulta(rs2)

        End If


        If atrok = "O" Or atrok = "S" Then


            sInput = Trim(rs(xsel))
            sHref = Trim(rs(2))
            sTarget = "info"
            sPath = session("CorPath")
            If rs(3) <> "" Then atrok = "N"
            spt = InStr(rs(2), ".asp")
            xesm = "N"
            If spt > 1 Then
                If Left(rs(2), spt - 1) = "Submenus" Then atrok = "N"
            End If
            sseg = InStr(sHref, "?")
            If sseg = 0 Then
                sHref = sHref & "?"
            Else
                sHref = sHref & "&"
            End If
            sHref = sHref & "seg=" & ofv.convertircar(sInput, " ", "_") & "&MNN=" & ofv.convertircar(UCase(rs(1)), " ", "_")

            If atrok <> "N" Then

                ii = ii + 1

                atit(ii) = UCase(rs(1))
                awmnasp(ii) = sPath & sHref

            End If





        End If




        rs.Movenext
    Loop
    Call ofv.cerrarconsulta(rs)
    Call ofv.cerrarconn(cx1)
    Call ofv.cerrarconn(cx0)




End Function

Function noveselec(xmenu)

    '' xscolor = "#999999 size=2 "
'response.write xmenu
    xscolor = " class=istt "    '"#0000 style='font-size:8pt;font-weight:normal;' "

    Do Until saleSel = "S"

        Select Case xmenu

            Case 1


            Case 2

                '''    session("ModMain") = "CL"

                StrSql = "SELECT * FROM csnovedades "
                StrSql = StrSql & " Where fecexp > " & dFechai
                StrSql = StrSql & " and fecinicio < " & dFechai + 1
                StrSql = StrSql & " And ("
                StrSql = StrSql & " (tipo = 2 And usugrp = '" & session("usuario") & "') or "
                StrSql = StrSql & " (tipo = 1 And usugrp = 'TODOS' and sector = 999 )    or "
                StrSql = StrSql & " (tipo = 1 And usugrp = 'TODOS' and sector = 0" & session("UN") & ") OR "
                StrSql = StrSql & " (tipo = 51 And usugrp = 'TODOS' and sector = 999 )    or "
                StrSql = StrSql & " (tipo = 51 And usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
                StrSql = StrSql & ") "
                StrSql = StrSql & " order by fecinicio desc, tipo desc, importancia, objid desc"
                 'RESPONSE.WRITE STRSQL & chr(13) & "<BR>" & session("ModMain")

                Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rsx.EOF Then
                    xnohay1 = "S"
                Else
                    xnohay1 = "N"
                    Do Until rsx.EOF
                        INTabla = INTabla + 1

                        xfromusr = rsx(11)

                        xxxx = getInfoUser(rsx(11))

                        xfromusr = session("SGU_nombre")

                        ANTabla(INTabla, 0) = rsx(1)
                        ANTabla(INTabla, 1) = CLng(rsx(2))
                        ANTabla(INTabla, 2) = CLng(rsx(6))
                        ANTabla(INTabla, 3) = CLng(rsx(3))
                        ANTabla(INTabla, 4) = rsx(12)
                        ANTabla(INTabla, 5) = "N"
                        ANTabla(INTabla, 6) = "N"
                        ANTabla(INTabla, 7) = xfromusr
                        ANTabla(INTabla, 8) = rsx(5)
                        ANTabla(INTabla, 9) = "&nbsp;"
                        ANTabla(INTabla, 10) = "XXMSGXX"
                        ANTabla(INTabla, 11) = rsx(2)
                        ANTabla(INTabla, 12) = rsx(13)
                        ANTabla(INTabla, 13) = Capitalize(rsx(14))
                        ANTabla(INTabla, 14) = Capitalize(rsx(15))
                        ANTabla(INTabla, 15) = rsx(16)
                        ANTabla(INTabla, 16) = rsx(18)

                        ANTabla(INTabla, 17) = rsx(19) & ""
                        ANTabla(INTabla, 18) = rsx(20) & ""
                        ANTabla(INTabla, 19) = rsx(21) & ""
                        ANTabla(INTabla, 20) = rsx(22) & ""

                        ANTabla(INTabla, 21) = "0"



                        rsx.Movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)

            Case 3, 9

                '''    session("ModMain") = "CL"

                StrSql = "SELECT  * FROM csnovedades "
                StrSql = StrSql & " Where fecexp > " & dFechai
                StrSql = StrSql & " and fecinicio <  " & dFechai + 1
                StrSql = StrSql & " And (tipo = 3 or tipo = 9)"
                StrSql = StrSql & " And ("
                StrSql = StrSql & "    (usugrp = '" & session("usuario") & "') "
                StrSql = StrSql & " or ( usugrp = 'TODOS' and sector = 999)"
                StrSql = StrSql & " or ( usugrp = 'TODOS' and sector = 0" & session("UN") & ")"
                StrSql = StrSql & "     ) "
                StrSql = StrSql & " order by fecinicio desc, importancia, objid desc"
                 '  RESPONSE.WRITE STRSQL & chr(13) & "<BR>"
                Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rsx.EOF Then
                    xnohay1 = "S"

                Else

                    xvinoant = "no"
                    Do Until rsx.EOF


                        If xvinoant <> rsx(1) Then
                            xvinoant = rsx(1)


                            xvale = "N"

                            StrSql = "SELECT * FROM objetos "
                            StrSql = StrSql & " Where codobjeto = 0" & rsx(8)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                            End If
                            Call ofv.cerrarconsulta(rsxp)

                            If xvale = "S" Then

                                xnohay1 = "N"
                                INTabla = INTabla + 1


                                xarc = "<font face=webdings color=#993333 size=4  >&#0157;</font>"


                                sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(8) & "&Tipo=" & rsx(9)
                                sActcon = sActcon & "&Datos=*smenu=zzzzz*lista=3"
                                sHref = "&nbsp;"
                               ' sActcon="http://www.google.com"
                                ' response.write sActcon
                                sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)


                                wmsgx = trim(rsx(5) & "")

                                If Len(wmsgx) < 5 Then
                                    wmsgx = EsqnomP(rsx(9), rsx(8), ofv) '& " VIGENTE A PARTIR DEL " & fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900")))) & ". "
                                Else
                                    'if xmenu=q3 then
                                   ' wmsgx = UCase(wmsgx) ' & "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
                                   ' else'response.write "aa" & len(wmsgx) & rsx(5)
                                    wmsgx = UCase(wmsgx)  '& "&nbsp;"& fechaArg(CDate(CLng(rsx(3)) + CLng(CDate("01-01-1900"))))
                                   ' end if
                                End If

                                wNomL = wmsgx & "&nbsp;" '& Len(wmsgx)



                                sImagen = wNomL

                                If xarc = "NO" Then
                                    xarc = "&nbsp;"
                                End If

                                sLink = sImagen
                                '   XTSLNKM = OFV.GenLink("", "cmb title='Consultar Documento...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                                XTSLNKM = ofv.GenLink("", "istt title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                                xfromusr = rsx(11)


                                xxxx = getInfoUser(rsx(11))

                                xfromusr = session("SGU_nombre")

                                ANTabla(INTabla, 0) = rsx(1)
                                ANTabla(INTabla, 1) = rsx(2)
                                ANTabla(INTabla, 2) = CLng(rsx(6))
                                ANTabla(INTabla, 3) = CLng(rsx(3))
                                ANTabla(INTabla, 4) = rsx(12)
                                ANTabla(INTabla, 5) = "S"
                                ANTabla(INTabla, 6) = "N"
                                ANTabla(INTabla, 7) = xfromusr
                                ANTabla(INTabla, 8) = sLink
                                ANTabla(INTabla, 9) = xarc
                                ANTabla(INTabla, 10) = XTSLNKM
                                ANTabla(INTabla, 11) = xmenu

                                If (rsx(13) & "") = "2" Then
                                    ANTabla(INTabla, 12) = rsx(13)
                                    ANTabla(INTabla, 13) = Capitalize(rsx(14))
                                    If rsx(2) = "9" Then ANTabla(INTabla, 13) =ANTabla(INTabla, 13) & " - " & EsqnomP(rsx(9), rsx(8), ofv)

                                    ANTabla(INTabla, 14) =Capitalize(rsx(15))
                                    ANTabla(INTabla, 15) = rsx(16)
                                    ANTabla(INTabla, 16) = rsx(18)
                                Else

                                    ANTabla(INTabla, 12) = 2
                                    ANTabla(INTabla, 13) = EsqnomP(rsx(9), rsx(8), ofv)
                                    ANTabla(INTabla, 14) = sLink
                                    ANTabla(INTabla, 15) = "0"
                                    ANTabla(INTabla, 16) = ""
                                End If

                                ANTabla(INTabla, 17) = rsx(19) & ""
                                ANTabla(INTabla, 18) = rsx(20) & ""
                                ANTabla(INTabla, 19) = rsx(21) & ""
                                ANTabla(INTabla, 20) = rsx(22) & ""

                                ANTabla(INTabla, 21) = rsx(8)

                            End If
                        End If

                        rsx.Movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)





            Case 4

                '''     session("ModMain") = "WK"
                session("Segce") = "S"
                session("Segco") = "S"

                StrSql = "SELECT * FROM csnovedades "
                StrSql = StrSql & " Where fecexp > " & dFechai
                StrSql = StrSql & " and fecinicio < " & dFechai + 1
                StrSql = StrSql & " And tipo = 4"
                StrSql = StrSql & " And usugrp = '" & session("usuario") & "' order by fecinicio desc, importancia, objid desc"
                'RESPONSE.WRITE STRSQL
                Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rsx.EOF Then
                    xnohay1 = "S"
                Else
                    xnohay1 = "N"
                    Do Until rsx.EOF


                        INTabla = INTabla + 1

                        StrSql = "SELECT archivo FROM tipobjeto "
                        StrSql = StrSql & " Where tipobjeto = '" & rsx(9) & "'"
                        Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                        If Not rs.EOF Then
                            xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
                        Else
                            xarc = "NO"
                        End If
                        Call ofv.cerrarconsulta(rs)


                        session("ASP") = "novedmenu.asp?menu=4"
                        session("sEstado") = "PARA APROBAR"
                        session("sEstadon") = "2"
                        session("Form2") = "apr"

                        aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(8) & "&Tipo=" & rsx(9) & "&Perfil=xxxx&Funcion=E"
                        aspeditar = CStr(aspeditar)
                        sHref = "&nbsp;"
                        ' response.write sActcon
                        sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                        wNomL = "&nbsp;" & Capitalize(rsx(5)) & "&nbsp;"

                        sImagen = wNomL


                        If xarc = "NO" Then
                            xarc = "&nbsp;"
                        End If

                        ''''  sLink = OFV.GenLink("", "cmb", sHref, "", "", sFuncion, "<font color=" & xscolor & " >" & sImagen & "</font>", "si")


                        sLink = sImagen
                        'XTSLNKM = OFV.GenLink("", "cmb ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                        XTSLNKM = ofv.GenLink("", "istt ", sHref, "", "", sFuncion, "XXMSGXX", "si")


                        xfromusr = rsx(11)


                        xxxx = getInfoUser(rsx(11))
                        xfromusr = session("SGU_nombre")

                        ANTabla(INTabla, 0) = rsx(1)
                        ANTabla(INTabla, 1) = 4
                        ANTabla(INTabla, 2) = CLng(rsx(6))
                        ANTabla(INTabla, 3) = CLng(rsx(3))
                        ANTabla(INTabla, 4) = rsx(12)
                        ANTabla(INTabla, 5) = "S"
                        ANTabla(INTabla, 6) = "N"
                        ANTabla(INTabla, 7) = xfromusr
                        ANTabla(INTabla, 8) = sLink
                        ANTabla(INTabla, 9) = xarc
                        ANTabla(INTabla, 10) = XTSLNKM
                        ANTabla(INTabla, 11) = xmenu
                        ANTabla(INTabla, 12) = rsx(13)
                        ANTabla(INTabla, 13) = rsx(14)
                        ANTabla(INTabla, 14) = rsx(15)
                        ANTabla(INTabla, 15) = rsx(16)
                        ANTabla(INTabla, 16) = rsx(18)

                        ANTabla(INTabla, 17) = rsx(19) & ""
                        ANTabla(INTabla, 18) = rsx(20) & ""
                        ANTabla(INTabla, 19) = rsx(21) & ""
                        ANTabla(INTabla, 20) = rsx(22) & ""

                        ANTabla(INTabla, 21) = "0"

                        rsx.Movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)

                '''''''''''''


                Set cn1 = ofv.conectar(ofv.strconn5)
                Set cn4 = ofv.conectar(ofv.strconn6)


                If wkaplic = "OR" Or wkaplic = "PL" Then

                    xtipos = " OBJTIPO = 'RSK' "


                Else

                    xtipos = " OBJTIPO in ("
                    separ = ""

                    StrSql = "SELECT tipobjeto FROM tipobjeto "
                    Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs.EOF Then
                        Do While Not rs.EOF
                            xtipos = xtipos & separ & "'" & rs(0) & "'"
                            separ = ","
                            rs.Movenext
                        Loop
                    Else
                        xtipos = xtipos & "' '"
                    End If
                    xtipos = xtipos & ")"

                End If





                SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & session("Usuario") & "') "


                xxxx = getInfoUser(session("Usuario"))
                xusrlog = session("SGU_nombre")
                'RESPONSE.WRITE xusrlog

                IF xusrlog ="NO" AND session("USRPRV")="LACAJA"  THEN
                    session("USRPRV")="I"
	            	session("Usuario")="USRCONSULTA"
                    call getInfoUser (session("Usuario"))
                   xusrlog = session("SGU_nombre")
                   session("USRPRV")="LACAJA"
				END IF

               ' xxxx = getInfoUser(session("Usuario"))
               ' xusrlog = session("SGU_nombre")


              '  RESPONSE.WRITE xusrlog &" 222222222222222222222222222" & CHR(13)
                If xusrlog <> "NO" Then


                    SEGWHERE = SEGWHERE & "OR ( ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & session("Usuario") & "') ) "
                    SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

                    '''''''''''''''''''''''''''''''''''''''''''
                    If session("DELLOG") = "S" Then

                        XMFTGRP = Split(session("SegGrpLst"), Chr(9))
                        xcount = UBound(XMFTGRP) + 1
                        XMFSEP = ""
                        SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

                        For XMFI = 0 To UBound(XMFTGRP)
                            SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                            XMFSEP = ","
                        Next

                        SEGWHERE = SEGWHERE & ") ) )"

                    Else
                        SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & session("SEGGRP") & "') ) "
                    End If


                    ''''''''''''''''''''''''''''''''''''''''''

                    '''   IF  RS(21) > 0 THEN


                    If CStr(session("SGU_idcostcenter") & "") <> "0" Then
                        SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                        SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                        SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & session("SEGNIVEL") & " ) "
                        SEGWHERE = SEGWHERE & " ) ) "
                    End If


                End If
                If xusrlog <> "NO" THEN
                  SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "
                ELSE
                  SEGWHERE = SEGWHERE & " ) ORDER BY PSCCOMIENZO "
                END IF

                StrSql = "SELECT * FROM CKTVIEW " & SEGWHERE

                'RESPONSE.WRITE STRSQL & "<BR>" & chr(13)
                Set rsx = ofv.crearconsultaEx(StrSql, cn1, 1, parametros)
                If rsx.EOF Then
                    xnohay2 = "S"
                Else

                    Do Until rsx.EOF

                        xvale = "N"

                        If rsx(3) = "RSK" Then

                            StrSql = "SELECT * FROM ORRSKCIRCUITO "
                            StrSql = StrSql & " Where OBJID = 0" & rsx(4)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)


                        Else



                            StrSql = "SELECT * FROM objetos "
                            StrSql = StrSql & " Where codobjeto = 0" & rsx(4)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)

                        End If




                        If xvale = "S" Then

                            xnohay2 = "N"

                            INTabla = INTabla + 1


                            xarc = "<font face=webdings color=#005500 size=4  >&#0157;</font>"

                            session("ASP") = "novedmenu.asp?menu=4"
                            session("sEstado") = "PARA APROBAR"
                            session("sEstadon") = "2"
                            session("Form2") = "apr"

                            aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8) & "&MAINAPL=" & xMAINAPL


                            ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                            aspeditar = CStr(aspeditar)
                            sHref = "&nbsp;"
                            '        response.write sActcon
                            sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                            wNomL = "" & Capitalize(rsx(7)) & "&nbsp;-&nbsp;Paso&nbsp;" & Capitalize(rsx(16)) & "&nbsp;"

                            sImagen = wNomL


                            sImagen = sImagen & "-&nbsp;" & EsqnomP(rsx(3), rsx(4), ofv)

                            If xarc = "NO" Then
                                xarc = "&nbsp;"
                            End If

                            sLink = sImagen
                            'XTSLNKM = OFV.GenLink("", "cmb title='Autorizar documento ...' ", sHref, "", "", sFuncion, "<font color=" & xscolor & " >XXMSGXX</font>", "si")
                            XTSLNKM = ofv.GenLink("", "istt title='Autorizar documento ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                            xfromusr = rsx(6)


                            xxxx = getInfoUser(rsx(6))
                            xfromusr = session("SGU_nombre")

                            ANTabla(INTabla, 0) = "CKT" & INTabla
                            ANTabla(INTabla, 1) = 5
                            ANTabla(INTabla, 2) = 0
                            ANTabla(INTabla, 3) = dFechai
                            ANTabla(INTabla, 4) = rsx(10)
                            ANTabla(INTabla, 5) = "S"
                            ANTabla(INTabla, 6) = "N"
                            ANTabla(INTabla, 7) = xfromusr
                            ANTabla(INTabla, 8) = sLink
                            ANTabla(INTabla, 9) = xarc
                            ANTabla(INTabla, 10) = XTSLNKM
                            ANTabla(INTabla, 11) = xmenu


                            ANTabla(INTabla, 12) = 2
                            ANTabla(INTabla, 13) = EsqnomP(rsx(3), rsx(4), ofv)
                            ANTabla(INTabla, 14) = sLink
                            ANTabla(INTabla, 15) = "0"
                            ANTabla(INTabla, 16) = ""

                            ANTabla(INTabla, 17) = ""
                            ANTabla(INTabla, 18) = ""
                            ANTabla(INTabla, 19) = ""
                            ANTabla(INTabla, 20) = ""

                            ANTabla(INTabla, 21) = "0"

                        End If



                        rsx.Movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)

            Case 5

                '''    session("ModMain") = "CL"

                cuenta = 0

                ''''''''''''''''''''''''''''''''
                XWHERE = " and USUARIO = '" & session("usuario") & "' "
                XWHERE = XWHERE & " AND CODOBJETO > 0 "

                If wkaplic = "OR" Or wkaplic = "PL" Then

                    XWHERE = XWHERE & " AND TIPOBJETO = 'RSK' "

                Else

                    XWHERE = XWHERE & " AND TIPOBJETO <> 'RSK' "

                End If

                XORDER = " group by codobjeto ORDER BY fecha desc "


                StrSql = ""
                StrSql = StrSql & "SELECT CODOBJETO, max(TIPOBJETO), max(finicio) as fecha  "
                StrSql = StrSql & " FROM OBJALOGDET B1"
                StrSql = StrSql & " WHERE TACC = 'CONSEN'  " & XWHERE

                StrSql = StrSql & XORDER

                ' RESPONSE.WRITE STRSQL & "<BR>"

                Set rsx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If rsx.EOF Then
                    xnohay1 = "S"

                Else

                    xvinoant = "no"
                    Do Until rsx.EOF


                        xvale = "N"



                        If wkaplic = "OR" Or wkaplic = "PL" Then


                            StrSql = "SELECT * FROM ORRSKCIRCUITO "
                            StrSql = StrSql & " Where OBJID = 0" & rsx(0)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)

                            WKSPTIPOOBJ = "RSK"


                        Else

                            StrSql = "SELECT * FROM objetos "
                            StrSql = StrSql & " Where codobjeto = 0" & rsx(0)
                            ''    RESPONSE.WRITE STRSQL & "<BR>"
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)


                            WKSPTIPOOBJ = "VER"

                        End If




                        If xvale = "S" Then


                            cuenta = cuenta + 1

                            xnohay1 = "N"
                            INTabla = INTabla + 1


                            xarc = "<font face=webdings color=#993333 size=4  >&#0157;</font>"


                            sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsx(0) & "&Tipo=" & rsx(1) & "&MAINAPL=" & xMAINAPL
                            sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                            sHref = "&nbsp;"
                            ' response.write sActcon
                            ''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

                            sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)


                            sTipobjeto = EsqnomP(WKSPTIPOOBJ, rsx(0), ofv)






                            wmsgx = sTipobjeto & ""
                            wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



                            sImagen = wNomL

                            If xarc = "NO" Then
                                xarc = "&nbsp;"
                            End If

                            sLink = sImagen

                            XTSLNKM = ofv.GenLink("", "nt  title='Consultar Documento...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")

                            ''' XTSLNKM = wNomL

                            xfromusr = "COREMANAGER"


                            xxxx = getInfoUser(xfromusr)

                            xfromusr = session("SGU_nombre")


                            ANTabla(INTabla, 0) = "0"
                            ANTabla(INTabla, 1) = 3
                            ANTabla(INTabla, 2) = 1
                            ANTabla(INTabla, 3) = dFechai
                            ANTabla(INTabla, 4) = 0
                            ANTabla(INTabla, 5) = "S"
                            ANTabla(INTabla, 6) = "N"
                            ANTabla(INTabla, 7) = xfromusr
                            ANTabla(INTabla, 8) = sLink
                            ANTabla(INTabla, 9) = xarc
                            ANTabla(INTabla, 10) = XTSLNKM
                            ANTabla(INTabla, 11) = xmenu
                            ANTabla(INTabla, 12) = "1"
                            ANTabla(INTabla, 13) = ""
                            ANTabla(INTabla, 14) = ""
                            ANTabla(INTabla, 15) = ""
                            ANTabla(INTabla, 16) = ""
                            ANTabla(INTabla, 17) = ""
                            ANTabla(INTabla, 18) = ""
                            ANTabla(INTabla, 19) = ""
                            ANTabla(INTabla, 20) = ""

                            ANTabla(INTabla, 21) = "0"

                        End If

                        If cuenta > 9 Then Exit Do


                        rsx.Movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)



            Case 6

                '''    session("ModMain") = "WK"
                session("Segce") = "S"
                session("Segco") = "S"

                ' RESPONSE.WRITE "6"

                Set cn1 = ofv.conectar(ofv.strconn5)
                Set cn4 = ofv.conectar(ofv.strconn6)

                xtipos = " OBJTIPO in ('RSK'"
                separ = ","
                StrSql = "SELECT tipobjeto FROM tipobjeto "
                Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                If Not rs.EOF Then
                    Do While Not rs.EOF
                        xtipos = xtipos & separ & "'" & rs(0) & "'"
                        separ = ","
                        rs.Movenext
                    Loop
                Else
                    xtipos = xtipos & "' '"
                End If
                xtipos = xtipos & ")"

                SEGWHERE = " WHERE " & xtipos & " AND  PSCFIN = 'NO' AND ( (RESPASG = '" & session("Usuario") & "') OR ("

                xxxx = getInfoUser(session("Usuario"))
                xusrlog = session("SGU_nombre")




                If xusrlog <> "NO" Then


                    SEGWHERE = SEGWHERE & " ( (RESPTIPO = 'USR') AND (PSCRESPONSABLE = '" & session("Usuario") & "') ) "
                    SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'GRP') AND ( "

                    '''''''''''''''''''''''''''''''''''''''''''
                    If session("DELLOG") = "S" Then

                        XMFTGRP = Split(session("SegGrpLst"), Chr(9))
                        xcount = UBound(XMFTGRP) + 1
                        XMFSEP = ""
                        SEGWHERE = SEGWHERE & " PSCRESPONSABLE in ("

                        For XMFI = 0 To UBound(XMFTGRP)
                            SEGWHERE = SEGWHERE & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                            XMFSEP = ","
                        Next

                        SEGWHERE = SEGWHERE & ") ) )"

                    Else
                        SEGWHERE = SEGWHERE & " PSCRESPONSABLE = '" & session("SEGGRP") & "') ) "
                    End If


                    ''''''''''''''''''''''''''''''''''''''''''

                    '''   IF  RS(21) > 0 THEN


                    If CStr(session("SGU_idcostcenter") & "") <> "0" Then
                        SEGWHERE = SEGWHERE & " OR ( (RESPTIPO = 'SEC') AND ( "
                        SEGWHERE = SEGWHERE & " (PSCRESPONSABLE = '" & session("SGU_idcostcenter") & "' OR SUBCODIGO = '" & session("SGU_idcostcenter") & "' ) "
                        SEGWHERE = SEGWHERE & " AND (NOT NIVEL < " & session("SEGNIVEL") & " ) "
                        SEGWHERE = SEGWHERE & " ) ) "
                    End If


                End If

                SEGWHERE = SEGWHERE & " ) ) ORDER BY PSCCOMIENZO "

                StrSql = "SELECT * FROM CKTVIEW " & SEGWHERE

                '  RESPONSE.WRITE STRSQL & "4444<BR>"
                Set rsx = ofv.crearconsultaEx(StrSql, cn1, 1, parametros)
                If rsx.EOF Then
                    xnohay2 = "S"
                Else

                    Do Until rsx.EOF

                        xvale = "N"

                        '   Strsql  = "SELECT * FROM objetos "
                        '   Strsql  = Strsql & " Where codobjeto = 0" & rsx(4)
                        '   set rsxp  = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                        '   if  not rsxp.eof then
                        '       xvale = "S"
                        '   end if
                        '   call ofv.cerrarconsulta(rsxp)

                        If rsx(3) = "RSK" Then

                            StrSql = "SELECT * FROM ORRSKCIRCUITO "
                            StrSql = StrSql & " Where OBJID = 0" & rsx(4)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)


                        Else



                            StrSql = "SELECT * FROM objetos "
                            StrSql = StrSql & " Where codobjeto = 0" & rsx(4)
                            Set rsxp = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rsxp.EOF Then
                                xvale = "S"
                                xMAINAPL = rsxp("APLIC")
                            End If
                            Call ofv.cerrarconsulta(rsxp)

                        End If


                        If xvale = "S" Then

                            xnohay2 = "N"

                            INTabla = INTabla + 1

                            StrSql = "SELECT archivo FROM tipobjeto "
                            StrSql = StrSql & " Where tipobjeto = '" & rsx(3) & "'"
                            Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rs.EOF Then
                                '         xarc = "<IMG SRC=Iconos\" & rs(0) & " style={border-style:outset;border-width:2;}>"
                                '      else
                                '          xarc = "NO"
                            End If
                            Call ofv.cerrarconsulta(rs)

                            xarc = "<font face=webdings color=#005500 size=4  >&#0157;</font>"

                            session("ASP") = "novedmenu.asp?menu=4"
                            session("sEstado") = "PARA APROBAR"
                            session("sEstadon") = "2"
                            session("Form2") = "apr"

                            aspeditar = "Editar.asp?DOCMODAL=WK&Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Perfil=xxxx&Funcion=CD&SEGCKTPSC=" & rsx(8)


                            ''''  aspeditar = "Consultar.asp?Codigo=" & rsx(4) & "&Tipo=" & rsx(3) & "&Datos=*smenu=PORESTADO*lista=3&SEGCKTPSC=" & RSX(8)

                            aspeditar = CStr(aspeditar)
                            sHref = "&nbsp;"
                            '        response.write sActcon
                            sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & aspeditar & "'); return false;" & Chr(34)

                            wNomL = "" & Capitalize(rsx(7)) & "&nbsp;-&nbsp;PASO&nbsp;" & Capitalize(rsx(16)) & "&nbsp;"

                            sImagen = wNomL


                            sImagen = sImagen & "-&nbsp;" & EsqnomP(rsx(3), rsx(4), ofv)


                            If xarc = "NO" Then
                                xarc = "&nbsp;"
                            End If

                            sLink = sImagen
                            XTSLNKM = ofv.GenLink("", "istt title='Ver mensaje ...' ", sHref, "", "", sFuncion, "XXMSGXX", "si")


                            xfromusr = rsx(6)


                            xxxx = getInfoUser(rsx(6))
                            xfromusr = session("SGU_nombre")


                            '' ANTabla(INTabla,0) = "0"
                            ANTabla(INTabla, 0) = "MISPEN" & INTabla
                            ANTabla(INTabla, 1) = 5
                            ANTabla(INTabla, 2) = 0
                            ANTabla(INTabla, 3) = dFechai
                            ANTabla(INTabla, 4) = rsx(10)
                            ANTabla(INTabla, 5) = "S"
                            ANTabla(INTabla, 6) = "N"
                            ANTabla(INTabla, 7) = xfromusr
                            ANTabla(INTabla, 8) = sLink
                            ANTabla(INTabla, 9) = xarc
                            ANTabla(INTabla, 10) = XTSLNKM
                            ANTabla(INTabla, 11) = xmenu
                            ANTabla(INTabla, 12) = 2
                            ANTabla(INTabla, 13) = EsqnomP(rsx(3), rsx(4), ofv)
                            ANTabla(INTabla, 14) = sLink
                            ANTabla(INTabla, 15) = "0"
                            ANTabla(INTabla, 16) = ""

                            ANTabla(INTabla, 17) = ""
                            ANTabla(INTabla, 18) = ""
                            ANTabla(INTabla, 19) = ""
                            ANTabla(INTabla, 20) = ""

                            ANTabla(INTabla, 21) = "0"

                        End If



                        rsx.Movenext
                    Loop
                End If
                Call ofv.cerrarconsulta(rsx)





            Case 10

                '''   session("ModMain") = "CL"

                xnohay1 = "N"


                sLink = "<INPUT TYPE=BUTTON VALUE='HIDE MENU' onclick='hdemenu();' ><INPUT TYPE=BUTTON VALUE='SHOW MENU' onclick='mostrarmenu();' >"


                INTabla = INTabla + 1


                ANTabla(INTabla, 0) = "0"
                ANTabla(INTabla, 1) = 3
                ANTabla(INTabla, 2) = 1
                ANTabla(INTabla, 3) = dFechai
                ANTabla(INTabla, 4) = 0
                ANTabla(INTabla, 5) = "S"
                ANTabla(INTabla, 6) = "N"
                ANTabla(INTabla, 7) = "COREMANAGER"
                ANTabla(INTabla, 8) = sLink
                ANTabla(INTabla, 9) = "&nbsp;"
                ANTabla(INTabla, 10) = "&nbsp;"
                ANTabla(INTabla, 11) = xmenu
                ANTabla(INTabla, 12) = "1"
                ANTabla(INTabla, 13) = ""
                ANTabla(INTabla, 14) = ""
                ANTabla(INTabla, 15) = ""
                ANTabla(INTabla, 16) = ""


                ANTabla(INTabla, 17) = ""
                ANTabla(INTabla, 18) = ""
                ANTabla(INTabla, 19) = ""
                ANTabla(INTabla, 20) = ""
                ANTabla(INTabla, 21) = "0"








            Case Else


        End Select

        If menuSel <> "" Then
            saleSel = "S"
        Else
            If INTabla > 0 Then
                saleSel = "S"
            Else
                menuSel = 2
                Menu = 2
            End If
        End If



    Loop

End Function

Function noveformat(xmenu, XMENSEL)


    XTablaN = ""

    acctitblks = 0
    acctitlin = 0


    'xsclase = " CLASS=VTI STYLE='background-color:#f0eedd;color:#000000;font-size:9pt;' "

    xsclase = " CLASS=ietx "
    xsclase2 = "class=istt "
    '''   XTABLAN  = XTablaN & "<div id='DTNOV" & xmenu & "' style='display: inline ;float: left; "

    XTablaN = XTablaN & "<div id='DTNOV" & xmenu & "' style='width:100%; "


    If xmenu = 2 Then

        '    XTABLAN = XTablaN & " width:40%;"

    Else

        '   XTABLAN = XTablaN & " width:30%;"

    End If

    '********* ENCABEZADO DE SECCIONES
    XTablaN = XTablaN & "border-style:solid;border-width:3;border-color:#ffffff;' >"
    XTablaN = XTablaN & "<table width=100% cellspacing=0 cellpadding=1    >"
    XTablaN = XTablaN & "<tr HEIGHT=25  >"
    XTablaN = XTablaN & "<td " & xsclase & "   >" & Replace(Capitalize(atit(xmenu)) & "", " ", "&nbsp;")  &  "</TD>"
    XTablaN = XTablaN & "</TR>"
    XTablaN = XTablaN & "</table>"

    XTablaN = XTablaN & "<div id='DTNOVINT" & xmenu & "' style='"

    XTablaN = XTablaN & "border-style:solid;border-width:1;border-color:#ffffff;' >"

    XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=0  >"


    If TNTabla = 0 Then


        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;TEST NOVEDADES&nbsp;DE&nbsp;" & atit(Menu) & "&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"

        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"


    Else

        For INTabla = 1 To TNTabla

            '    acctitlin  = acctitlin + 2


            xtmcod = ANTabla(INTabla, 0)
            xtipom = ANTabla(INTabla, 1)
            xtpri = ANTabla(INTabla, 2)
            xtfvig = ANTabla(INTabla, 3)
            xthvig = ANTabla(INTabla, 4)
            xtsadj = ANTabla(INTabla, 5)
            xtslnk = ANTabla(INTabla, 6)
            xfromusr = ANTabla(INTabla, 7)
            xtmsg = ANTabla(INTabla, 8)
            xtmsgimg = ANTabla(INTabla, 9)
            xtmsglnk = ANTabla(INTabla, 10)
            xtmsgTYPE = ANTabla(INTabla, 11)



            xfclas = "dt"
            sImage = "&nbsp;"

            Select Case xtpri
                Case 0

                    xfclas = "btbn"
                    sglmod = 5
                    sgltit = "Prioridad Alta."

                Case 1

                    xfclas = "btba"
                    sglmod = 4
                    sgltit = "Prioridad Normal."

                Case 2

                    xfclas = "btbs"
                    sglmod = 6
                    sgltit = "Prioridad baja."


            End Select

            sglmod = "<div class='" & xfclas & "' ><font class=btic >" & sglmod & "</font></div>"
            sImage = ofv.GenLink("", xfclas & " title='" & sgltit & "' ", "", "", "", "", sglmod, "si")

            If xtfvig = dFechai Then
                XMARQ = sglmar
            Else
                XMARQ = "<font color=#999999 SIZE=1 >" & sglman & "</font>"
            End If

            If xtipom = 5 Then
                xtfechavig = xthvig
                XTCOUNT = InStr(xtfechavig, " ")
                If XTCOUNT > 0 Then
                    xtfechavig = Left(xtfechavig, XTCOUNT)
                    xtfechavig = GTFechadorsh(CLng(CDate(xtfechavig)) - CLng(CDate("01-01-1900")))
                End If
            Else
                xtfechavig = GTFechadorsh(xtfvig)
                XTCOUNT = InStr(xtfechavig, " ")
                If XTCOUNT > 0 Then xtfechavig = Left(xtfechavig, XTCOUNT)


            End If

            Xnlnk = "<font color=#dddddd SIZE=1 >" & sgllnk & "</font>"
            Xnadj = "<font color=#dddddd SIZE=1 >" & sgladj & "</font>"

            If xtsadj = "S" Then Xnadj = sgladj
            If xtslnk = "S" Then Xnlnk = sgllnk









            xtmsgVRS = "0" & ANTabla(INTabla, 12)
            xtmsgVRS = CInt(xtmsgVRS)
            xtmsgTIT = ANTabla(INTabla, 13)
            xtmsgCOP = ANTabla(INTabla, 14)
            xtmsgimg = ANTabla(INTabla, 15)
            xtmsgVID = ANTabla(INTabla, 16)


            xtmsgVAENC = ANTabla(INTabla, 17)
            xtmsgVAFORO = ANTabla(INTabla, 18)
            xtmsgENCCOD = ANTabla(INTabla, 19)
            xtmsgFOROCOD = ANTabla(INTabla, 20)

            xtmsgdoccod = ANTabla(INTabla, 21)


            If xtmsgVRS = 2 Then

                acctitlin = acctitlin + 2

                'xtmsglnk    = replace(xtmsglnk,"XXMSGXX","<font STYLE='color:#990000;font-size:10pt;font-weight:bold;' >" & xtmsgTIT & "</font>")
                xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

                '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
                '  XTABLAN = XTablaN &  "<td class=UT colspan=8  >&nbsp;</TD>"
                '  XTABLAN = XTablaN &  "</TR>"

                XTablaN = XTablaN & "<tr HEIGHT='20' >"
                XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
                For yy = 1 To 6
                    XTablaN = XTablaN & "<td class=UT width='15%'  >&nbsp;</TD>"
                Next
                XTablaN = XTablaN & "<td class=UT width='5%'  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"

                ''  TIPO DE NOVEDAD NOTICIA DEL DIA
                If xtipom = "1" Then

                    xSolasp = ""
                    xSolasp = xSolasp & "novedmenuview.asp?xxmsg=" & CStr(xtmcod) & "&Pantalla=MAIN"
                    XTablaN = XTablaN & "<form name='VNGP" & CStr(xtmcod) & "' action='" & xSolasp & "' "
                    XTablaN = XTablaN & " method=post target='fesquemas' style='display:none;' >"
                    ' XTablaN = XTablaN & " method=post target='_new' style='display:none;' >"
                    XTablaN = XTablaN & "</form>"



                    xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsgTIT)

                    XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"
                    XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8   >" & xtmsglnk & "</a></TD>"
                    XTablaN = XTablaN & "</TR>"


                    '' Otros tipos de novedad
                Else



                    XTablaN = Celda2 & XTablaN & "<tr HEIGHT='30' >"
                    XTablaN = XTablaN & "<td " & xsclase2 & " colspan=8  >" & xtmsglnk & "</TD>"

                    XTablaN = XTablaN & "</TR>"

                End If


                If xtipom = "9" Then

                    acctitlin = acctitlin + 2

                    XTMSGSEND = ""
                    XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP

                    XTablaN = XTablaN & "<tr  >"
                    XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B>"
                    XTablaN = XTablaN & "<font color='#AA0000' size=2 >&#0187;</font>"
                    XTablaN = XTablaN & "<font color='#0000AA' size=1 >&#0124;</font>"
                    XTablaN = XTablaN & "</B></TD>"
                    XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;text-align:justify;' >"
                    XTablaN = XTablaN & XTMSGSEND & "</TD>"
                    XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"


                    XTMSGSEND = ""
                    XTMSGSEND = XTMSGSEND & "<font color=#888888 style='font-weight:bold;' size=1 >recomendado por " & ANTabla(INTabla, 7) & "</font>"

                    XTablaN = XTablaN & "<tr HEIGHT='20' >"
                    XTablaN = XTablaN & "<td class=UT align='right' colspan=8 STYLE='background-color:#dddddd;' >" & XTMSGSEND & "&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"

                End If


                '''''''''''''''''''''''''''''
                XTablaN = XTablaN & "<tr HEIGHT='20' >"
                XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"


                If  xtmsgdoccod <> "0" Then



                    StrSql = "Select * From objinfo where codobjeto = 0" & xtmsgdoccod & " "
                    ' response.write strsql
                    Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs1.EOF Then
'aca esta el error
                        xtmsgCOP =xtmsgCOP &"<br><br>" &rs1(4) '& xtmsgdoccod
                        xtmsgimg = rs1(7)
                        xtmsgVID = rs1(9)

                    End If

                    StrSql = "Select * From objinfoIMG  where codobjeto = 0" & xtmsgdoccod & " ORDER BY objSECUEN, OBJID"

                    wspimgdat = 6

                Else

                    StrSql = "Select * From NOVEDADESIMG  where NOVEDAD = '" & xtmcod & "' ORDER BY SECUEN, OBJID"
                    wspimgdat = 5
                End If



                If (xtmsgimg & "") = "1" Then

                    '   response.write Strsql

                    Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs1.EOF Then



                        XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
                        XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"

                        XTablaN = XTablaN & "<table cellspacing=0 cellpadding=1  width=100% >"
                        XTablaN = XTablaN & "<tr   >"


                        XSFSTYLE = " STYLE='display:block;' "

                        Do While Not rs1.EOF

                            acctitlin = acctitlin + 10

                            xCelda = ""
                            xCelda = xCelda & "<div id='DIMGN" & INTabla & "' style='margin-top:1;margin-left:1;"
                            xCelda = xCelda & "width:90%;overflow:auto;"
                            xCelda = xCelda & ""
                            xCelda = xCelda & "' >" & Chr(13)
                            xCelda = xCelda & "<IMG style='border-style:solid;border-width:1;border-color:#888888;' SRC="&chr(34) & rs1(wspimgdat) & chr(34) &">"
                            xCelda = xCelda & "</div>" & Chr(13)

                            XTablaN = XTablaN & "<td CLASS=UT align='center' " & XSFSTYLE & " width=100% >" & xCelda & "</TD>"
                            XSFSTYLE = " STYLE='display:none;' "


                            rs1.Movenext
                        Loop

                        XTablaN = XTablaN & "</TR>"
                        XTablaN = XTablaN & "</table>"

                        XTablaN = XTablaN & "</TD>"
                        XTablaN = XTablaN & "</TR>"


                        XTablaN = XTablaN & "<tr HEIGHT='20' >"
                        XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
                        XTablaN = XTablaN & "</TR>"


                    End If






                ElseIf (xtmsgimg & "") = "2" Then

                    acctitlin = acctitlin + 10


                    XTablaN = XTablaN & "<tr bgcolor='#F0F0F0' >"
                    XTablaN = XTablaN & "<td class=ut colspan=8 ALIGN='CENTER' STYLE='color:#000000;' >"
                    XTablaN = XTablaN & xtmsgVID
                    '   XTABLAN = XTablaN &  REPLACE(xtmsgVID & "",CHR(34),CHR(34) & CHR(34))
                    XTablaN = XTablaN & "</TD>"
                    XTablaN = XTablaN & "</TR>"


                    XTablaN = XTablaN & "<tr HEIGHT='20' >"
                    XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"

                End If

                acctitlin = acctitlin + 2

                '**** COPETE NOTICIA DEL DIA
                XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
                XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                XTMSGSEND = XTMSGSEND & "&nbsp;" & xtmsgCOP
                XTablaN = XTablaN & "<tr  >"
                XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' ><B><font color='#AA0000' size=2 >&#0149;</font></B></TD>"
                XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='text-align:justify;' >"
                ' LINK VER MAS
                If xtipom = "1" Then
                    url = "novedmenuview.asp?xxmsg=" & xtmcod
                    sFuncion = " onclick=" & Chr(34) & "document.all.VNGP" & CStr(xtmcod) & ".submit();WKlevanta('" & CStr(xtmcod) & "');" & Chr(34)
                    xtmsglnk = ofv.GenLink("", " istt style='cursor:hand ;' title=' Ver Mas ...' ", "", "", "", sFuncion, "Ver Mas", "si")
                    XTablaN = XTablaN & XTMSGSEND & "<br> <br>" & xtmsglnk
                else
                XTABLAN = XTablaN &   XTMSGSEND & "</TD>"
                End If
               ' XTablaN = XTablaN & "</TD>"







                ''' ********
                XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"




                ''''' PREGUNTA SI VA ENCUESTA '''''''''''''''''''''''

                '    xtmsgVAENC      = ANTabla(INTabla,17)
                '    xtmsgVAFORO     = ANTabla(INTabla,18)
                '    xtmsgENCCOD     = ANTabla(INTabla,19)
                '    xtmsgFOROCOD    = ANTabla(INTabla,20)



                If xtmsgVAENC = "S" Then

                    acctitlin = acctitlin + 2

                    StrSql = "Select * From emencuestas  where objid = 0" & xtmsgENCCOD
                    Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                    If Not rs1.EOF Then

                        xtmsgTITENC = Capitalize(rs1(17) & "")
                        xtmsgENCDER = "S"


                        If (rs1(2) & "") = "1" Then
                            StrSql = "Select * From EMENCUESTASBLK WHERE ENCcodigo = 0" & xtmsgENCCOD
                            StrSql = StrSql & " Order By BLKSECUEN,objid DESC"
                            Set rs1 = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
                            If Not rs1.EOF Then
                                sblktvoto = rs1(4) & ""
                                sblkcodigo = rs1(0) & ""
                                If sblktvoto = "2" Then
                                    xtmsgENCDER = "N"
                                    xtmsgTITENC = Capitalize(rs1(3) & "")



                                    XTablaN = XTablaN & "<tr  >"
                                    XTablaN = XTablaN & "<td class=ut  VALIGN='TOP' align='center' >&nbsp;</TD>"
                                    XTablaN = XTablaN & "<td class=ut colspan=6 STYLE='color:#000000;' align='center' ><br>"

                                    XTablaN = XTablaN & "<table width=99% cellspacing=0 cellpadding=1  style='background-color:#eeeeee;border-style:solid;border-width:1;border-color:#dddddd;' >"

                                    XTablaN = XTablaN & "<tr HEIGHT=20  >"
                                    XTablaN = XTablaN & "<td class=vtI COLSPAN=3  >Encuesta</TD>"
                                    XTablaN = XTablaN & "</TR>"


                                    XTablaN = XTablaN & "<tr HEIGHT=20  >"
                                    XTablaN = XTablaN & "<td class=utI width=85% style='font-size:8pt;font-weight:normal;' >" & xtmsgTITENC & "</TD>"

                                    WSPDIVN = "DENC" & xtmsgENCCOD

                                    StrSql2 = "Select * From EMENCUESTASITM WHERE ENCcodigo = 0" & xtmsgENCCOD
                                    StrSql2 = StrSql2 & " and BLKcodigo = 0" & sblkcodigo
                                    StrSql2 = StrSql2 & " Order By ITMSECUEN,objid DESC"

                                    Set rs2 = ofv.crearconsultaEx(StrSql2, cn, 1, parametros)
                                    Do While (Not rs2.EOF)

                                        acctitlin = acctitlin + 2


                                        WSPFRMN = "FENC" & xtmsgENCCOD & "IT" & rs2(0)
                                        WSPFRMA = "EMENCUESTASACT.ASP?TP=ES&OP=ACT&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
                                        WSPFRMR = "EMENCUESTASACT.ASP?TP=ES&OP=RES&ENC=" & xtmsgENCCOD & "&BLK=" & sblkcodigo & "&ITM=" & rs2(0)
                                        XTablaN = XTablaN & "<td class=ut width=5% align='center' >"
                                        WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='fesquemas' style='display:none;' METHOD='POST' ></FORM>"
                                        XTablaN = XTablaN & "<input type='button' style='width:45px;background-color:#cc0000;color:#ffffff;font-weight:bold;border-style:solid;border-width:1;border-color:#888888;' "
                                        XTablaN = XTablaN & "  value='&nbsp;" & rs2(4) & "&nbsp;' "
                                        XTablaN = XTablaN & "  onclick='"
                                        '  XTABLAN = XTablaN &  "alert("" voto:" & rs2(4) & " "");"
                                        '  XTABLAN = XTablaN &  "alert(document.getElementById(" & chr(34) & WSPDIVN & chr(34) & ").id);"
                                        XTablaN = XTablaN & "document." & WSPFRMN & ".submit();"
                                        XTablaN = XTablaN & "document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";"
                                        XTablaN = XTablaN & "' >"
                                        XTablaN = XTablaN & "</TD>" & WSPFORME
                                        rs2.Movenext
                                    Loop

                                    acctitlin = acctitlin + 2

                                    xCelda = ""
                                    xCelda = xCelda & "<div id='" & WSPDIVN & "' style='z-index:1;margin-top:1;margin-left:1;display:none;"
                                    xCelda = xCelda & "width:99%;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
                                    xCelda = xCelda & "background-color:#F9F9F9;"
                                    xCelda = xCelda & "' >" & Chr(13)
                                    xCelda = xCelda & "&nbsp;"
                                    xCelda = xCelda & "</div>" & Chr(13)


                                    XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"
                                    XTablaN = XTablaN & "</TR>"

                                    XTablaN = XTablaN & "<tr HEIGHT=20  >"
                                    XTablaN = XTablaN & "<td class=utI width=85%  >&nbsp;</TD>"
                                    XTablaN = XTablaN & "<td class=vt COLSPAN=2 ALIGN='CENTER' ><A CLASS=NT HREF='" & WSPFRMR & "' TARGET='fesquemas' "
                                    XTablaN = XTablaN & " ONCLICK='document.getElementById(" & Chr(34) & WSPDIVN & Chr(34) & ").style.display=" & Chr(34) & "block" & Chr(34) & ";' "
                                    XTablaN = XTablaN & "><B>Resultados</B></TD>"
                                    XTablaN = XTablaN & "<td class=ut width=5% >&nbsp;</TD>"

                                    XTablaN = XTablaN & "</TR>"

                                    XTablaN = XTablaN & "<tr   >"
                                    XTablaN = XTablaN & "<td class=vtI COLSPAN=4 >&nbsp;</TD>"
                                    XTablaN = XTablaN & "</TR>"



                                    XTablaN = XTablaN & "</table>" & xCelda
                                    XTablaN = XTablaN & "</TD>"
                                    XTablaN = XTablaN & "<td class=ut   >&nbsp;</TD>"
                                    XTablaN = XTablaN & "</TR>"





                                End If

                            End If


                        End If





                    End If

                End If


                ''''''''''''''''''''''''''''

                XTablaN = XTablaN & "<tr HEIGHT='20' >"
                XTablaN = XTablaN & "<td class=UT colspan=8 STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"











            Else

                acctitlin = acctitlin + 1


                ''''  xtmsg    = replace(ltrim(xtmsg & "")," ","&nbsp;")
                xtmsg = LTrim(xtmsg & "")
                xtmsglnk = Replace(xtmsglnk, "XXMSGXX", xtmsg & "")


                If xtmcod = "0" Then

                    XTMSGSEND = "<font color=#000000 >" & xtmsglnk & "</font>"

                Else

                    XTMSGSEND = "<B><font color=#000000 >" & xtfechavig & "</font></B>"
                    XTMSGSEND = XTMSGSEND & "<B><font color='#0000AA' size=2 >&#0124;</font></B>"
                    XTMSGSEND = XTMSGSEND & "<font color=#000000 >&nbsp;" & xtmsglnk & "</font>"

                End If



                XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " HEIGHT='25' >"
                XTablaN = XTablaN & "<td class=" & xfclas & " align='center' >" & xtmsgimg & "</TD>"
                XTablaN = XTablaN & "<td class=uti colspan=6 STYLE='text-align:justify;' >" & XTMSGSEND & "</TD>"
                XTablaN = XTablaN & "<td class=UT  >&nbsp;</TD>"
                XTablaN = XTablaN & "</TR>"


                If xtmcod <> "0" Then

                    ' acctitlin  = acctitlin + 1


                    XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=5 >"
                    XTablaN = XTablaN & "<td class=ut colspan=8 valign=top align=right style='border-top-style:solid;border-top-width:1;border-top-color:#eeeeee;' >&nbsp;</TD>"
                    XTablaN = XTablaN & "</TR>"

                End If


            End If


            If xTrColor = "#FFFFFF" Then
                xTrColor = "#eeeeff"
            Else
                xTrColor = "#FFFFFF"
            End If


            '  XTABLAN = XTablaN &  "<tr HEIGHT='20' >"
            '  XTABLAN = XTablaN &  "<td class=UT colspan=8  ><HR  size=1 style='width:400;color:#ff0000;' >&nbsp;</TD>"
            '  XTABLAN = XTablaN &  "</TR>"




        Next


    End If

    XTablaN = XTablaN & "<tr HEIGHT='20' >"
    XTablaN = XTablaN & "<td class=UT colspan=8  >&nbsp;</TD>"
    XTablaN = XTablaN & "</TR>"



    XTablaN = XTablaN & "</table>"
    XTablaN = XTablaN & "<HR  size=1 style='width:350;color:#eeeeee;' >"
    XTablaN = XTablaN & "</DIV>"

    XTablaN = XTablaN & "</DIV>"



    If TNTabla = 0 Then
        noveformat = ""
    Else

        If xmenu = 4 Then atmensajes = TNTabla

        ttitsecc = ttitsecc + 1

        atitsecciones(ttitsecc, 0) = Capitalize(atit(xmenu))
        atitsecciones(ttitsecc, 1) = TNTabla
        atitsecciones(ttitsecc, 2) = acctitlin + 2
        atitsecciones(ttitsecc, 3) = XTablaN
        atitsecciones(ttitsecc, 4) = 0
        atitsecciones(ttitsecc, 5) = 0

        noveformat = XTablaN
    End If


End Function


function Favoritos()
dim XTablaN
strsql = "SELECT * "
strsql = strsql &  "FROM csnovedades   a "
'strsql = strsql  & "inner join objmsg b on( a.codobjeto=b.MSGOBJCOD and b.MSGTIPOMESS=2 and  b.usr = '" & SESSION("USUARIO") & "')"
xTrColor =""
set rsfav=ofv.crearconsultaEx(StrSql,cn,1,parametros)

 XTablaN =  "<table  width=80% cellspacing=0 cellpadding=0  >"


    If not rsfav.eof  = 0 Then


        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;FAVORITOS&nbsp;"& "&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"

        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"


    Else

      do while not rsfav.eof

       CLAVE= rsfav(0)
	   If xTrColor = "#FFFFFF" Then
		xTrColor = "#eeeeff"
	   Else
		xTrColor = "#FFFFFF"
       End If

               XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=25 >"
               sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsfav("MSGOBJCOD") & "&Tipo=" & rsfav("MSGOBJTIP") & "&MAINAPL=" & xMAINAPL
			   sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			   sHref = "&nbsp;"
			   ' response.write sActcon
			   ''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

			   sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)

               IF RSFAV("ESTADO") & "" ="3" THEN

                TITULOLNK="Consultar Documento..."

               ELSE
               sFuncion=""
			   sHref="#"
			   TITULOLNK="EL Documento NO encuentra vigente debe eliminarse de Favoritos..."
               END IF

			   sTipobjeto = EsqnomP(rsfav("MSGOBJTIP"), rsfav("MSGOBJCOD"), ofv)



			   wmsgx = sTipobjeto & ""
			   wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



			   sImagen = wNomL

			   If xarc = "NO" Then
				   xarc = "&nbsp;"
			   End If

			   sLink = sImagen

               XTSLNKM = ofv.GenLink("", "nt  title='"& TITULOLNK &"' ", sHref, "", "", sFuncion, wmsgx, "si")

               sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
               VLK="../novedmenu.asp?tab=view2"
			   sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=novegenpar&volver=" & vlk
			   sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

			   Celda = BotonEliminar("", sHref,"fav" & CLAVE )


                XTablaN =XTablaN  &  CELDA

                    XTablaN = XTablaN & "<td class=ut STYLE='color:#000000;text-align:justify;' >"
                    XTablaN = XTablaN & XTSLNKM & "</TD>"

                    XTablaN = XTablaN & "</TR>"













     rsfav.movenext
        loop


    End If




    XTablaN = XTablaN & "</table>"
Favoritos=XTablaN

end function


function SUSCRIPCION()
dim XTablaN
'Strsql=" Select *,B.ESTADO   From objmsg a "
'STRSQL =STRSQL  & " LEFT JOIN OBJPRTY B ON(A.MSGOBJCOD= B.CODOBJETO) "
'strsql=strsql & " where a.MSGTIPOMESS=2  AND usr = '" & SESSION("USUARIO") & "'"
strsql = "SELECT * "
strsql = strsql &  "FROM csnovedades   a "
strsql = strsql  & "inner join objmsg b on( a.codobjeto=b.MSGOBJCOD and b.MSGTIPOMESS=2 and  b.usr = '" & SESSION("USUARIO") & "')"

 StrSql = StrSql & " Where a.fecinicio > " & dFechai-1
 StrSql = StrSql & " and a.fecinicio <  " & dFechai + 6
StrSql = StrSql & " And a.tipo = 3 "
'response.write strsql
xTrColor =""
' vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID
 'response.write vuelta
set rsfav=ofv.crearconsultaEx(StrSql,cn,1,parametros)

 XTablaN =  "<table  width=80% cellspacing=0 cellpadding=0  >"


    If not rsfav.eof  = 0 Then


        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=VT width=100% COLSPAN=8 align=center valign=middle ><br><br>SIN&nbsp;SUSCRIPCIONES&nbsp;"& "&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"

        XTablaN = XTablaN & "<tr>"
        XTablaN = XTablaN & "<td class=ut colspan=8 >&nbsp;</TD>"
        XTablaN = XTablaN & "</TR>"


    Else

      do while not rsfav.eof

       CLAVE=  rsfav(0)
	   If xTrColor = "#FFFFFF" Then
		xTrColor = "#eeeeff"
	   Else
		xTrColor = "#FFFFFF"
       End If

               XTablaN = XTablaN & "<tr bgcolor=" & xTrColor & " height=25 >"
               sActcon = "Consultar.asp?DOCMODAL=CL&Codigo=" & rsfav("MSGOBJCOD") & "&Tipo=" & rsfav("MSGOBJTIP") & "&MAINAPL=" & xMAINAPL
			   sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
			   sHref = "&nbsp;"
			   ' response.write sActcon
			   ''   sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

			   sFuncion = " onclick=" & Chr(34) & " AbrirVentana('" & sActcon & "'); return false;" & Chr(34)

          '     IF RSFAV("ESTADO") & "" ="3" THEN
'
          '      TITULOLNK="Consultar Documento..."
'
          '     ELSE
          ''     sFuncion=""
		'	   sHref="#"
		'	   TITULOLNK="EL Documento NO se encuentra vigente debe eliminarse de la Suscripcion..."
          '     END IF

			   sTipobjeto = EsqnomP(rsfav("MSGOBJTIP"), rsfav("MSGOBJCOD"), ofv)



			   wmsgx = sTipobjeto & ""
			   wNomL = "&deg;&nbsp;&nbsp;" & wmsgx & "&nbsp;"



			   sImagen = wNomL

			   If xarc = "NO" Then
				   xarc = "&nbsp;"
			   End If

			   sLink = sImagen

               XTSLNKM =wmsgx & " (Se ha Modificado)"' ofv.GenLink("", "nt  title='"& TITULOLNK &"' ", sHref, "", "", sFuncion, wmsgx, "si")

               sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
               vlk="../novedmenu.asp?tab=view3"
			   sHref = "asp/EliminarEN.asp?ID=" & CLAVE & "&Tabla=novegenpar&volver=" & vlk
			   sHref = sHref & "&Codobjeto=" & rsfav("MSGOBJCOD") & "&Tipobjeto=" & rsfav("MSGOBJTIP") & "&Eliminar=SI"

			   Celda = "" 'BotonEliminar("", sHref, "sus" &CLAVE)


                XTablaN =XTablaN  &  CELDA

                    XTablaN = XTablaN & "<td class=ut STYLE='color:#000000;text-align:justify;' >"
                    XTablaN = XTablaN & XTSLNKM & "</TD>"

                    XTablaN = XTablaN & "</TR>"













     rsfav.movenext
        loop


    End If




    XTablaN = XTablaN & "</table>"
SUSCRIPCION=XTablaN

end function



Public Function BotonEliminar(sNombre, sHref,CLAVE)

              wktitdelete="Confirma la eliminacion "
              sImageELS    = "<font face='wINGdings' size=3 color='#bb0000'>&#0253;</font>"
             sformobj   = "FRMOBJ" & CLAVE
             sAccion    = sHref
             formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
             sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

            BotonEliminar=ofv.GenCelda("","ut WIDTH=2% style='cursor:hand;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs ,"si")
            xxform=xxform & formobj


End Function




%>



