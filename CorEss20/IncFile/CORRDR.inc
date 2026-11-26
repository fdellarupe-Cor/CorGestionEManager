 <%

' ****************************************************************************************
' ** Cor E-Solution v48 -    Modulo de Visualizacion  ------------------ enero, 2020   **
' ****************************************************************************************

XGBinput1   = ""
XGBinput2   = ""
XGBinput3   = ""
XGBinput4   = ""
XGBinput5   = ""
XGBinput6   = ""
XGBinput7   = ""

Styl    = ""

SfidkA  = ""
SfidkC  = ""
SfidkB  = ""

Sfidk0  = ""
Sfidk1  = ""
Sfidk2  = ""
Sfidk3  = ""
Sfidk4  = ""

Sglcmp = ""
Sglcnt = ""
Sglprf = ""
Sgllnk = ""



'xxx=armaglos()

If Not (Session("APLBGD") & "") <> "" Then
   Conset(session("CONSET"))
   ofv.wstyles = CorStyler()
End If


ofv.CMBFiller   = "----- Sin Datos -----"
ofv.CMBFillOPC  = "1"



'call ofv.Browser2(Request)

 ofv.Browser()
bc = Request.servervariables("HTTP_USER_AGENT")

bcn = instr(bc,"; MSIE")
If bcn > 0 Then
  xexplor = mid(bc,bcn + 7,3)
Else
  xexplor = "4"
End If



    If xexplor > "5.0" Then
    If xexplor > "5.5" Then
      Session("FULLGUI") = "S"
   Else
      Session("FULLGUI") = "M"
   End If
    Else
   Session("FULLGUI") = "N"
End If

    If Session("USRES") = "B" Then Session("FULLGUI") = "N"

  If session("SMPLMENU") = "S" Then Session("FULLGUI") = "N"


  '''''''''''''''''''''''''''''' Session("FULLGUI") = "S"

If Session("FULLGUI") = "S" Then

    ofv.TitCls    = "ttr"
    ofv.ModTitCls = "1"
    ofv.ModBot    = "2"
    ofv.FBFcnt    = 0
  ''  XFBClas       = "btsb"
    XFBClas       = "bietx"

Else

    ofv.TitCls    = "ttr"
    ofv.ModTitCls = "2"
    ofv.ModBot    = "2"
    ofv.FBFcnt    = 0
  ''  XFBClas       = "btsb"
    XFBClas       = "bietx"

End If

    ofv.TitCls    = "ttx"
    ofv.ModTitCls = "2"


''''''' U - Botonera Arriba
''''''' D - Botonera Abajo

CorBarra          = "U"

ofv.wstyles       = CorStyler()
ofv.CorLogo       = CorLogo()
ofv.wbody         = CStr("bgcolor=" & Session("APLBGD") & " TOPMARGIN=1 LEFTMARGIN=5  ")

ofv.CorFormHeader = CorFormHeader2(session("FormN"),"S")

ofv.CorSetMHeader = CorSetMHeader()

ofv.CorSetFHeader = CorSetFHeader()

xxxx              = CorGenBoton(ofv.ModBot)
xxxx              = CorGenTit(ofv.ModTitCls)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Funciones de administracion de Windows
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Function CorGenTit(CModBot)

        bc = Request.servervariables("HTTP_USER_AGENT")

        bcn = InStr(bc, "; MSIE")
        If bcn > 0 Then

            xexplor = Mid(bc, bcn + 7, 1)

        Else

            xexplor = "4"

        End If

        If xexplor > "5" Then
            xalign = "center"
        Else
            xalign = "left"
        End If

        xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:4;position:RELATIVE;width:50px;height:20px;' >"

        '''''''' 0. Tit.Columnas del Panel

        XinD = 0

        Xinput = ""
        Xinput2 = "Titulo"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"

        '''  Xinput = Xinput & xfstilo


        Xinput = Xinput & " <v:line from='0,0' to='1,0' id='XXXIDXXX' style='z-index:9;align:left;position:relative;top:-1;width:10px;height:10px;' > "
        Xinput = Xinput & "   <v:fill on='True' color='" & session("CMFPNTIT") & "' color2='" & session("CMFPNTIT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
        Xinput = Xinput & "   <v:stroke on='False' /> "
        Xinput = Xinput & "   <v:shadow on='True' color='#ffffff' weight='4px' /> "
        Xinput = Xinput & "   <v:path textpathok='True' /> "
        'response.write Session("CMFPNTITFTSZ")
        Xinput = Xinput & "     <v:textpath on='True' string='XXXDATOXXX' style=""font-size:'" & Session("CMFPNTITFTSZ") + 1 & "pt';font-family:'" & Session("CMFPNTITFTFC") & "';font-weight:bold;""  /> "
        '    Xinput = Xinput & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
        Xinput = Xinput & " </v:line>"


        xxxx = OFV.TitBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))



        '''''''' 1. Tit.Columnas del Panel

        XinD = 1

        Xinput = ""
        Xinput2 = "Titulo"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"

        '''  Xinput = Xinput & xfstilo

        Xinput = Xinput & "  <v:roundrect arcsize='25%'  id='XXXIDXXX'  style='z-index:9;margin-top:0;position:relative;left:0px;top:0px;width:XXXWIDTHXXX;height:20px;'  >"
        Xinput = Xinput & "	<v:fill  type='gradientradial' method='sigma' color='" & session("CMFPNTIT") & "' color2='#eeeeee' colors='95% " & session("CMFPNTIT") & ", 0% #eeeeee' "
        Xinput = Xinput & "      opacity='100%' angle='45'  focusposition=""95%, 15%""  />"
        Xinput = Xinput & "       <v:stroke on='True' weight='2px' color='#eeeeee' />	"
        Xinput = Xinput & "       <v:shadow on='False' color='#dddddd' weight='3px' offset='3px, 2px' />"
        Xinput = Xinput & "        <v:textbox inset='10px,1px,0px,0px' style='font-size:" & Session("CMFPNTITFTSZ") & "pt;font-family:" & Session("CMFPNTITFTFC") & ";font-weight:900;color:" & Session("CMFPNTITFT") & ";text-transform:uppercase;' >XXXDATOXXX</v:textbox>"
        '  Xinput = Xinput & "        <v:extrusion on='True' metal='True'  type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
        Xinput = Xinput & "    </v:roundrect>"




        xxxx = OFV.TitBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 5. Menus

        XinD = 5

        Xinput = ""
        Xinput2 = "Menus"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"




        Xinput = Xinput & " <v:line from='0,0' to='1,0' id='XXXIDXXX' style='z-index:9;align:left;position:relative;top:-1;width:10px;height:10px;' > "
        Xinput = Xinput & "   <v:fill on='True' color='" & Session("CMFMNBOT") & "' color2='" & Session("CMFMNBOT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
        Xinput = Xinput & "   <v:stroke on='False' /> "
        Xinput = Xinput & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
        Xinput = Xinput & "   <v:path textpathok='True' /> "
        Xinput = Xinput & "     <v:textpath on='True' string='XXXDATOXXX' style=""font-size:'" & Session("CMFMNFTSZ") + 1 & "pt';font-family:'" & Session("CMFMNFTFC") & "';font-weight:bold;""  /> "
        Xinput = Xinput & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
        Xinput = Xinput & " </v:line>"


        xxxx = OFV.TitBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))





        '''''''' 6. Menus

        XinD = 6

        Xinput = ""
        Xinput2 = "Menus"
        Xinput3 = Session("CMFPNTITFTFC")
        Xinput4 = ofv.titcls
        Xinput5 = "Encabezados"
        Xinput6 = "no"
        Xinput7 = "Encabezados"



        Xinput = Xinput & "  <v:roundrect arcsize='75%'  id='XXXIDXXX'  style='z-index:9;margin-top:0;position:relative;left:0px;top:0px;width:XXXWIDTHXXX;height:20px;'  >"
        Xinput = Xinput & "	<v:fill  type='gradientradial' method='sigma' color='" & Session("CMFMNBOT") & "' color2='#eeeeee' colors='95% " & Session("CMFMNBOT") & ", 0% #eeeeee' "
        Xinput = Xinput & "      opacity='100%' angle='45'  focusposition=""95%, 15%""  />"
        Xinput = Xinput & "       <v:stroke on='True' weight='1px' color='#eeeeee' />	"
        Xinput = Xinput & "       <v:shadow on='True' color='#eeeeee' weight='2px'  />"
        Xinput = Xinput & "        <v:textbox inset='10px,1px,0px,0px' style='font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";font-weight:900;color:" & Session("CMFMNFT") & ";text-transform:uppercase;' >XXXDATOXXX</v:textbox>"
        '  Xinput = Xinput & "        <v:extrusion on='True' metal='True' type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
        Xinput = Xinput & "    </v:roundrect>"




        xxxx = OFV.TitBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))




    End Function


    '''''''''
    Function CorGenBoton(CModBot)

        bc = Request.servervariables("HTTP_USER_AGENT")

        bcn = InStr(bc, "; MSIE")
        If bcn > 0 Then

            xexplor = Mid(bc, bcn + 7, 1)

        Else

            xexplor = "4"

        End If

        If xexplor > "5" Then
            xalign = "center"
        Else
            xalign = "left"
        End If

        ''color='#aaCCEE' color2='#225588'

        xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:10;position:RELATIVE;width:45px;height:20px;' >"
        '' xfsextr = "  <v:extrusion on='True' metal='True' type='perspective' backdepth='6'  rotationcenter='0,0,0'  rotationangle = '30, 0'/>"
        '' xfsshdw = "  <v:shadow on='False' />"
        xfsshdw = "  <v:shadow on='True' color='#eeeeee' weight='1px' /><v:stroke on='True' color='#eeeeee' weight='1px' />"
        xfsfill = "  <v:fill on='true' color='XXXCOL1XXX' color2='XXXCOL2XXX' type='gradient' angle='30' opacity='100%'  />"
        '' xfsfill = " <v:fill  type='gradientradial' method='sigma' color='XXXCOL2XXX' color2='XXXCOL1XXX' colors='95% XXXCOL2XXX, 0% XXXCOL1XXX' opacity='100%' angle='45'  focusposition=""95%, 15%""  />"

        XSCOLOR2 = Session("CMFPNBOT")
        XSCOLOR1 = Session("CMFMENU")

        xfsfillD = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
        xfsfillD = Replace(xfsfillD, "XXXCOL2XXX", XSCOLOR2)

        Xtypm = "D"
        Xweig = "600"


        '''''''' 0. Bot.Help

        XinD = 0

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "s"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Ayuda On Line"
        Xinput6 = "_self"
        Xinput7 = "Ayuda"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        '''  Xinput = Xinput & "     <v:fill on='true' color='#aaCCEE' color2='#225588' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"


        xxxx = OFV.AddBot(CLng(XinD), "", CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 1. Bot.ANTERIOR

        XinD = 1

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "3"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Pagina Anterior"
        Xinput6 = "_self"
        Xinput7 = "Anterior"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''   Xinput = Xinput & "     <v:fill on='true' color='#aaCCEE' color2='#225588' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"


        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 2. Bot.Siguiente

        XinD = 2

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "4"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Pagina Siguiente"
        Xinput6 = "_self"
        Xinput7 = "Siguiente"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''  Xinput = Xinput & "     <v:fill on='true' color='#aaCCEE' color2='#225588' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"




        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))



        '''''''' 3. Bot.Agregar

        XinD = 3

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "5"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Agregar Elemento"
        Xinput6 = "_self"
        Xinput7 = "Agregar"
        '''''''Xinput7 = "Accion"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#ff0000"
        XSCOLOR2 = "#770000"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''  Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 4. Bot.Buscar

        XinD = 4

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "L"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Buscar Elemento"
        Xinput6 = "_self"
        Xinput7 = "Buscar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''     Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:12pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 5. Bot.Consultar

        XinD = 5

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "1"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Consultar - Lista"
        Xinput6 = "_new"
        Xinput7 = "Consultar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#ffbb22"
        XSCOLOR2 = "#bbaa22"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        '' Xinput = Xinput & "     <v:fill on='true' color='#ffbb22' color2='#bbaa22' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 6. Bot.Test

        XinD = 6

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "J"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Test"
        Xinput6 = "_new"
        Xinput7 = "Test"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''   Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:13pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))

        '''''''' 7. Bot.Volver

        XinD = 7

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "&#0231;"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Volver"
        Xinput6 = "_self"
        Xinput7 = "Volver"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#005500"
        XSCOLOR2 = "#00cc00"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''   Xinput = Xinput & "     <v:fill on='true' color='#005500' color2='#00cc00' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:12pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))

        '''''''' 8. Bot.Eliminar

        XinD = 8

        Xtyp = "C"
        ''  Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "4"
        If BrowserName() = "IE" Or BrowserName() = "CHROME" Then
            Xinput3 = "Corflaticon"
        Else
            Xinput3 = "Corflaticon"
        End If
        Xinput4 = "bte"
        Xinput5 = "Eliminar"
        Xinput6 = "_self"

        If ofv.modbot = "2" Then

            Xinput7 = "&#xf2ed;"

        Else
            Xinput7 = "Eliminar"
        End If

        ''' Xinput = Xinput & xfstilo
        Xinput = Xinput & "  <v:roundrect arcsize='65%' id='AP'  style='z-index:4;position:RELATIVE;width:12px;height:12px;' >"
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#ff0000"
        XSCOLOR2 = "#770000"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-4px,0px,0px' style=""text-align:" & xalign & ";font-size:7pt;font-family:'XXXFONTXXX';font-weight:400;color:#FFFFFF;"" >XXXDATOXXX"
        Xinput = Xinput & "     </v:textbox>"
        Xinput = Xinput & xfsshdw
        ' Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 11. Bot.ANTERIOR gris

        XinD = 11

        Xtyp = "C"
        ''  Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "3"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas & "G"
        Xinput5 = "Pagina Anterior no Disponible"
        Xinput6 = "_self"
        Xinput7 = "Anterior"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#eeeeee"
        XSCOLOR2 = "#aaaaaa"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''  Xinput = Xinput & "     <v:fill on='true' color='#eeeeee' color2='#aaaaaa' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))



        '''''''' 12. Bot.Siguiente gris

        XinD = 12

        Xtyp = "C"
        ''  Xtyp   = Xtypm

        Xinput = ""
        Xinput2 = "4"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas & "G"
        Xinput5 = "Pagina Siguiente no Disponible"
        Xinput6 = "_self"
        Xinput7 = "Siguiente"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#eeeeee"
        XSCOLOR2 = "#aaaaaa"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''     Xinput = Xinput & "     <v:fill on='true' color='#eeeeee' color2='#aaaaaa' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))

        '''''''' 13. Bot.Aceptar

        XinD = 13

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "a"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Aceptar Operacion"
        Xinput6 = "_self"
        Xinput7 = "Aceptar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''   Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))

        '''''''' 14. Bot.Cancelar

        XinD = 14

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "r"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Cancelar Operacion"
        Xinput6 = "_self"
        Xinput7 = "Cancelar"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#ff0000"
        XSCOLOR2 = "#770000"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''   Xinput = Xinput & "     <v:fill on='true' color='#ff0000' color2='#770000' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))

        '''''''' 15. Bot.Confirmar

        XinD = 15

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "a"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Confirmar"
        Xinput6 = "_self"
        Xinput7 = "Confirmar Operacion"

        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#0000ff"
        XSCOLOR2 = "#336699"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''   Xinput = Xinput & "     <v:fill on='true' color='#0000ff' color2='#336699' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-1px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 16. Bot.Esp1

        XinD = 16

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "2"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Consultar"
        Xinput6 = "_new"
        Xinput7 = "Consultar"



        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#f7aa22"
        XSCOLOR2 = "#b79911"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''    Xinput = Xinput & "     <v:fill on='true' color='#f7aa22' color2='#b79911' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))

        '''''''' 17. Bot.Indice

        XinD = 17

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "&#0105;"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Abrir/Cerrar indice"
        Xinput6 = "_self"
        Xinput7 = "Indice"



        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#f7aa22"
        XSCOLOR2 = "#b79911"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''    Xinput = Xinput & "     <v:fill on='true' color='#f7aa22' color2='#b79911' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        ' xxxx = OFV.AddBot(clng(XinD), cstr(Xinput), cstr(Xinput2), cstr(Xinput3), cstr(Xinput4), cstr(Xinput5), cstr(Xinput6), cstr(Xinput7))

        '''''''' 18. Bot.Glosario

        XinD = 18

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "&#062;"
        Xinput3 = "corflaticon"
        Xinput4 = XFBClas
        Xinput5 = "Glosario"
        Xinput6 = "_self"
        Xinput7 = "Glosario"



        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#f7aa22"
        XSCOLOR2 = "#b79911"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If

        ''    Xinput = Xinput & "     <v:fill on='true' color='#f7aa22' color2='#b79911' type='gradient' angle='30' opacity='80%'  />"
        Xinput = Xinput & "	<v:textbox inset='0px,-2px,0px,0px' style=""text-align:" & xalign & ";font-size:10pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"



        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


        '''''''' 20. Bot.INICIO

        XinD = 20

        ''  Xtyp   = "C"
        Xtyp = Xtypm

        Xinput = ""
        Xinput2 = "INGRESAR"
        Xinput3 = "verdana"
        Xinput4 = XFBClas
        Xinput5 = "Ingresar"
        Xinput6 = "_self"
        Xinput7 = "Ingresar"

        xfstilo = "  <v:roundrect arcsize='75%' id='AP'  style='z-index:10;position:RELATIVE;width:90px;height:25px;' >"
        Xinput = Xinput & xfstilo
        Xinput = Xinput & "     <v:stroke on='FALSE' color='#eeeeee' weight='1px' />"

        XSCOLOR1 = "#aaCCEE"
        XSCOLOR2 = "#225588"

        If Xtyp = "C" Then
            xfsfillC = Replace(xfsfill, "XXXCOL1XXX", XSCOLOR1)
            xfsfillC = Replace(xfsfillC, "XXXCOL2XXX", XSCOLOR2)
            Xinput = Xinput & xfsfillC
        Else
            Xinput = Xinput & xfsfillD
        End If


        Xinput = Xinput & "	<v:textbox inset='0px,2px,0px,0px' style=""text-align:" & xalign & ";font-size:7pt;font-family:'XXXFONTXXX';font-weight:" & Xweig & ";color:#FFFFFF;" & XspDShw("1", "1", session("CMFPNTIT")) & ";"" >XXXDATOXXX</v:textbox>"
        Xinput = Xinput & xfsshdw
        Xinput = Xinput & xfsextr
        Xinput = Xinput & "  </v:roundrect>"

        xxxx = OFV.AddBot(CLng(XinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))
'RESPONSE.WRITE  CStr(Xinput) & "22222222222"
    End Function

    '''''''''
    Function genbotoncambios(xsaccion)

        '   XSTITULOS="Cambios" '& session("APLMDCLI")
        'if  session("APLMDCLI")="BGAL" THEN
        XSTITULOS = "Resumen de cambios" '& len(cstr(xsaccion &""))

        xxxx = CorGenBotonXp(3, "NO", "&#0113;", "NO", "NO", "Lista de cambios", "NO", XSTITULOS)

        If Len(CStr(xsaccion & "")) = 0 Then
            genbotoncambios = ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

        Else
            genbotoncambios = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", xsaccion & "")
        End If



    End Function






    Function CorGenBotonXp(XCPinD, XCPinput, XCPinput2, XCPinput3, XCPinput4, XCPinput5, XCPinput6, XCPinput7)


        xxxx = OFV.GetBot(CLng(XCPinD))


        '    Xinput = ""
        '    Xinput2 = "2"               ' caracter
        '    Xinput3 = "corflaticon"        ' font
        '    Xinput4 = XFBClas
        '    Xinput5 = "Consultar"       ' hint
        '    Xinput6 = "_blank"          ' target
        '    Xinput7 = "Consultar"       ' nombre


        Xinput = OFV.FBotDat1
        Xinput2 = OFV.FBotDat2
        Xinput3 = OFV.FBotDat3
        Xinput4 = OFV.FBotDat4
        Xinput5 = OFV.FBotDat5
        Xinput6 = OFV.FBotDat6
        Xinput7 = OFV.FBotDat7


        If XCPinput <> "NO" Then Xinput = XCPinput
        If XCPinput2 <> "NO" Then Xinput2 = XCPinput2
        If XCPinput3 <> "NO" Then Xinput3 = XCPinput3
        If XCPinput4 <> "NO" Then Xinput4 = XCPinput4

        If XCPinput5 <> "NO" Then Xinput5 = XCPinput5


        If XCPinput6 <> "NO" Then Xinput6 = XCPinput6

        If XCPinput7 <> "NO" Then Xinput7 = XCPinput7



        xxxx = OFV.AddBot(CLng(XCPinD), CStr(Xinput), CStr(Xinput2), CStr(Xinput3), CStr(Xinput4), CStr(Xinput5), CStr(Xinput6), CStr(Xinput7))


    End Function


    ''''''''''''

    Function GenTitulo(xfDato, xfModo)



        xGenCont = xfDato

        If ofv.ModTitCls = "1" Or ofv.ModTitCls = "5" Or ofv.ModTitCls = "0" Then

            If xfModo = 1 Then

                xfInd = CLng(ofv.ModTitCls)
            Else

                xfInd = 0
            End If

            xGenConta = ofv.GetTit(xfInd + 5)
            xGenConta = ofv.FTitDat1

            xGenCont = xGenConta

            ''    xGenCont = FTitTab(CLng(ModTitCls), 0)

            If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then

                xGenWidth = CStr(Len(xfDato) * 10) & "px"
                xGenCont = Replace(xGenCont, "XXXDATOXXX", xfDato)
                xGenCont = Replace(xGenCont, "XXXWIDTHXXX", xGenWidth)
                ''  xGenCont = Replace(xGenCont, "XXXIDXXX", "r" & xGenNom)

            End If

        End If

        GenTitulo = xGenCont

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''' FUNCIONES DE MASTER-PANELES / PANELES DE MENUS / PANELES DE DOCUMENTOS Y COMPONENTES ''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Function CorPnlFnc(XCopc, XCdsc)

        Dim Whtml
        Dim Whtml1

        Whtml = ""
        Whtml1 = ""


        '''''''''''''''''''''''''

        Whtml = Whtml & "<style ID='oStyle' > span.gls {cursor:pointer;background-color:#bbbbbb;color:white!important;"
        Whtml = Whtml & "border-style:outset;border-width:1;border-color:#888888;line-height: 25px;font-weight:bold;} </style>"



        Whtml = Whtml & "<public>" & Chr(13)
        Whtml = Whtml & "<PROPERTY name='myWin' />"
        Whtml = Whtml & "<method name='creawin' />"
        Whtml = Whtml & " </public>" & Chr(13)

        Whtml = Whtml & "<SCRIPT language='javascript' for='mt2' event='onmouseout' >" & Chr(13) & Chr(13)
        Whtml = Whtml & " if (document.selection.type=='Text') {;"
        Whtml = Whtml & "     myGlos = document.selection.createRange().text;"
        Whtml = Whtml & " };"
        Whtml = Whtml & "</SCRIPT>" & Chr(13) & Chr(13)

        Whtml = Whtml & "<SCRIPT language=javascript >             " & Chr(13) & Chr(13)



        Whtml = Whtml & "    var myWin;                            " & Chr(13)
        Whtml = Whtml & "    var myGlos;                           " & Chr(13)
        Whtml = Whtml & "      var Aglos = new Array();     " & Chr(13)
        Whtml = Whtml & " oStyle.disabled = true;                  " & Chr(13)
        Whtml = Whtml & " function creawin(spmk) {                 " & Chr(13)
        ' Whtml = Whtml & " document.all['tmtshwx'].style.visibility='hidden'; " & chr(13)
        '''' Whtml = Whtml & " document.all['DCEX1'].style.visibility='hidden'; " & chr(13)

        Whtml = Whtml & "document.getElementById(spmk).scrollIntoView(true);  " & Chr(13)
        Whtml = Whtml & " document.getElementById('THM').scrollIntoView(true); "
        ' Whtml = Whtml & " document.all['tmtshwx'].style.visibility='visible'; " & chr(13)
        Whtml = Whtml & " clIndice();                              " & Chr(13)
        Whtml = Whtml & " }                                        " & Chr(13)

        Whtml = Whtml & "function AbrirCons(Datos,nom) {" & Chr(13)
        Whtml = Whtml & " var w=570;h=790 ;"
        ' if  session("APLMDCLI")="BC" THEN
        Whtml = Whtml & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
        ' END IF
        Whtml = Whtml & " myWin= open(Datos"
        Whtml = Whtml & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom," '& Chr(34)
        'Whtml = Whtml & "status=yes,menubar=no,toolbar=no,height=570,width=790,resizable=yes" & Chr(34)
        Whtml = Whtml & "  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'" '& CHR(34)
        Whtml = Whtml & " ); " & Chr(13) & " } " & Chr(13)

        ''''''
        Whtml = Whtml & "function setglos()  " & Chr(13)
        Whtml = Whtml & "{                   " & Chr(13)
        Whtml = Whtml & "   myGlos = document.selection.createRange().text;" & Chr(13)
        Whtml = Whtml & "}                   " & Chr(13)

        Whtml = Whtml & "function verglos()" & Chr(13)
        Whtml = Whtml & "{" & Chr(13)
        Whtml = Whtml & " var xtxt = myGlos;"
        Whtml = Whtml & " var sasp = 'GLOSCONS.ASP?DATOS=" & Chr(34) & "' + xtxt + '" & Chr(34) & "';"
        ' Whtml = Whtml & " alert(xtxt);"
        If BrowserName() = "IE" Then
            Whtml = Whtml & " window.showModalDialog(sasp,window,"
            Whtml = Whtml & Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
            Whtml = Whtml & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
            Whtml = Whtml & " ); " & Chr(13)
        Else
            Whtml = Whtml & " window.open(sasp,window,"
            Whtml = Whtml & Chr(34) & "Width=600,Height=350,"
            Whtml = Whtml & "Top=250,Left=300,status=no" & Chr(34)
            Whtml = Whtml & " ); " & Chr(13)
        End If
        Whtml = Whtml & " myGlos = '';"
        Whtml = Whtml & "}" & Chr(13)

        Whtml = Whtml & " function versup() { "
        Whtml = Whtml & xxx
// whtml=whtml & "var zzzz=document.getElementById('mt2') [0];"
 Whtml = Whtml & " if (oStyle.disabled == true) {"

        Whtml = Whtml & " oStyle.disabled = false;var zzz='';"
        Whtml = Whtml & armaglos()
 //Whtml = Whtml & " for (var i=0;i<Aglos.length;i++){ "
/// Whtml = Whtml & "	zzz=zzz +  Aglos[i] + ' ';}"
 Whtml = Whtml & "localSearchHighlight2();}"
 //Whtml = Whtml & "}"
 Whtml = Whtml & " else {"
        Whtml = Whtml & " oStyle.disabled = true;"
        Whtml = Whtml & " unhighlight(document.getElementsByTagName('body')[0]);} }"


        Whtml = Whtml & " function verglos2(xdata)" & Chr(13)
        Whtml = Whtml & " {" & Chr(13)
        Whtml = Whtml & " var xtxt = xdata;"
        Whtml = Whtml & " var sasp = 'GLOSCONS.ASP?DATOS=" & Chr(34) & "' + xtxt + '" & Chr(34) & "';"
        Whtml = Whtml & " window.showModalDialog(sasp,window,"
        Whtml = Whtml & Chr(34) & "dialogWidth:600px;dialogHeight:350px;help:no;resizable:yes;"
        Whtml = Whtml & "dialogTop:250px;dialogLeft:300px;status:no" & Chr(34)
        Whtml = Whtml & " ); " & Chr(13)
        Whtml = Whtml & "}" & Chr(13)


        Whtml = Whtml & "function clIndice() {" & Chr(13)
        Whtml = Whtml & "  if  (document.getElementById('DHEX1').style.display=='none')   " & Chr(13)
        Whtml = Whtml & "      { document.getElementById('DHEX1').style.display='';} " & Chr(13)
        Whtml = Whtml & "  else                                             " & Chr(13)
        Whtml = Whtml & "      { document.getElementById('DHEX1').style.display='none';}  " & Chr(13)
        Whtml = Whtml & " } " & Chr(13)


        Whtml = Whtml & "function ApreIndice() { " & Chr(13)

        Whtml = Whtml & "document.getElementById('tmtshwxy').style.display  ='';"
        Whtml = Whtml & "document.getElementById('mxbtopynx').style.display ='none';"
        Whtml = Whtml & "document.getElementById('mxbtclynx').style.display ='';"
        Whtml = Whtml & " } " & Chr(13)

        Whtml = Whtml & "function ClosIndice() { " & Chr(13)
        Whtml = Whtml & "document.getElementById('tmtshwxy').style.display  ='none';"
        Whtml = Whtml & "document.getElementById('mxbtopynx').style.display ='none';"
        Whtml = Whtml & "document.getElementById('mxbtclynx').style.display ='none';"
        Whtml = Whtml & " } " & Chr(13)

        Whtml = Whtml & "function DCEXRCLC() {"
        Whtml = Whtml & " var xrclc0 = document.body.clientWidth;"
        Whtml = Whtml & " var xrclc1 = document.body.clientHeight;"
        Whtml = Whtml & " var xrclc2 = document.getElementById('TCEX1').clientHeight;"
        Whtml = Whtml & " var xrclc3 = document.getElementById('TCEX1').offsetTop;"
        '' Whtml = Whtml & " document.all.mt2.style.height = xrclc1 - (xrclc3 + xrclc2) - 5;"
        '' Whtml = Whtml & " document.all.mt2.style.width  = xrclc0 - 10;"
        Whtml = Whtml & "}"

        Whtml = Whtml & "</SCRIPT>" & Chr(13) & Chr(13)


        styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee;"
        styl = styl & " SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa;"
        styl = styl & " SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"

        If BrowserName() = "IE" Then
            sglos2 = "<b><font face=corflaticon size=3 color=#ffffff>&</font></b>"
            sglos1 = "<b><font face=corflaticon size=3 color=#ffffff>q</font></b>"
            sglos3 = "<b><font face=corflaticon size=3 color=#ffffff>n</font></b>"

            '  sfidkA  = "<font face=corflaticon size=1 color=#99ccee >&#060;</font>"

            sfidkA = "<font face=corflaticon size=2 color=#99ccee >&#0157;</font>"
            sfidkC = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
            sfidkB = "<font face=corflaticon size=2 color=#6699cc >&#0159;</font>"

            sfidk0 = "<font face=corflaticon size=3 color=" & session("CMFMNBOT") & " >&#0157;</font>"
            sfidk1 = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
            sfidk2 = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
            sfidk3 = "<font face=corflaticon size=2 color=#00aa00 >4</font>"
            sfidk4 = "<font face=corflaticon size=2 color=#00aa00 >&#060;</font>"

            sglcmp = "<div class=btba style='xwidth:1px;clear:right;' ><font class=btic >&#0157;</font></div>"
            sglcnt = "<div class=btbs style='xwidth:1px;clear:right;' ><font class=btic >G</font></div>"
            sglprf = "<div class=btbn style='xwidth:1px;clear:right;' ><font class=btic >&#0128;</font></div>"
            sgllnk = "<div class=btbn style='xwidth:1px;clear:right;' ><font class=btic >&#0126;</font></div>"
        Else

            sglos2 = "<b><font face=corflaticon size=3 color=#ffffff>&</font></b>"
            sglos1 = "<b><font face=corflaticon size=3 color=#ffffff>q</font></b>"
            sglos3 = "<b><font face=corflaticon size=3 color=#ffffff>n</font></b>"

            '  sfidkA  = "<font face=corflaticon size=1 color=#99ccee >&#060;</font>"

            sfidkA = "<font face=corflaticon size=2 color=#99ccee >&#0157;</font>"
            sfidkC = "<i class='las la-angle-right' style=color:#132574 >4</i>"
            sfidkB = "<font face=corflaticon size=2 color=#6699cc >&#0159;</font>"

            sfidk0 = "<font face=corflaticon size=3 color=" & session("CMFMNBOT") & " >&#0157;</font>"
            sfidk1 = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
            sfidk2 = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
            sfidk3 = "<font face=corflaticon size=2 color=#00aa00 >4</font>"
            sfidk4 = "<font face=corflaticon size=2 color=#00aa00 >&#060;</font>"

            sglcmp = "<div class=btba style='xwidth:1px;clear:right;' ><i class='las la-file-alt'></i></div>"
            sglcnt = "<div class=btbs style='xwidth:1px;clear:right;' ><font class=btic >G</font></div>"
            sglprf = "<div class=btbn style='xwidth:1px;clear:right;' ><font class=btic >&#0128;</font></div>"
            sgllnk = "<div class=btbn style='xwidth:1px;clear:right;' ><font class=btic >&#0126;</font></div>"

        End If


        ''''''''''''''''''''''''
        CorPnlFnc = Whtml

    End Function

    Function CorPnlIndx(XCopc, XCdsc)

        Dim Whtml
        Dim Whtml1

        Whtml = ""
        Whtml1 = ""

        StyL = " STYLE='cursor:pointer;font-weight:normal;padding-left:7px;padding-right:7px;line-height:20px;"
        StyL = StyL & "border-style:solid;border-width:1;border-color:#000000;background-color:#ff99cc;"
        StyL = StyL & "' "


        Whtml1 = ofv.GenCelda("", "vt", "center", "", "", "", "", "<br><big><big>Cargando DOCUMENTO...</big><br><br>un momento por favor</big><br>", "si")
        Whtml1 = ofv.GenRow("", "", "", "", "", "", "", Whtml1, "si")

        Whtml = Whtml & "<table id='btnreco'  style='margin-left:16px;padding-bottom:5px' ><tr>"
        Whtml = Whtml & "<td id='mxbtopynx' style='display:none'  class='bietx' ><a href='javascript:void(0);' onclick='ApreIndice();' class='' >.</a></td>"

        wwwkadisplay = "none"

        ' if session("WEBDESK") = "S" and session("ModMain") = "CL"  then
        If session("WEBDESK") = "S" Then
            wwwkadisplay = "block"
        End If

        Whtml = Whtml & "<td id='mxbtclynx' style='display:none' class='bietx' ><a href='javascript:void(0);' xonclick='ClosIndice();' class='' ><< Ocultar indice.</a></td>" '& session("ModMain")


       '  editarcomp session("asp") ="CompTree.asp"  si request.querystring("tipocomp")="" destruye esa variable ya que viene de edicion de documentos
        ' esto hay que mejorar
     '   response.write "2222222222222222222" & session("asp")
        if (request.querystring("tipocomp") ="" and instr(ucase(Session("asp")),"NOVED") =0) and request.querystring("family")<>"K" then session("asp")=""

    '   response.write request.querystring & " " & session("asp")
        If session("WEBDESK") = "S" And session("ModMain") = "CL" Then

            If Trim(nCodigo & "") <> ""   Then

             if  UCase(session("asp")) <> UCase("CompTree.asp") then
                xverasp = "CoObjRecomendar.asp?Tipobjeto=" & sTipo & "&Codobjeto="
                xverasp = xverasp & nCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
                xverasp = xverasp & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2 & "&NIV3=" & sNIV3 & "&NIV4=" & sNIV4 & "&NIV5=" & sNIV5
                Whtml = Whtml & "<td id='mxbtreco' style='udisplay:block'  class='bietx'  ><a href='" & xverasp & " target='Principal'  class='' ><i class='las la-hand-point-right la-1x'></i>&nbsp;Recomendar el documento.</a></td>"
              elseif  UCase(session("asp")) = UCase("CompTree.asp")  then

                 xverasp = "GenerarComp.asp?id=" & nCodigo
               Whtml = Whtml & "<td id='mxbtreco' style='udisplay:block'  class='bietx'  ><a href='" & xverasp & "' target='Principal'  class='' >Imprimir.</a></td>"
              end if
            End If
           'response.write session("WEBDESK")  & session("ModMain")

            If (InStr(UCase(session("asp")), UCase("novedmenu.asp")) > 0 Or InStr(UCase(session("asp")), UCase("ObjTreeSubTemaCons.asp")) > 0 Or InStr(UCase(session("asp")), UCase("SUBCONSULTA.ASP")) > 0 Or InStr(UCase(session("asp")), UCase("Objetos.asp")) > 0) Or UCase(session("asp")) <> UCase("CompTree.asp") Then
                'session("asp") = ucase("SUBCONSULTA.ASP") or session("asp") = ucase("Objetos.asp") THEN

                If Trim(nCodigo & "") <> "" Then
                    xverasp = "CoObjsuscripcion.asp?Tipobjeto=" & sTipo & "&Codobjeto="
                    xverasp = xverasp & nCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
                    xverasp = xverasp & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2 & "&NIV3=" & sNIV3 & "&NIV4=" & sNIV4 & "&NIV5=" & sNIV5



                    Whtml = Whtml & "<td id='mxbtreco' style='udisplay:block'  class='bietx'  ><a href='" & xverasp & " target='Principal'  class='' ><i class='las la-bell la-1x'></i>&nbsp;Suscripción</a></td>" '& session("ASP")
                End If

                If Trim(nCodigo & "") <> "" Then
                    xverasp = "CoObjfavoritos.asp?Tipobjeto=" & sTipo & "&Codobjeto="
                    xverasp = xverasp & nCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
                    xverasp = xverasp & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2 & "&NIV3=" & sNIV3 & "&NIV4=" & sNIV4 & "&NIV5=" & sNIV5



                    Whtml = Whtml & "<td id='mxbtreco' style='udisplay:block'  class='bietx'  ><a href='" & xverasp & " target='Principal'  class='' ><i class='lar la-star la-1x'></i>&nbsp;Favoritos</a></td>" '& session("ModMain")
                End If
            End If

            If SESSION("COMENTARIOS") = "S" Then
                If Trim(nCodigo & "") <> "" and UCase(session("asp")) <> UCase("CompTree.asp") Then
                    xx = GTCPTOFUN()

                    ' sAccion = "gencomentario.asp?wasotxt=N&Codobjeto=" & nCodigo
                    '  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nCodigo '  & "&ID2=" & rs(0)  & "&ID3=" & rs(0)

                    sAccion = "gencomentario.asp?wasotxt=N&Codobjeto=" & nCodigo
                    sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL & "&NIV1=" & sNIV1 & "&NIV2=" & sNIV2 & "&NIV3=" & sNIV3 & "&NIV4=" & sNIV4 & "&NIV5=" & sNIV5

                    xsffun = " onclick=""document.getElementById('DX4').style.display='';"" "
                    sHref = sAccion
                    sInput = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
                    '  Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","center","","1","","",sInput,"si")
                    Celdaobs = ofv.GenLink("", "biet" & "  " & xsffun, sHref, "WK4", "", "", "<i class='las la-comment la-1x'></i>&nbsp;	Comentarios", "si")

                    GTWhtm = GTWhtm & "<SCRIPT language=javascript >" & Chr(13)
                    GTWhtm = GTWhtm & "function GTCOBSCL() {" & Chr(13)
                    GTWhtm = GTWhtm & "   document.getElementById('DX4').style.display='none';" & Chr(13)
                    GTWhtm = GTWhtm & " } " & Chr(13)
                    GTWhtm = GTWhtm & "</SCRIPT>"

                    Whtml = Whtml & "<td id='mxbtreco' style='udisplay:block'  class='bietx'  >" & Celdaobs&"</td>"  & xx & GTWhtm
                 End If
            End If


            If SESSION("WKFOBSBTN") = "S<<" Then


                sAccion = "COOBSHIST.asp?Codobjeto=" & nCodigo
                sAccion = sAccion & "&Tipobjeto=" & sTipo


                sHref = sAccion
                ' Celdaobs  = ofv.GenLink("","bietx" ,sHref,"WK4","","","Observaciones","si")

                Whtml = Whtml & "<td id='mxbtreco' style='xdisplay:block'  class='bietx'  ><a href='" & sAccion & "' target='Principal'  class='bietx' >Observaciones </a></td>" '& session("ModMain")

                sAccion = "COOBSHIST.asp?FILTRO=USR&Codobjeto=" & nCodigo
                sAccion = sAccion & "&Tipobjeto=" & sTipo
                sHref = sAccion
                ' Celdaobs  = ofv.GenLink("","bietx" ,sHref,"WK4","","","Observaciones","si")
                Whtml = Whtml & "<td id='mxbtreco' style='xdisplay:block'  class='bietx'  ><a href='" & sAccion & "' target='Principal'  class='bietx' >Mis Observaciones </a></td>" '& session("ModMain")

            End If


        ElseIf session("WEBDESK") = "S" And session("ModMain") = "ED" Then

            If SESSION("COMENTARIOS") = "S2" Then


                sAccion = "consultarcmd.asp?codigo=" & nCodigo
                sAccion = sAccion & "&tipo=" & sTipo & "&pend=S"

                sHref = sAccion
                Celdaobs = ofv.GenLink("", "" & "  " & xsffun, sHref, "_new", "", "", "Comentarios", "si")
                Whtml = Whtml & "<td id='mxbtreco' style='background-color:red'  class='bietx'  >" & Celdaobs&"</td>"
  ElseIf SESSION("COMENTARIOS") = "S1" Then
                sAccion = "consultarcmd.asp?codigo=" & nCodigo
                sAccion = sAccion & "&tipo=" & sTipo & "&pend=N"
                sHref = sAccion
                Celdaobs = ofv.GenLink("", "" & "  " & xsffun, sHref, "_new", "", "", "Comentarios", "si")
                Whtml = Whtml & "<td id='mxbtreco' style='background-color:green'  class='bietx'  >" & Celdaobs&"</td>"


  End If

        ElseIf session("WEBDESK") = "S" And session("ModMain") = "WK" Then

            If SESSION("WKFOBSBTN") = "S" Or SESSION("WKFOBSBTN") = "S1" Or SESSION("WKFOBSBTN") = "S2" Then


                sAccion = "COOBSHIST.asp?ACT=N&Codobjeto=" & nCodigo
                sAccion = sAccion & "&Tipobjeto=" & sTipo


                sHref = sAccion
                ' Celdaobs  = ofv.GenLink("","bietx" ,sHref,"WK4","","","Observaciones","si")

                Whtml = Whtml & "<td id='mxbtreco' style='xdisplay:block'  class='bietx'  ><a href='" & sAccion & "' target='Principal'  class='bietx' >Observaciones </a></td>" '& SESSION("WKFOBSBTN")
                If SESSION("WKFOBSBTN") = "S1" Then
                    sAccion = "COOBSHIST.asp?ACT=N&FILTRO=USR&Codobjeto=" & nCodigo
                    sAccion = sAccion & "&Tipobjeto=" & sTipo
                    sHref = sAccion
                    Whtml = Whtml & "<td id='mxbtrobs' style='xdisplay:block'  class='bietx'  ><a href='" & sAccion & "' target='Principal'  class='bietx' >Mis Observaciones </a></td>" '& session("ModMain")

                Else

                    sAccion = "COOBSHIST.asp?ACT=N&FILTRO=USR&Codobjeto=" & nCodigo
                    sAccion = sAccion & "&Tipobjeto=" & sTipo
                    sHref = sAccion
                    Whtml = Whtml & "<td id='mxbtrobs' style='display:none'  class='bietx'  ><a href='" & sAccion & "' target='Principal'  class='bietx' >Mis Observaciones </a></td>" '& session("ModMain")
                End If

            End If
        End If
        Whtml = Whtml & "</tr></table>" '& bmodificar & SESSION("x")


        Whtml = Whtml & "<div id='mt1' style='display:block' ><br><br><br><br><center>"
        Whtml = Whtml & ofv.GenTabla("", "", "", "60% style='border-style:SOLID;border-width:2;border-color:#AAAAAA;' ", "", "0", "", "0", "0", "#ffffff", "", Whtml1, "si")
        Whtml = Whtml & "</center></div>"


        Whtml1 = ofv.GenCelda("tmtshwxy", "uti style='display:none;border-right-style:solid;border-right-width:2;border-right-color:#cccccc;' valign='top' ", "", "20%", "", "", "", "&nbsp;INDEX", "si")
        Whtml1 = Whtml1 & ofv.GenCelda("tmtshwxz", "ut xstyle='border-style:solid;border-width:1;border-color:#cccccc;' valign='top'  ", "", "80%", "", "", "", "&nbsp;", "si")
        Whtml1 = ofv.GenRow("", "", "", "", "", "", "", Whtml1, "si")
        'Whtml1 = whtml1 & "<div id=""divtmtshwxz"" style=""height: 600px; overflow: auto;"" name=""divtmtshwxz"">"

        '''border-style:solid;border-width:1;border-color:#000000;

        ' Whtml = Whtml & "<div id='mtshwx' style='display:block;' >"
        'class aaa
        Whtml = Whtml & ofv.GenTabla("tmtshwx", "", "", "100% xstyle='border-top-style:solid;border-top-width:2;border-top-color:#ccaaaa;xdisplay:none;z-index:99;max-width: 1400px;' ", "", "0", "", "0", "0", "", "", Whtml1, "si")
        ' Whtml = Whtml & "</div>"

        'alert('llego aca');
        '  document.getElementsByTagName(""div"")[0]

        Whtml = Whtml & "<script language=javascript >"

        Whtml = Whtml & "document.getElementById('mt1').style.display='none';"

        Whtml = Whtml & " var pp=document.createElement(""div"");var pp2=document.getElementById(""mt2"");"
        'whtml = whtml & "alert(pp2);"
        Whtml = Whtml & " pp.setAttribute(""id"",""divtmtshwxz"");pp.style.height =""600px"";"
        Whtml = Whtml & "  pp.style.overflow=""auto"";"
        'Whtml = Whtml & "   if (pp2.firstChild  != null) {xx=pp2.firstChild;} else{xx=pp2; alert('1');}"

        If BrowserName() <> "IE" And session("ModMain") = "ED" Then
            Whtml = Whtml & "if (pp2.hasChildNodes()) { pp2.removeAttribute('style');"

            Whtml = Whtml & "  var children = pp2.childNodes;"

            Whtml = Whtml & "   for (var i = 0; i <children.length; i++) { "
            Whtml = Whtml & " if (children[i].nodeType ==1) { "
            Whtml = Whtml & "      document.getElementById('tmtshwxz').appendChild(children[i]);}"

            Whtml = Whtml & "} "
            'ale 16/9/2025 duplicado info mt2
			 'Whtml = Whtml & "pp2.removeAttribute('style');var parentelem = document.getElementById('mt2').parentNode;var theFirstChild = parentelem.firstChild;"
            ' Whtml = Whtml & "parentelem.insertBefore(document.getElementById('btnreco'), theFirstChild);"
            Whtml = Whtml & "} "

        Else
            'ale 16/9/2025 duplicado info mt2
             Whtml = Whtml & "pp2.removeAttribute('style');var parentelem = document.getElementById('zglos').parentNode;;var theFirstChild = parentelem.firstChild;"
             Whtml = Whtml & "parentelem.insertBefore(document.getElementById('btnreco'), theFirstChild);"
           ' Whtml = Whtml & " alert('1');    ' document.getElementById('tmtshwxz').innerHTML=pp2.innerHTML;"
        End If

        ' Whtml = Whtml & "document.getElementById('tmtshwxz').innerHTML='<div NAME=divtmtshwxz id=divtmtshwxz style=height:600px;overflow:auto;}>' + pp2.innerHTML  + '</div>';"
        'Whtml = Whtml & "document.getElementById('tmtshwxz').innerHTML=document.getElementById('mt2').innerHTML;"
      '  Whtml = Whtml & " var txtid='' ;var et=document.getElementById('TXHI3M'); if (et  != null) {txtid=document.getElementById('TXHI3M').innerHTML}else {textid=''}"
        Whtml = Whtml & "document.getElementById('tmtshwxy').innerHTML= '<div id=xdestino style=""width:400px;height:600px;overflow:auto;"">' + document.getElementById('TXHI3M').innerHTML + '</div>';"

        'Whtml = Whtml & " alert(document.getElementById('TXHI3M').innerHTML);"


        ' LINEA agregado
        ' Whtml = Whtml & "document.getElementById('tmtshwxz').innerHTML='<div NAME=" & chr(34) & "divtmtshwxz" & chr(34) & " id=" & chr(34) & "divtmtshwxz" & chr(34) & "  style=height:600px;overflow:auto;>' + document.getElementById('mt2').innerHTML  + '</div>';"

        ' Whtml = Whtml & "document.getElementById('tmtshwxy').innerHTML= '<div style=width:200px;height:600px;overflow:auto;>' + document.getElementById('TXHI3M').innerHTML + '</div>';"
        ' LINEA AGREgada  fin'

        '------------------------------------------------------------------
        ' Whtml = Whtml & "document.all.tmtshwxy.style.display='';"

        'Whtml = Whtml & "document.all.TXHI3M.style.display='none';"
        'Whtml = Whtml & "document.all.mt2.style.display='none';"
        '------------------------------------------------------------------

        Whtml = Whtml & "var et=document.getElementById('tmtshwxy'); if (et  != null) { document.getElementById('tmtshwxy').style.display='none';}"

        Whtml = Whtml & " var et=document.getElementById('TXHI3M'); if (et  != null) { document.getElementById('TXHI3M').style.display='none';}"
        Whtml = Whtml & " var et=document.getElementById('mt2');  if (et  != null)  {document.getElementById('mt2').style.display='none';}"
        '''// Obtenemos el elemento
        '''// Obtenemos el padre de dicho elemento
        '''// con la propiedad “parentNode”
        '''// Eliminamos el hijo (el) del elemento padre

        Whtml = Whtml & " var el = document.getElementById('mt2'); "
        Whtml = Whtml & " var padre = el.parentNode; "
        Whtml = Whtml & " padre.removeChild(el); "

        Whtml = Whtml & " var el = document.getElementById('mt1'); "
        Whtml = Whtml & " var padre = el.parentNode; "
        Whtml = Whtml & " padre.removeChild(el); "

        Whtml = Whtml & " var el = document.getElementById('TXHI3M'); "
        Whtml = Whtml & " var padre = el.parentNode; "
        Whtml = Whtml & " padre.removeChild(el); "


        Whtml = Whtml & " DCEXRCLC();"
        ' Whtml = Whtml & " clIndice();"
        Whtml = Whtml & "</script>"



        If XCopc = "S" Then
            Whtml1 = ""
            Whtml1 = Whtml1 & "<script language=javascript >"
            If SESSION("FULLGUI") = "S" Then
                XCdsc = ""
                Whtml1 = Whtml1 & "   var element = document.getElementById('TBH1VML'); if (element != null ){ document.getElementById('TBH1VML').textpath.string='" & XCdsc & "';}" & Chr(13)
            Else
                Whtml1 = Whtml1 & "   document.getElementById('TBH1DV').innerHTML= '" & XCdsc & "';" & Chr(13)
            End If
            Whtml1 = Whtml1 & "</script>"

            ofv.FBBarr = ofv.FBBarr & Whtml1

        End If






        ''''''''''''''''''''''
        If UCase(Session("ASP")) <> "COMPTREE.ASP" Then
           'response.write request.querystring & " ---------------"
           ' CorPnlIndx = Whtml & "<div id=contm>" & armasolapa(nCodigo, sTipo) & "</div>"  '& "<div id=stickx style='display:none'></div>"
            CorPnlIndx = Whtml '& "<div id=contm>" & armasolapa(nCodigo, sTipo) & "</div>"  '& "<div id=stickx style='display:none'></div>"
            '' CorPnlIndx =CorPnlIndx & "<div id='toolbar-options' style='left: 50%;position: absolute;' class='hidden'>"
            'CorPnlIndx =  CorPnlIndx & "<a href='#' title='recomendar'><i class='la la-users'></i></a>"
            ' CorPnlIndx =  CorPnlIndx & "<a href='#'><i class='fa fa-car'></i></a>"
            '' CorPnlIndx = CorPnlIndx &  "<a href='#'><i class='fa fa-bicycle'></i></a>"
            '  CorPnlIndx =CorPnlIndx &"</div>"
        Else
            CorPnlIndx = Whtml
        End If
        ' For Each i in Session.Contents
        '  Response.Write(i & "---" & session(i) & "<br>")
        'Next

    End Function

    '''' funciones del panel ejecutivo Cor On-Time ''''''''''''''''''''''''''''


    Function CorPnlCOT()

        ''''''''''''''''''''''
        Whtml = ""
        Whtml = Whtml & "<DIV class='DINDM' ID='DHEJ1'  NAME='DHEX1'  style='display:NONE;position:absolute;top:160;left:180;HEIGHT=310px;WIDTH=558px;' >"

        Whtml = Whtml & "<DIV class='DINDC' ID='DHEJ1C' NAME='DHEX1C' style='position:relative;top:+2px;left:+2;HEIGHT=300px;WIDTH=550px;' >"
        Whtml = Whtml & "<CENTER>"

        Whtml = Whtml & "<table id=""MTIJX"" class='TWIND' width=100% cellspacing=0 cellpadding=2 >"
        Whtml = Whtml & "<tr>"
        Whtml = Whtml & "<td CLASS=mbt width=99%  >&nbsp;Panel Ejecutivo</td>"
        Whtml = Whtml & "<td CLASS=mbt width=1% onclick=""clpanel();"" title='Ocultar panel.' >"
        Whtml = Whtml & "<b><font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1   >r</font></b>"
        Whtml = Whtml & "</td>"
        Whtml = Whtml & "</tr>"
        Whtml = Whtml & "</table>"

        Whtml = Whtml & "<DIV ID='DHEXJB' NAME='DHEX2B' style='display:BLOCK;position:relative;top:+1px;left:+2;HEIGHT=270px;WIDTH=544px;overflow:auto;border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >"

        Whtml = Whtml & "</div>"

        Whtml = Whtml & "</CENTER>"
        Whtml = Whtml & "</div>"
        Whtml = Whtml & "</div>"

        Whtml = Whtml & "<SCRIPT language=javascript >" & Chr(13) & Chr(13)
        Whtml = Whtml & " var xspnl = 'NO'; "
        Whtml = Whtml & "function clpanel() { " & Chr(13)
        Whtml = Whtml & "  if  (xspnl=='NO')   " & Chr(13)
        Whtml = Whtml & "      { document.getElementById('DHEXJB').innerHTML=window.top.verfrm4(); " & Chr(13)
        Whtml = Whtml & "        xspnl='SI';} " & Chr(13)

        Whtml = Whtml & "  if  (document.getElementById('DHEJ1').style.display=='none')   " & Chr(13)
        Whtml = Whtml & "      { document.getElementById('DHEJ1').style.display='';} " & Chr(13)
        Whtml = Whtml & "  else                                             " & Chr(13)
        Whtml = Whtml & "      { document.getElementById('DHEJ1').style.display='none';}  " & Chr(13)
        Whtml = Whtml & " } " & Chr(13)

        Whtml = Whtml & "</SCRIPT>" & Chr(13) & Chr(13)

        ''''''''''''''''''''''

        CorPnlCOT = Whtml

    End Function


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Function CorLogo()

        Dim Whtml

        Whtml = ""

        Whtml = Whtml & "<table border=0 cellspacing=0 cellpadding=0 bgcolor=#eeeeee width=100% height=100% style='border-style:ridge;border-width:3;' >"
        Whtml = Whtml & "<tr>"
        Whtml = Whtml & "<td align=center style='border-style:ridge,border-width:2; " & XspSShw("#777777", "135") & ";' >"
        Whtml = Whtml & "<IMG border=0 src='imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
        Whtml = Whtml & "<FONT color=#ffffff face=verdana size=2>"
        Whtml = Whtml & "<B>&nbsp;Cor&nbsp;Consulting</B>"
        Whtml = Whtml & "</FONT>"
        Whtml = Whtml & "</td></tr></table>"

        CorLogo = Whtml

    End Function

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Function CorPanelTop()

        Dim Whtml
        Dim WhtmlX

        WStyl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; "
        WStyl = WStyl & "SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; "
        WStyl = WStyl & "SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; "
        WStyl = WStyl & "SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"

        Whtml = ""
        WhtmlX = ""

        xsscpnm = request.servervariables("script_name")
        xsscpnmi = InStrRev(xsscpnm, "/")
        xsscpnmy = Right(UCase(xsscpnm), Len(xsscpnm) - xsscpnmi)

        If xsscpnmy = "MENUS.ASP" Or xsscpnmy = "SUBMENUS.ASP" Or xsscpnmy = "FOOTER.ASP" Then

            WhtmlX = ""

        Else




            If xsscpnmy = "NOVEDMENU.ASP" Then
                CLASE = ""
            Else
                CLASE = " class=pagina "
            End If
            Whtml0i = "<DIV id='XCPT000I' STYLE='display:block;'  ><DIV  id='MDPTOP1' STYLE='z-index:1000;display:block;' >"

            Whtml1i = ""
            Whtml1i = Whtml1i & " <DIV id='THM' STYLE='display:none;z-index:0;position:absolute;top:0;left:0;width:100%;height:40;' >"
            Whtml1i = Whtml1i & "<br style='line-height:45px;' ></DIV>"


            Whtml0o = "<DIV id='XCPT000O'  ><DIV id='MDPTOP1' STYLE='z-index:0;display:block;' >"

            Whtml1o = ""
            Whtml1o = Whtml1o & " <DIV id='THM' STYLE='z-index:0;position:absolute;top:0;left:0;width:100%;height:40;"

            If session("WEBDESK") = "S" Then


            Else

                Whtml1o = Whtml1o & "background-color:" & Session("CMFMENU") & ";"
                Whtml1o = Whtml1o & XspEx("#f0f0f0", Session("CMFMENU"), 0) & ""

            End If


            Whtml1o = Whtml1o & "' >"
            Whtml1o = Whtml1o & "<br style='line-height:45px;' ></DIV>"


            '''''''''
            Whtml = ""


            If session("WEBDESK") = "S" Then
                Whtml = Whtml & "<TABLE id='MTH'  width=100% marginwidth=0  marginheight=0  "
                Whtml = Whtml & " cellspacing=0 "
            Else
                Whtml = Whtml & "<TABLE id='MTH'  width=98% marginwidth=0  marginheight=0  "
                Whtml = Whtml & " cellspacing=2 "
            End If

            Whtml = Whtml & " cellpadding=2 "
            Whtml = Whtml & " style=" & Chr(34) & "display:none;z-index:90;position:relative;top:2;left:0;"
            Whtml = Whtml & Chr(34)
            Whtml = Whtml & " >"




            Whtml = Whtml & "<TR height='40' >"

            If session("WEBDESK") = "S" Then

                xscliente = "<font id='TXHIS' style='display:block;' face=" & Session("CMFMEFTFC")
                xscliente = xscliente & " size=3 color=" & Session("CMFMENU") & " ><b>&nbsp;" & Session("Cliente") & "</b></font>" '& SESSION("CLIEMP")

                xsclilogo = "<IMG id='TXHIM' border=0 src='" & Session("CliLogo") & "' style='display:block;' >"


            Else

                xscliente = "<font id='TXHIS' style='display:block;' face=" & Session("CMFMEFTFC")
                xscliente = xscliente & " size=2 color=" & Session("CMFMEFT") & " ><b>&nbsp;" & Session("Cliente") & "</b></font>"

                xsclilogo = "<IMG id='TXHIM' border=0 src='" & Session("CliLogo") & "' style='display:block;' >"

            End If


            Whtml = Whtml & "<TD  id='TXHI' CLASS=vt valign=middle width=4% align=center style='display:block;' >"
            Whtml3 = Whtml
            Whtml3i = Whtml3 & "&nbsp;"
            Whtml3o = Whtml3 & xsclilogo

            Whtml = ""
            Whtml = Whtml & "</TD>"




            Whtml = Whtml & "<TD  id='TXHI2' CLASS=vt valign=middle width=96% align='right' >"

            Whtml4 = Whtml


            Whtml4i = Whtml4 & "&nbsp;"

            Whtml4o = Whtml4 & xscliente


            '''''''' sublogo '''''''''''''''''''''

            If session("SUBLOGO") <> "" Then

                Whtml4o = Whtml4o & "</TD>"

                Whtml4o = Whtml4o & "<TD  id='TXHI2ll' CLASS=vt valign=middle width=4% align='center' >"

                Whtml4o = Whtml4o & "<IMG id='TXHIMAD' border=0 src='../coress20/empresa/" & session("SUBLOGO") & "' style='display:block;' >"
            Else


            End If


            '''''''''''''''''''''''''''''''''''''

            Whtml = ""

            Whtml = Whtml & "</TD>"


            xsmenumtr = xsmenumtr & "<TABLE  id=XMH  marginwidth=0  marginheight=0 "
            xsmenumtr = xsmenumtr & " cellspacing=0  cellpadding=0  "
            xsmenumtr = xsmenumtr & " style=" & Chr(34)
            xsmenumtr = xsmenumtr & "visibility:hidden;"
            xsmenumtr = xsmenumtr & Chr(34)
            xsmenumtr = xsmenumtr & " >"
            xsmenumtr = xsmenumtr & "<TR >"
            xsmenumtr = xsmenumtr & "<TD id=XMHZ  CLASS=ut style='font-size:6pt;' >&nbsp;</TD>"
            xsmenumtr = xsmenumtr & "</TR>"
            xsmenumtr = xsmenumtr & "</TABLE>"

            Whtml = Whtml & "<TD id='TXMH' CLASS=vt valign=middle style='visibility:visible;' width=3 >" & xsmenumtr & "</TD>"
            Whtml = Whtml & "</TR>"

            xsFuncion = ""


            If session("WEBDESK") = "S" Then
                Whtml = Whtml & "<TR id='TXHR3S'  >"

            Else
                Whtml = Whtml & "<TR id='TXHR3S' style='display:none;' >"
            End If


            Whtml = Whtml & "<TD id='TXHI3S' colspan=3 CLASS=ut width=100%  valign=middle style='border-style:solid;border-width:1;border-color:#666666;'  >"
            Whtml = Whtml & "</TD>"
            Whtml = Whtml & "</TR>"



            If session("WEBDESK") = "S" Then
                ''''     Whtml = Whtml & "<TR bgcolor='" & Session("CMFMENU") & "' >"
                Whtml = Whtml & "<TR bgcolor='" & Session("CMFMENU") & "' >"

            Else
                Whtml = Whtml & "<TR>"
            End If



            Whtml = Whtml & "<TD id='TXHI3' colspan=3 CLASS=vt width=100% " & xsFuncion & " valign=middle   ><small>&nbsp;</small></TD>"
            Whtml = Whtml & "</TR>"

            If session("WEBDESK") = "S" Then
                Whtml = Whtml & "<TR id='TXHR3M'  bgcolor='" & Session("CMFMNBOT") & "' >"

            Else
                Whtml = Whtml & "<TR id='TXHR3M' style='display:none;' >"
            End If

            ''''''''''align=center valign=middle

            Whtml = Whtml & "<TD id='TXHI3M' colspan=3 CLASS=ut width=100%  style='border-style:solid;border-width:1;border-color:#666666;' ><small>&nbsp;</small>"
            Whtml = Whtml & "</TD>"
            Whtml = Whtml & "</TR>"


            Whtml = Whtml & "<TR id='TXHR4' style='display:none;' >"
            Whtml = Whtml & "<TD id='TXHI4' colspan=3 CLASS=TTT width=100% " & xsFuncion & " valign=top   >"
            Whtml = Whtml & "</TD>"
            Whtml = Whtml & "</TR>"


            'Whtml = Whtml & "</TABLE><div id=cormenux style='z-index:6000'></div>"
            Whtml = Whtml & "</TABLE>"'<div id=cormenux style='z-index:1000'></div>"
  '          if session("APLMDCLI")="BPN" THEN
           ' Whtml = Whtml &  "<div id=menuSH  class=menuSH><i style='font-family:corflaticon;font-size: 40px; ' >q</i></div>"
   '         end if

            Whtml5 = Whtml

            '''''''''''''''''''''''''''''''''''''

            Whtml5i = Whtml5 & Whtml1i
            Whtml5o = Whtml5 & Whtml1o


            '''''''''''''''''''''''''''''''''''''

            Whtml = ""

            xsAccion = "../coress20/inicio.asp?Opcion=I&ddat="
            xsInput = "<input type=text style='display:none;' value='ok' >"
            xFormulario = "<form action='" & xsaccion & "' method='post' id='logoff' name='logoff' style='display:none;'  >" & xsinput & "</form>"

            Whtml = Whtml & xFormulario

            Whtml = Whtml & "</DIV>"

            Whtml = Whtml & "</DIV>"

            Whtml6 = Whtml

            ''''''''''

            Whtml = ""

            Whtml = Whtml & "<script id='XCPT002I' language='javascript' for='TABLE' event='onBLUR' >" & Chr(13)

            Whtml7 = Whtml

            Whtml = ""
            Whtml = Whtml & "<DIV id='XCPT001O' >"

            ''''''''''
            sImagenx = "<B><FONT FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >r</FONT></B>"
            sImagenx1 = "<B><FONT FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >2</FONT></B>"
            sImagenx2a = "<B><FONT FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >5</FONT></B>"
            sImagenx2b = "<B><FONT FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >6</FONT></B>"
            sImagenx3 = "<B><FONT FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0204;</FONT></B>"
            sImagenx4 = "<B><FONT FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >2</FONT></B>"
            xsimagena = "<b><font FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >s</font></b>"
            xsimagenb = "<b><font FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0128;</font></b>"
            xsimagenc = "<b><font FACE=corflaticon COLOR=" & session("CMFMNBOT") & " SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >r</font></b>"

            '''''''
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''' marco de alojamiento del menu desplegable
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Whtml = Whtml & "<DIV id='x1' STYLE='width:400px;display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:5px;top:50px;border-style:outset;border-color:#ffffff;border-width:0;" '1'

            If Session("FULLGUI") = "N" Then
                Whtml = Whtml & "background-color:;" '#aaaaaa'
            Else
                Whtml = Whtml & XspEx("#aaaaaa", "#dddddd", 1) & ";"
            End If

            Whtml = Whtml & "'>"


            ''' tabla superior del menu: aplicacion y botonera
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Whtml = Whtml & "<table id='HX1'   class='HX1CL' width=100% cellspacing=1 cellpadding=0 "
            Whtml = Whtml & " style='display:none;border-style:outset;border-width:1;border-color:#ffffff;"

            If Session("FULLGUI") = "N" Then
                Whtml = Whtml & "background-color:#EEEEEE;"
            Else
                Whtml = Whtml & XspEx(Session("APLBGD"), "#AAAAAA", 0) & ";"
            End If

            Whtml = Whtml & "' >"
            Whtml = Whtml & "<tr height=20 >"
            Whtml = Whtml & "<td  "
            Whtml = Whtml & "class=vT width=2 STYLE='" & XspSShw("#AAAAAA", "125") & ";' >" & "<IMG valign=middle border=0 src='../coress20/imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "<td id='HX1TP' onclick='submenu();' "
            Whtml = Whtml & "class=vT STYLE='" & XspDShw("1", "1", "#f9f9f9") & ";' ><font color=" & session("CMFMNBOT") & " >" & "&nbsp;" & Session("CorTitApli") & "/" & Session("cormenu") & "</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "<td id='SUB' CLASS=mbt onclick='submenu();' title='Menus' "
            Whtml = Whtml & " width=2 STYLE='display:block;' >" & sImagenx4
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "<td id='SOL' CLASS=mbt onclick='solmenu();' title='Solapas' "
            Whtml = Whtml & " width=2 STYLE='display:block;' >" & sImagenx3
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "<td id='MIN' CLASS=mbt onclick='minimenu();' title='Minimizar' "
            Whtml = Whtml & " width=2 STYLE='display:block;' >" & sImagenx2a
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "<td id='MAX' CLASS=mbt onclick='maximenu();' title='Maximizar' "
            Whtml = Whtml & " width=2 STYLE='display:none;' >" & sImagenx2b
            Whtml = Whtml & "</td>"
            '''''''
            Whtml = Whtml & "<TD id=XMH1 CLASS=mbt width=2 onclick='mostrarHLP();' STYLE='display:block;;' >" & xsimagena & "</TD>"
            Whtml = Whtml & "<TD id=XMH2 CLASS=mbt width=2 onclick='mostrarUSR();' STYLE='display:block;' >" & xsimagenb & "</TD>"
            Whtml = Whtml & "<TD id=XMH3 CLASS=mbt width=2 onclick='document.logoff.submit();return false;' STYLE='display:block;' >" & xsimagenc & "</TD>"
            ''''''''
            Whtml = Whtml & "</tr>"
            Whtml = Whtml & "</table>"


            ''' marcos inferiores del menu: solapas y menues
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Whtml = Whtml & "<div id='MM0' style='display:block;overflow:auto;" & WStyl & "' >"
            Whtml = Whtml & "</div>"

            Whtml = Whtml & "<div id='MM1' style='display:block;" & WStyl
            Whtml = Whtml & "' >"
            Whtml = Whtml & "</div>"


            Whtml = Whtml & "<div id='MM2' style='display:block;"
            Whtml = Whtml & "' >"
            Whtml = Whtml & "</div>"

            Whtml = Whtml & "<div id='MM3' style='display:block;' >&nbsp;"
            Whtml = Whtml & "</div>"
            Whtml = Whtml & "</DIV>"


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            ''' fin marco de alojamiento del menu desplegable
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            ''''''''''
            Whtml = Whtml & "<DIV id='PDUSR' class='DINDM'  STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:60px;top:100px; cursor:pointer;"
            Whtml = Whtml & "'>"
            Whtml = Whtml & "<table id=""MTUSR1""  class='TWIND' width=400 cellspacing=0 cellpadding=2 >"
            Whtml = Whtml & "<tr>"
            Whtml = Whtml & "<td class=mbt width=99%  >&nbsp;SESION&nbsp;&nbsp;" & Session.sessionid & "</td>"
            Whtml = Whtml & "<td CLASS=mbt width=1% onclick=""document.getElementById('PDUSR').style.display='none';""  >"
            Whtml = Whtml & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1  >r</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "</tr>"
            Whtml = Whtml & "</table>"
            Whtml = Whtml & "<table id=""MTUSR1""  class='aaa' width=400 cellspacing=4 cellpadding=2 "
            Whtml = Whtml & " style=""border-style:solid;border-width:2;border-color:#eeeeee;background-color:#f9f9f9;"" "
            Whtml = Whtml & " >"
            Whtml = Whtml & "<tr>"
            Whtml = Whtml & "<td class=tt width=25% >&nbsp;Usuario</td>"
            Whtml = Whtml & "<td class=vti width=75%  >"
            Whtml = Whtml & "<font color=#999999 >" & Session("SegUsuario") & "</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "</tr>"
            Whtml = Whtml & "<tr>"
            Whtml = Whtml & "<td class=tt   >&nbsp;Empresa</td>"
            Whtml = Whtml & "<td class=vti    >"
            Whtml = Whtml & "<font color=#999999 >" & session("cliente") & "</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "</tr><tr>"
            Whtml = Whtml & "<td class=tt   >&nbsp;Email</td>"
            Whtml = Whtml & "<td class=vti    >"
            Whtml = Whtml & "<font color=#999999 >" & UCase(Session("UsrMail") & "") & "</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "</tr>"



            Whtml = Whtml & "<tr>"
            Whtml = Whtml & "<td class=tt  colspan=2 >&nbsp;Grupos</td>"
            Whtml = Whtml & "</tr>"

            wktgrupos = Split(session("SegGrpdsc"), Chr(9))

            For wkti = 0 To UBound(wktgrupos)

                Whtml = Whtml & "<tr>"
                Whtml = Whtml & "<td class=vti  colspan=2  >"
                Whtml = Whtml & "<font color=#999999 >-&nbsp;" & wktgrupos(wkti) & "</font>"
                Whtml = Whtml & "</td>"
                Whtml = Whtml & "</tr>"

            Next
            Whtml = Whtml & "</table>"
            Whtml = Whtml & "</DIV>"
            ''''''''''
            Whtml = Whtml & "<DIV id='PDHLP' class='DINDM'  STYLE='display:none;z-index:99;margin-top:1;margin-left:1;position:absolute;left:60px;top:100px;"
            Whtml = Whtml & "'>"
            Whtml = Whtml & "<table id=""MTHLP1""  class='TWIND' width=400 cellspacing=0 cellpadding=2 >"
            Whtml = Whtml & "<tr>"
            Whtml = Whtml & "<td class=mbt width=99% >&nbsp;AYUDA&nbsp;ON-LINE</td>"
            Whtml = Whtml & "<td CLASS=mbt width=1% onclick=""document.getElementById('PDHLP').style.display='none';""  >"
            Whtml = Whtml & "<font onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1  >r</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "</tr>"
            Whtml = Whtml & "</table>"
            'Whtml = Whtml & "<table id=""MTHLP2""  class=aaa width=400 cellspacing=4 cellpadding=2 "
            Whtml = Whtml & " style=""border-style:solid;border-width:3;border-color:#eeeeee;background-color:#f9f9f9;"" "
            Whtml = Whtml & " >"
            Whtml = Whtml & "<tr>"
            Whtml = Whtml & "<td class=tt width=25% >&nbsp;</td>"
            Whtml = Whtml & "<td class=vti width=70%  >"
            Whtml = Whtml & "<font color=#999999 >en breve ......</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "<td class=btbs width=5% >"
            Whtml = Whtml & "<font class=btic >4</font>"
            Whtml = Whtml & "</td>"
            Whtml = Whtml & "</tr>"
            Whtml = Whtml & "</table>"
            Whtml = Whtml & "</DIV>"

            Whtml = Whtml & "</DIV>"

            ''''''''''

            '''''     if session("WEBDESK") = "S" then TXHI3


            Whtml = Whtml & "<script id='XCPT001' language=javascript>" & Chr(13)

            If session("SMPLMENU") = "S" Then
                Whtml = Whtml & " var smpmenu = 'S'; " & Chr(13)
            Else
                Whtml = Whtml & " var smpmenu = 'N'; " & Chr(13)
            End If


            Whtml = Whtml & " function verfrm() { "
            Whtml = Whtml & "    mostrarmenu(); " & Chr(13)
            Whtml = Whtml & " } "

            Whtml = Whtml & "function mostrarmenu(){" & Chr(13)

            '''  if session("SMPLMENU") = "S" or session("WEBDESK") = "S" then

            If session("SMPLMENU") = "S" Then

                Whtml = Whtml & " if (PCorType == 'D' || PCorType == 'K') {" & Chr(13)
                Whtml = Whtml & "   document.getElementById('x1').style.display='';" & Chr(13)
                Whtml = Whtml & "                                                              }" & Chr(13)

            Else
                Whtml = Whtml & "   document.getElementById('x1').style.display='';" & Chr(13)
            End If

            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function hdemenu(){" & Chr(13)
            Whtml = Whtml & "   document.getElementById('x1').style.display='';" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function shwhdemenu(){" & Chr(13)
            Whtml = Whtml & "   if (document.getElementById('MDPTOP1').style.display=='block') {" & Chr(13)
            Whtml = Whtml & "       window.top.CorMMnu='H';                     " & Chr(13)
            Whtml = Whtml & "       document.getElementById('MDPTOP1').style.display='none';  " & Chr(13)

            If SESSION("FULLGUI") = "S" Then
                Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='';  " & Chr(13)
            Else
                Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='';  " & Chr(13)
            End If

            Whtml = Whtml & "        hdemenu();                                }" & Chr(13)
            Whtml = Whtml & "   else {                                          " & Chr(13)
            Whtml = Whtml & "       window.top.CorMMnu='S';                     " & Chr(13)
            Whtml = Whtml & "       document.getElementById('MDPTOP1').style.display=''; " & Chr(13)

            If SESSION("FULLGUI") = "S" Then
                Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='none';  " & Chr(13)
            Else
                Whtml = Whtml & "       document.getElementById('tbhMAX').style.display='none';  " & Chr(13)
            End If

            Whtml = Whtml & "       mostrarmenu();                             }" & Chr(13)
            Whtml = Whtml & "   var Fbtnas = 0; "
            Whtml = Whtml & "   MoveBtnas(); " & Chr(13)
            If SESSION("FULLGUI") = "S" Then
                Whtml = Whtml & " " & Chr(13)
            Else
              '  Whtml = Whtml & " ubimenu(); " & Chr(13)
            End If
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function mostrarUSR(){" & Chr(13)
            Whtml = Whtml & "   document.getElementById('PDUSR').style.display='';" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function mostrarHLP(){" & Chr(13)
            Whtml = Whtml & "   document.getElementById('PDHLP').style.display='';" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function solmenu(){" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MIN').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MAX').style.display='none';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('SOL').style.display='none';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('SUB').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM0').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM1').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM2').style.display='none';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM3').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('x1').style.borderStyle='outset';" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function submenu(){" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MIN').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MAX').style.display='none';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('SOL').style.display='';" & Chr(13)
            Whtml = Whtml & "   if (PCorType == 'D' || PCorType == 'K' ) {       " & Chr(13)
            Whtml = Whtml & "       document.getElementById('SOL').style.display='none';" & Chr(13)
            Whtml = Whtml & "                                 }       " & Chr(13)
            Whtml = Whtml & "   document.getElementById('SUB').style.display='none';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM0').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM2').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM3').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('x1').style.borderStyle='outset';" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)


            Whtml = Whtml & "function subrefresh(){" & Chr(13)
            Whtml = Whtml & " var frm2 = window.top.verfrm2();"
            Whtml = Whtml & " document.getElementById('MM2').innerHTML= '' + frm2 + ''; " & Chr(13)
            Whtml = Whtml & "}" & Chr(13)



            Whtml = Whtml & "function despmenu(){" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM0').style.display='';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM2').style.display='none';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('MM3').style.display='none';" & Chr(13)

            Whtml = Whtml & "   if (PCorType == 'D' || PCorType == 'K') {       " & Chr(13)
            Whtml = Whtml & "       document.getElementById('x1').style.width='400px';                            " & Chr(13)
            Whtml = Whtml & "       document.getElementById('x1').style.left=document.getElementById('TXMH').offsetLeft-400;    " & Chr(13)
            Whtml = Whtml & "                                 }                                     " & Chr(13)
            Whtml = Whtml & "   else {                                                              " & Chr(13)
            Whtml = Whtml & "       document.getElementById('x1').style.width='450px';                            " & Chr(13)
            Whtml = Whtml & "       document.getElementById('x1').style.left=document.getElementById('TXMH').offsetLeft-430;    " & Chr(13)
            Whtml = Whtml & "                                 }                                     " & Chr(13)

            If session("WEBDESK") = "S" Then

                Whtml = Whtml & "   document.getElementById('x1').style.top=document.getElementById('TXHI3').offsetTop +  document.getElementById('MTH').offsetTop - 4;" & Chr(13)

            Else

                Whtml = Whtml & "   document.getElementById('x1').style.top=document.getElementById('TXMH').offsetTop +  document.getElementById('MTH').offsetTop - 4;" & Chr(13)

            End If



            Whtml = Whtml & "   document.getElementById('HX1').cellSpacing=4;" & Chr(13)
            Whtml = Whtml & "   document.getElementById('HX1').style.borderStyle='outset';" & Chr(13)
            Whtml = Whtml & "   document.getElementById('x1').style.borderStyle='none';" & Chr(13)
            Whtml = Whtml & "   if (document.getElementById('MDPTOP1').style.display=='none') {" & Chr(13)
            Whtml = Whtml & "       hdemenu();                                  }" & Chr(13)
            Whtml = Whtml & "                                                    " & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function minimenu(){" & Chr(13)
            Whtml = Whtml & "   despmenu();" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function maximenu(){" & Chr(13)
            Whtml = Whtml & "  mostrarmenu();                         " & Chr(13)
            Whtml = Whtml & "  document.getElementById('SOL').style.display='';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('SUB').style.display='none';" & Chr(13)
            Whtml = Whtml & "  if (PCorType == 'D' || PCorType == 'K') {       " & Chr(13)
            Whtml = Whtml & "  document.getElementById('SOL').style.display='none';" & Chr(13)
            Whtml = Whtml & "                                 }       " & Chr(13)
            Whtml = Whtml & "  document.getElementById('MIN').style.display='';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('MAX').style.display='none';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('MM0').style.display='';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('MM1').style.display='none';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('MM2').style.display='';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('MM3').style.display='';" & Chr(13)
            Whtml = Whtml & "   if (window.top.CorType == 'D' || window.top.CorType == 'K') {       " & Chr(13)
            Whtml = Whtml & "       document.getElementById('x1').style.width='400px';                            " & Chr(13)
            Whtml = Whtml & "                                 }                                     " & Chr(13)
            Whtml = Whtml & "   else {                                                              " & Chr(13)
            Whtml = Whtml & "       document.getElementById('x1').style.width='450px';                            " & Chr(13)
            Whtml = Whtml & "                                 }                                     " & Chr(13)
            Whtml = Whtml & "  document.getElementById('x1').style.left='5px';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('x1').style.top=document.getElementById('MTH').clientHeight;" & Chr(13)
            Whtml = Whtml & "  document.getElementById('HX1').cellSpacing=4;" & Chr(13)
            Whtml = Whtml & "  document.getElementById('HX1').style.borderStyle='outset';" & Chr(13)
            Whtml = Whtml & "  document.getElementById('x1').style.borderStyle='outset';" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function ubimenu(){" & Chr(13)
           ' Whtml = Whtml & " var element = document.getElementById('TBHST'); if (element != null ){ "
            Whtml = Whtml & "   document.getElementById('TBHST').style.left=document.getElementById('TBH1').offsetLeft + 8;" & Chr(13)
            Whtml = Whtml & "   document.getElementById('TBHST').style.top=document.getElementById('TBH').offsetTop + 5;" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)


            Whtml = Whtml & "function locmenu(){" & Chr(13)
            Whtml = Whtml & "   var frmxx = window.top.verInfrm2();"
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function locmenu2(){" & Chr(13)
            '  Whtml = Whtml & "   alert(window.top.CorType);" & chr(13)
            ' Whtml = Whtml & "   alert('vis ' + document.all.XMH.style.visibility);" & chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function MoveBtnas(){" & Chr(13)
           '  Whtml = Whtml & "  alert('1'); "
            Whtml = Whtml & "  if (document.getElementById('btnas') ) { " & Chr(13)

            Whtml = Whtml & "     document.getElementById('btnas').style.border='2';" & Chr(13)

            If CorBarra = "U" Then
                If session("WEBDESK") = "S" Then
                    Whtml = Whtml & " var btnoffset=02;if (document.getElementById('sublogoid') ) {btnoffset=document.getElementById('sublogoid').height ;}"
                    Whtml = Whtml & "     document.getElementById('btnas').style.position='absolute';" & Chr(13)

                    Whtml = Whtml & "     document.getElementById('btnas').style.top=document.getElementById('TBH').offsetTop +  btnoffset;" & Chr(13)
                    Whtml = Whtml & "     document.getElementById('btnas').style.left=document.getElementById('TBH2').offsetLeft + (document.getElementById('TBH2').clientWidth-document.getElementById('btnas').clientWidth);" & Chr(13)
                    'whtml = whtml & "alert(document.getElementById('TBH2').clientWidth);"
                    'Whtml = Whtml & "     document.getElementById('btnas').style.left=document.getElementById('TBH2').offsetLeft + (document.getElementById('TBH2').clientWidth-document.getElementById('btnas').clientWidth);" & chr(13)

                    ''   Whtml = Whtml & "     document.getElementById('TXHI5').innerHTML = document.getElementById('btnas').innerHTML;" & chr(13)

                Else

                    Whtml = Whtml & "     document.getElementById('btnas').style.position='absolute';" & Chr(13)
                    Whtml = Whtml & "     document.getElementById('btnas').style.top=document.getElementById('TBH').offsetTop - 2;" & Chr(13)
                    Whtml = Whtml & "     document.getElementById('btnas').style.left=document.getElementById('TBH2').offsetLeft + (document.getElementById('TBH2').clientWidth-document.getElementById('btnas').clientWidth);" & Chr(13)
                    'Whtml = Whtml & "     document.getElementById('btnas').style.left=document.getElementById('TBH2').offsetLeft + (document.getElementById('TBH2').clientWidth-document.getElementById('btnas').clientWidth);" & chr(13)
                End If

            End If
            Whtml = Whtml & "     document.getElementById('btnas').style.visibility='visible';"
            Whtml = Whtml & " } else { " & Chr(13)
            Whtml = Whtml & "           if (Fbtnas > 10) {} else { Fbtnas++; " & Chr(13)
            Whtml = Whtml & " setTimeout (""MoveBtnas()"", 350); } " & Chr(13)



            Whtml = Whtml & " } " & Chr(13)
                    ' rutina que mueve al header
                    if false then
                    Whtml = Whtml & " var iframeBody = window.top.document;var iframe = iframeBody.getElementById('xheader');" & Chr(13)
			        Whtml = Whtml & "var iframeBody = window.top.document;"
			        Whtml = Whtml & "var innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
				    Whtml = Whtml & "var iframeBody = innerDoc;"
			 		Whtml = Whtml & "var cc=iframeBody.getElementById('XHDM');"
                    Whtml = Whtml & "cc.innerHTML=document.getElementById('TBH').outerHTML;"

                    Whtml = Whtml & " for (const element of document.getElementsByClassName('BBB')){"
                    Whtml = Whtml &   " element.action=element.action;"

                   ' Whtml = Whtml &   " var len =document.getElementById('btnas').length;"

					'Whtml = Whtml &   "    for(var i = 0 ; i< len;i++){"
					'Whtml = Whtml &   "        if(element.elements[i].type == 'input')"
					'Whtml = Whtml &   "            alert(document.getElementById('btnas').elements[i].type);"
    				'Whtml = Whtml &   "}"
				    Whtml = Whtml &   " cc.innerHTML=cc.innerHTML + element.outerHTML;"
					Whtml = Whtml & " }"
				 '	Whtml = Whtml & " document.getElementById('btnas').style.display='none';"
				   'Whtml = Whtml & " document.getElementById('TBH').style.display='none';"
                    end if

                    '----------------------------------------------


            Whtml = Whtml & "}" & Chr(13)


            Whtml = Whtml & "function SetMnn(MnnCod){" & Chr(13)
            Whtml = Whtml & " window.top.CorAMNN = MnnCod;"
            '     Whtml = Whtml & " alert(window.top.CorAMNN); " & chr(13)
            ' Whtml = Whtml & " if (window.top.CorAMNN !=undefined) {" & chr(13)
            Whtml = Whtml & " if (!document.getElementById(window.top.CorAMNN)) {" & Chr(13)
            Whtml = Whtml & "      document.getElementById(window.top.CorAMNN).className='XMN'; " & Chr(13)
            Whtml = Whtml & "         }" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)

            Whtml = Whtml & "function SetSlp(SlpCod){" & Chr(13)
            '      Whtml = Whtml & " alert( window.top.CorASLPANT); " & chr(13)
            '      Whtml = Whtml & " alert( window.top.CorASLP); " & chr(13)

            Whtml = Whtml & " window.top.CorASLPANT=window.top.CorASLP ;window.top.CorASLP = SlpCod;"

            '  Whtml = Whtml & " if (!document.getElementById(window.top.CorASLP) ) {" & chr(13)
            ' Whtml = Whtml & "      document.getElementById(window.top.CorASLP).className='XSL'; " & chr(13)
            '     Whtml = Whtml & "   if (window.top.CorASLPANT !=window.top.CorASLP ) { alert(window.top.CorASLPANT); "
            Whtml = Whtml & " if (window.top.CorASLP !=undefined   &&  window.top.CorASLP !='NO' ){" & Chr(13)
            Whtml = Whtml & "      document.getElementById(window.top.CorASLP).className='XSL'; }" & Chr(13)

            Whtml = Whtml & "  if (window.top.CorASLPANT !=undefined &&  window.top.CorASLPANT !='NO' ){ document.getElementById(window.top.CorASLPANT).className='xx';}     " & Chr(13)

            '  Whtml = Whtml & "else{ window.top.CorASLPANT=window.top.CorASLP;        }" & chr(13)

            Whtml = Whtml & "}" & Chr(13)



            Whtml = Whtml & "function SetMnnSOL(MnnCod){" & Chr(13)
            Whtml = Whtml & " window.top.CorASLPEX = MnnCod;"
            Whtml = Whtml & " alert(document.getElementById(window.top.CorASLPEX)); " & Chr(13)
            Whtml = Whtml & " if (!document.getElementById(window.top.CorASLPEX)) {" & Chr(13)
            Whtml = Whtml & "      document.getElementById(window.top.CorASLPEX).className='XMN'; " & Chr(13)
            Whtml = Whtml & "         }" & Chr(13)
            Whtml = Whtml & "}" & Chr(13)


            Whtml = Whtml & "</script>" & Chr(13)





            Whtml = Whtml & "<script id='XCPT007O' language='javascript' for='TABLE' event='onBLUR' >" & Chr(13)





            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '''' Acciones realizadas al cargar el Panel
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Whtml = Whtml & " var PCorType;" & Chr(13)



            Whtml = Whtml & " if (!window.top.CorType) {" & Chr(13)
            Whtml = Whtml & "     PCorType = 'SA';" & Chr(13)
            Whtml = Whtml & "                                                              }" & Chr(13)
            Whtml = Whtml & " else" & Chr(13)
            Whtml = Whtml & "     {" & Chr(13)
            Whtml = Whtml & "     PCorType = window.top.CorType;" & Chr(13)
            Whtml = Whtml & "                                                              }" & Chr(13)
            Whtml = Whtml & " if (PCorType == 'SA') {" & Chr(13)
            Whtml = Whtml & "                                                              }" & Chr(13)
            Whtml = Whtml & " else" & Chr(13)
            Whtml = Whtml & "     {" & Chr(13)



            Whtml = Whtml & " var frmz = window.top.xclrstr();"


            '''''     if session("WEBDESK") = "S" then TXHI3

            '    Whtml = Whtml & " alert('uno');"
            ' Whtml = Whtml & " PCorType='D';"
            If session("WEBDESK") = "S" Then




                Whtml = Whtml & " if (PCorType == 'D' || PCorType == 'K') {" & Chr(13)

                ' Whtml = Whtml & " alert('aaa');"

                Whtml = Whtml & "      var frm0 = window.top.verfrm0();"
                ''     Whtml = Whtml & " alert('aaa1');"

                Whtml = Whtml & "      var frm2 = window.top.verfrm2();"
                ''     Whtml = Whtml & " alert('aaa2');"

                Whtml = Whtml & "      document.getElementById('TXHI3').innerHTML="

                Whtml = Whtml & "      '<div style=" & Chr(34) & "width:60%;float:left;" & Chr(34) & " >' + frm0 "
                Whtml = Whtml & "      + '</div><div style=" & Chr(34) & "width:40%;float:left;" & Chr(34) & " >' + frm2 + '</div>'; " & Chr(13)
                ' Whtml = Whtml & " alert('aaa3');"

                Whtml = Whtml & "      var frm1 = window.top.verfrm1SP();"
                Whtml = Whtml & "      document.getElementById('TXHI3M').innerHTML='' + frm1 + ''; " & Chr(13)

                'variable donde arma la solapa
              '  Whtml = Whtml & "      var frm1z = window.top.verfrm1SP2();"

                ''  Whtml = Whtml & " alert('bbb');"

                ''  Whtml = Whtml & "      document.getElementById('x1').style.top=document.getElementById('MTH').clientHeight;" & chr(13)

                If Session("FULLGUI") = "S" Then
                    ''       Whtml = Whtml & "  document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight - 4;" & chr(13)
                Else
                    ''       Whtml = Whtml & "  document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight;" & chr(13)
                End If

                '''   Whtml = Whtml & " alert('ccc');"

                Whtml = Whtml & "     }" & Chr(13)

                Whtml = Whtml & " else" & Chr(13)

                Whtml = Whtml & "     {" & Chr(13)

                '  wwwMNN = request.querystring("MNN") & ""
                '   if  wwwMNN <> "" then
                '    Whtml = Whtml & "     window.top.CorwwwAMNN = 'MMN" & "+" & wwwMNN  & "';"
                '   end if




                'Whtml = Whtml & " alert(window.top);"

                Whtml = Whtml & "      var frm0 = window.top.verfrm0();"
                Whtml = Whtml & "      document.getElementById('TXHI3S').innerHTML='' + frm0 + ''; " & Chr(13)

                Whtml = Whtml & "      var frm1 = window.top.verfrm1();"
                Whtml = Whtml & "      document.getElementById('TXHI3').innerHTML='' + frm1 + ''; " & Chr(13)

                Whtml = Whtml & "      var frm2 = window.top.verfrm2();"
                Whtml = Whtml & "      document.getElementById('TXHI3M').innerHTML='' + frm2 + ''; " & Chr(13)


                Whtml = Whtml & " if (document.getElementById(window.top.CorAMNN) )  {   document.getElementById(window.top.CorAMNN).className='XMN'; }" & Chr(13)

                Whtml = Whtml & "  if (document.getElementById(window.top.CorASLP) )  {     document.getElementById(window.top.CorASLP).className='XSL'; }" & Chr(13)




                Whtml = Whtml & "      document.getElementById('x1').style.top=document.getElementById('MTH').clientHeight;" & Chr(13)

                If Session("FULLGUI") = "S" Then
                    Whtml = Whtml & "  document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight - 4;" & Chr(13)
                Else
                    Whtml = Whtml & "  document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight;" & Chr(13)
                End If


                Whtml = Whtml & "     }" & Chr(13)


            Else



                Whtml = Whtml & " var frm0 = window.top.verfrm0();"
                Whtml = Whtml & " document.getElementById('MM0').innerHTML='1' + frm0 + ''; " & Chr(13)

                Whtml = Whtml & " var frm1 = window.top.verfrm1();"
                Whtml = Whtml & " document.getElementById('MM1').innerHTML='2' + frm1 + ''; " & Chr(13)

                Whtml = Whtml & " var frm2 = window.top.verfrm2();"
                Whtml = Whtml & " document.getElementById('MM2').innerHTML= '3' + frm2 + ''; " & Chr(13)


                Whtml = Whtml & " document.getElementById('x1').style.top=document.getElementById('MTH').clientHeight;" & Chr(13)

                If Session("FULLGUI") = "S" Then
                    Whtml = Whtml & " document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight - 4;" & Chr(13)
                Else
                    Whtml = Whtml & " document.getElementById('THM').style.height=document.getElementById('MTH').clientHeight;" & Chr(13)
                End If

            End If

            Whtml = Whtml & " if (PCorType == 'D' || PCorType == 'K') {" & Chr(13)
            Whtml = Whtml & "     document.getElementById('XMH2').style.display='none';" & Chr(13)
            Whtml = Whtml & "     document.getElementById('XMH3').style.display='none';" & Chr(13)
            Whtml = Whtml & "     document.getElementById('SOL').style.display='none';" & Chr(13)

            '  Whtml = Whtml & "     document.getElementById('TXHI4').innerHTML= '' + window.top.CorEsq + '';" & chr(13)

            Whtml = Whtml & "     document.getElementById('TXHR4').style.display='';" & Chr(13)
            Whtml = Whtml & "                                                              }" & Chr(13)
            Whtml = Whtml & " else" & Chr(13)
            Whtml = Whtml & "     {" & Chr(13)
            Whtml = Whtml & "     document.getElementById('TXHI').style.display='';" & Chr(13)
            Whtml = Whtml & "                                                              }" & Chr(13)


            Whtml = Whtml & " despmenu();    " & Chr(13)

            Whtml = Whtml & " mostrarmenu(); " & Chr(13)

            Whtml = Whtml & "                                                              }" & Chr(13)


            Whtml = Whtml & " var Fbtnas = 0; "
            Whtml = Whtml & " MoveBtnas();" & Chr(13)

            Whtml8 = Whtml

            Whtml = ""

            Whtml = Whtml & " var xaction = document.logoff.action; "
            Whtml = Whtml & " document.logoff.action = xaction + window.top.CorUsr + '___' "
            Whtml = Whtml & " + window.top.CorNetUsr  + '___' + window.top.CorDomain + '___' + window.top.CorLogid;  " & Chr(13)

            Whtml = Whtml & "</SCRIPT>" & Chr(13) & Chr(13)
            '  Whtml = Whtml & "<script>var storage = window.localStorage;storage.setItem("&chr(34)&"XCPT007O" &chr(34)&", document.getElementById('XCPT007O').innerHTML);</script>"
            Whtml9 = Whtml

            ''''''''''''' arma rutinas ''''''''''''''''

            WhtmlX = ""

            WhtmlX = WhtmlX & Whtml0i & Whtml3i & Whtml4i & Whtml5i & Whtml6 & Whtml7 & Whtml9

            WhtmlX = WhtmlX & Whtml0o & Whtml3o & Whtml4o & Whtml5o & Whtml6 & Whtml8 & Whtml9

            '''''''''''''''''''''''''''''''''''''''''''


            WhtmlX = WhtmlX & ""

            WhtmlX = WhtmlX & "<SCRIPT id='XCPT003O' language='javascript' for='table' event='onresize' >" & Chr(13) & Chr(13)
            WhtmlX = WhtmlX & " if (PCorType == 'SA') {     " & Chr(13)
            WhtmlX = WhtmlX & "                       }     " & Chr(13)
            WhtmlX = WhtmlX & " else                        " & Chr(13)
            WhtmlX = WhtmlX & "                       {     " & Chr(13)
            WhtmlX = WhtmlX & "      minimenu();            " & Chr(13)
            WhtmlX = WhtmlX & "                       }     " & Chr(13)
            WhtmlX = WhtmlX & "  MoveBtnas();             " & Chr(13)
            WhtmlX = WhtmlX & "</SCRIPT>" & Chr(13) & Chr(13)



            '''''''''
        End If


        CorPanelTop = WhtmlX

    End Function

    Function CorPanelTop2()

        Whtmlx = ""

        xsscpnm = request.servervariables("script_name")
        xsscpnmi = InStrRev(xsscpnm, "/")
        xsscpnmy = Right(UCase(xsscpnm), Len(xsscpnm) - xsscpnmi)

        If xsscpnmy = "MENUS.ASP" Or xsscpnmy = "SUBMENUS.ASP" Or xsscpnmy = "FOOTER.ASP" Then

            WhtmlX = ""

        ElseIf xsscpnmy = "INICIO.ASP" Or xsscpnmy = "COREXPRESS.ASP" Or xsscpnmy = "BARRAENCIERRA2.ASP" Then

            Whtmlx = Whtmlx & "<SCRIPT id='XCPT999I' language='javascript' >"
            Whtmlx = Whtmlx & " var xcptvar = window.top.verfuncion(); "

            Whtmlx = Whtmlx & " document.body.innerHTML += xcptvar.getElementById('XCPT000I').innerHTML; "

            Whtmlx = Whtmlx & " 	var j2   = document.createElement('script'); "
            Whtmlx = Whtmlx & "	j2.type  = 'text/javascript'; "
            Whtmlx = Whtmlx & "	j2.text  = xcptvar.getElementById('XCPT002I').innerHTML; "
            Whtmlx = Whtmlx & "	document.body.appendChild(j2); "




            Whtmlx = Whtmlx & "</SCRIPT>"

            WhtmlX = ""

        Else



            Whtmlx = Whtmlx & "<SCRIPT id='XCPT999O' language='javascript' >"
            ''  Whtmlx = Whtmlx & " var xcptvar = window.top.verfuncion(); "
            '  Whtmlx = Whtmlx & "	alert( ' ' + 'corlogf'); "
            Whtmlx = Whtmlx & "  if (window.top.CorMtopm) { "
            Whtmlx = Whtmlx & "      var wcortopf = window.top.CorMtopf; } "
            Whtmlx = Whtmlx & "  else { "
            Whtmlx = Whtmlx & "      var wcortopf = window.opener.top.CorMtopf; } "
            '  Whtmlx = Whtmlx & "      var wcortopf = window.opener.top.CorMtopf; } "

            '   Whtmlx = Whtmlx & "	alert(wcortopf.name + ' ' + 'corlogf'); "


            '  Whtmlx = Whtmlx & "      var wcortopf = window.top; "
            Whtmlx = Whtmlx & " var xcptvar = wcortopf.verfuncion(); "
            Whtmlx = Whtmlx & "  if (document.body.innerHTML.substring(0,2)=='()') { "
            Whtmlx = Whtmlx & "   document.body.innerHTML=document.body.innerHTML.substring(3,document.body.innerHTML.length) } "
            '----
            Whtmlx = Whtmlx & "  if (xcptvar.getElementById('XCPT000O')){"
            Whtmlx = Whtmlx & " document.body.innerHTML += xcptvar.getElementById('XCPT000O').innerHTML; "
            Whtmlx = Whtmlx & " document.body.innerHTML += '<div>' + xcptvar.getElementById('XCPT001O').innerHTML + '</div>'; "
            Whtmlx = Whtmlx & "};"
            Whtmlx = Whtmlx & " 	var j5   = document.createElement('script'); "
            Whtmlx = Whtmlx & "	j5.type  = 'text/javascript'; "
            Whtmlx = Whtmlx & "  if (xcptvar.getElementById('XCPT001')){"
            Whtmlx = Whtmlx & "	j5.text  = xcptvar.getElementById('XCPT001').innerHTML; "
            Whtmlx = Whtmlx & "};"
            Whtmlx = Whtmlx & "	document.body.appendChild(j5); "

            Whtmlx = Whtmlx & " 	var j6    = document.createElement('script'); "
            Whtmlx = Whtmlx & "	j6.type   = 'text/javascript'; "
            Whtmlx = Whtmlx & "  if (xcptvar.getElementById('XCPT007O')){"
            Whtmlx = Whtmlx & "	j6.text   = xcptvar.getElementById('XCPT007O').innerHTML; "
            Whtmlx = Whtmlx & "};"
            Whtmlx = Whtmlx & "	document.body.appendChild(j6); "


            Whtmlx = Whtmlx & " 	var j3   = document.createElement('script'); "
            ' Whtmlx = Whtmlx & "   alert (xcptvar.getElementById('XCPT003O')); "'
            Whtmlx = Whtmlx & "	j3.type  = 'text/javascript'; "
            Whtmlx = Whtmlx & "	j3.htmlFor   = 'window';"
            Whtmlx = Whtmlx & "	j3.event = 'onresize'; "
            Whtmlx = Whtmlx & "  if (xcptvar.getElementById('XCPT003O')){"
            Whtmlx = Whtmlx & "	j3.text  = xcptvar.getElementById('XCPT003O').innerHTML;" '// alert('2');"
            Whtmlx = Whtmlx & "};"
            Whtmlx = Whtmlx & "	document.body.appendChild(j3);"

           ' Whtmlx = Whtmlx & "     var j15   = document.createElement('script'); "
           ' Whtmlx = Whtmlx & "   j15.type  = 'text/javascript'; "
           ' Whtmlx = Whtmlx & "   j15.text  = window.addEventListener('resize',function (){ if (PCorType == 'SA') { }"
            'Whtmlx = Whtmlx & "    else {  minimenu(); }  MoveBtnas();   }   ,true);"
            'Whtmlx = Whtmlx & "   document.body.appendChild(j15);"


           ' Whtmlx = Whtmlx & " if (PCorType == 'D' || PCorType == 'K') {" & Chr(13)
            ' Whtmlx = Whtmlx & "    document.all.TXHI3M.style.display='none'; "
            '	Whtmlx = Whtmlx & "    document.getElementById('TXHI3M').align='left'; "
            '  Whtmlx = Whtmlx & "    alert(document.all.TXHI3M.innerHTML); "
           ' Whtmlx = Whtmlx & "   } else { "
            ' Whtmlx = Whtmlx & "    document.all.TXHI3M.align='center'; "
            If session("APLMDCLI") = "BC" Then
                '	Whtmlx = Whtmlx & "    document.getElementById('TXHI3M').align='left'; "
            Else
                '	Whtmlx = Whtmlx & "    document.getElementById('TXHI3M').align='center'; "
            End If

            'Whtmlx = Whtmlx & "   } " & Chr(13)


            Whtmlx = Whtmlx & "</SCRIPT>"


        End If


        CorPanelTop2 = Whtmlx

    End Function

    Function CorSetMHeader()

        Dim Whtml

        Whtml = ""

        ' Whtml = Whtml & "<html  xmlns:v='urn:schemas-microsoft-com:vml' >"
        Whtml = Whtml & "<html   >"
        Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"">"

      If session("NUEVOMENU") = "S" Then
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/line-awesome.css"">"
            'Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/rtl.asp"">"
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/jquery-ui.css"">"
          '  Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/jquery.toolbar.css"">"
            Whtml = Whtml & " <link id=packery name=packery  rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/packery.css"">"
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/jquery.minicolors.css"">"
            Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/swiper-bundle.css"">"
            Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/swipercor.css"">"
             Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/select2.css"">"
            ' Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/quill.snow.css"">"
             Whtml = Whtml & " <link type=""text/css"" href=""../v48/line-awesome/css/summernote-lite.min.css"" rel=""stylesheet"">"
            ' Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/quill-table-better.css"">"
     '
        End If
        If session("NUEVOMENU") = "S" Then
            Whtml = Whtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"
             Whtml = Whtml & " <script src='../v48/line-awesome/jquery-ui.js'></script>"
   			Whtml = Whtml & " <script src='../v48/line-awesome/charts.js'></script>"
          	 Whtml = Whtml & " <script src='../v48/line-awesome/jqex.js?v='" & GetVersion("../v48/line-awesome/jqex.js")& "'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/jqbsq.js?v='" & GetVersion("../v48/line-awesome/jqbsq.js")& "'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/jq.js?v='" & GetVersion("../v48/line-awesome/jq.js")& "'></script>"
          '  Whtml = Whtml & " <script src='../v48/line-awesome/packery-docs.min.js'></script>"
           	Whtml = Whtml & " <script src='../v48/line-awesome/jquery.minicolors.js'></script>"
           	Whtml = Whtml & " <script src='../v48/line-awesome/swiper-bundle.min.js'></script>"
          ' 	Whtml = Whtml & " <script src='../v48/line-awesome/quill.js'></script>"
           '	Whtml = Whtml & " <script src='../v48/line-awesome/quill-table-better.js'></script>"
            Whtml = Whtml & "  <script src='../v48/line-awesome/summernote-lite.min.js'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/select2.js'></script>"
            Whtml = Whtml & " <script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & Chr(13)
        End If
        Whtml = Whtml & "<HEAD><meta http-equiv=""expires"" content=""0"" /><meta http-equiv=""pragma"" content=""no-cache"" /><meta http-equiv=""Cache-Control"" content=""no-cache, no-store, must-revalidate"" /> <meta http-equiv='Content-Type' content='text/html; charset=" & XLCHARSET & "'></HEAD>"
        Whtml = Whtml & "<Title>Cor E-Solution Suite V4.8.5</title>"
        Whtml = Whtml & ofv.FormStyler()
        ' Whtml = Whtml & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>"
        ' Whtml = Whtml & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>"
        Whtml = Whtml & "<body class=pagina  " & ofv.wbody & " alink=#fffff vlink=#ffffff link=#ffffff    XSFNK  >"

        Whtml = Whtml & CorPanelTop2()

        CorSetMHeader = Whtml

    End Function

Function CorSetFHeaderEx()

        Dim Whtml

        Whtml = ""

        'Whtml = Whtml & "<html xmlns:v='urn:schemas-microsoft-com:vml'>"
        Whtml = Whtml & "<html >"
       If session("NUEVOMENU") = "S" Then
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/line-awesome.css"">"
            'Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/rt2.asp"">"
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/jquery-ui.css"">"
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/packery.css"">"
    		Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/jquery.minicolors.css"">"
     		Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/swiper-bundle.css"">"
     		Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/swipercor.css"">"
     		 Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/select2.css"">"
     		' Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/quill.snow.css"">"
            ' Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/quill-table-better.css"">"
            Whtml = Whtml & " <link type=""text/css"" href=""../v48/line-awesome/css/summernote-lite.min.css"" rel=""stylesheet"">"


            Whtml = Whtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"
             Whtml = Whtml & " <script src='../v48/line-awesome/jquery-ui.js'></script>"
             Whtml = Whtml & " <script src='../v48/line-awesome/charts.js'></script>"

    		'Whtml = Whtml & " <script src='../v48/line-awesome/jquery.stylesheet.js'></script>"

        'ver si es galicia
       '    Whtml = Whtml & " <script src='../v48/line-awesome/coriconos.asp'></script>"
        '   Whtml = Whtml & " <script src='../v48/line-awesome/packery-docs.min.js'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/select2.js'></script>"
           Whtml = Whtml & " <script src='../v48/line-awesome/jqex.js?v='" & GetVersion("../v48/line-awesome/jqex.js")& "'></script>"
           Whtml = Whtml & " <script src='../v48/line-awesome/jqbsq.js?v='" & GetVersion("../v48/line-awesome/jqbsq.js")& "'></script>"
           Whtml = Whtml & " <script src='../v48/line-awesome/jq.js?v='" & GetVersion("../v48/line-awesome/jq.js")& "'></script>"

          Whtml = Whtml & " <script src='../v48/line-awesome/jquery.minicolors.js'></script>"
          Whtml = Whtml & " <script src='../v48/line-awesome/ swiper-bundle.min.js'></script>"
          'Whtml = Whtml & " <script src='../v48/line-awesome/quill.js'></script>"
		  'Whtml = Whtml & " <script src='../v48/line-awesome/quill-table-better.js'></script>"
          Whtml = Whtml & "  <script src='../v48/line-awesome/summernote-lite.min.js'></script>"

          Whtml = Whtml & " <script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & Chr(13)
         ' Whtml = Whtml & " <script src='../v48/line-awesome/swipercor.js'></script>"

        End If


        Whtml = Whtml & "<HEAD><meta http-equiv=""expires"" content=""0"" /><meta http-equiv=""pragma"" content=""no-cache"" /><meta http-equiv=""Cache-Control"" content=""no-cache, no-store, must-revalidate"" /><META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " content='text/html; charset=" & XLCHARSET & "'></HEAD>"
        ' Whtml = Whtml &  "<link rel=""stylesheet"" type=""text/css"" href=""font/corflaticon/corflaticon.css"">"
        Whtml = Whtml & "<Title>Cor E-Solution Suite V4.8.5</title>"
        ' Whtml = Whtml & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>" & ofv.FormStyler()
      '  Whtml = Whtml & ofv.FormStyler()
        Whtml = Whtml & "<body xclass=pagina  bgcolor=" & session("CMIBTTBI") &">"

     '  Whtml = Whtml & CorPanelTop2()

        CorSetFHeaderEx = Whtml

    End Function

    Function CorSetFHeader()

        Dim Whtml

        Whtml = ""

        'Whtml = Whtml & "<html xmlns:v='urn:schemas-microsoft-com:vml'>"
        Whtml = Whtml & "<html >"
        Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"" >"
        Whtml = Whtml & "<link href=""../coremanager/template6/tabcontent.css"" rel=""stylesheet"" type=""text/css"" />"
        If session("NUEVOMENU") = "S" Then
            Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/line-awesome.css"">"
         '   Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/rtl.asp"">"
           Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/jquery-ui.css"">"
          ' Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/jquery.toolbar.css"">"
        '    Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/jquery.bxslider.css"">"
             Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/packery.css"">"
             Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/jquery.minicolors.css"">"
             Whtml = Whtml & " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/swiper-bundle.css"">"
             Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/swipercor.css"">"
             Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/select2.css"">"
            ' Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/quill.snow.css"">"
            ' Whtml = Whtml & " <link   rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/quill-table-better.css"">"
            Whtml = Whtml & " <link type=""text/css"" href=""../v48/line-awesome/css/summernote-lite.min.css"" rel=""stylesheet"">"
            Whtml = Whtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"

           ' Whtml = Whtml & " <script src='../v48/line-awesome/packery-docs.min.js'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/jquery.minicolors.js'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/swiper-bundle.min.js'></script>"
            'Whtml = Whtml & " <script src='../v48/line-awesome/quill.js'></script>"
			'Whtml = Whtml & " <script src='../v48/line-awesome/quill-table-better.js'></script>"
            Whtml = Whtml & "  <script src='../v48/line-awesome/summernote-lite.min.js'></script>"
            'Whtml = Whtml & " <script src='../v48/line-awesome/swipercor.js'></script>"

            Whtml = Whtml & " <script src='../v48/line-awesome/jquery-ui.js'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/charts.js'></script>"
             Whtml = Whtml & " <script src='../v48/line-awesome/select2.js'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/jqex.js?v='" & GetVersion("../v48/line-awesome/jqex.js")& "'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/jqbsq.js?v='" & GetVersion("../v48/line-awesome/jqbsq.js")& "'></script>"
            Whtml = Whtml & " <script src='../v48/line-awesome/jq.js?v='" & GetVersion("../v48/line-awesome/jq.js")& "'></script>"
          '  Whtml = Whtml & " <script src='../v48/line-awesome/jquery.mark.min.js'></script>"
            Whtml = Whtml & " <script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & Chr(13)

        End If


        Whtml = Whtml & "<HEAD><meta http-equiv=""expires"" content=""0"" /><meta http-equiv=""pragma"" content=""no-cache"" /><meta http-equiv=""Cache-Control"" content=""no-cache, no-store, must-revalidate"" /><META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " content='text/html; charset=" & XLCHARSET & "'></HEAD>"
        ' Whtml = Whtml &  "<link rel=""stylesheet"" type=""text/css"" href=""font/corflaticon/corflaticon.css"">"
        Whtml = Whtml & "<Title>Cor E-Solution Suite V4.8.5</title>"
        ' Whtml = Whtml & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>" & ofv.FormStyler()
        Whtml = Whtml & ofv.FormStyler()
        '   Whtml = Whtml & "<div id=cormenux style='z-index:1000'></div>" &  "<body class=pagina " & ofv.wbody & " alink=#ffffff vlink=#ffffff link=#ffffff    XSFNK  >"
         Whtml = Whtml &  "<body class=pagina " & ofv.wbody & " alink=#ffffff vlink=#ffffff link=#ffffff    XSFNK  >"

        Whtml = Whtml & CorPanelTop2()

        CorSetFHeader = Whtml

    End Function



    Function CorFormHeader()

        Dim Whtml

        Whtml = ""

        Whtml = Whtml & "<TABLE  name=headm id=headm width=100%  marginwidth=0  marginheight=0 "
        Whtml = Whtml & " cellspacing=0  cellpadding=0 style='BORDER-style:outset;border-width:2;' >"
        Whtml = Whtml & "<TR  align=top>"
        Whtml = Whtml & "<TD  CLASS=tt  >"

        Whtml = Whtml & "<table border=0 cellspacing=0 cellpadding=2 bgcolor=#eeeeee  style='border-style:ridge;border-width:3;' >"
        Whtml = Whtml & "<tr>"
        Whtml = Whtml & "<td valign=middle align=center style='border-style:ridge,border-width:2;" & XspSShw("#777777", "135") & ";' >"
        Whtml = Whtml & "<IMG valign=middle border=0 src='imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
        Whtml = Whtml & "<FONT color=#ffffff >"
        Whtml = Whtml & "<B>&nbsp;&nbsp;" & Session("FormN") & "&nbsp;&nbsp;</B>"
        Whtml = Whtml & "</FONT>"
        Whtml = Whtml & "</td></tr></table>"

        Whtml = Whtml & "</TD>"
        Whtml = Whtml & "<TD  CLASS=tt align=right>Usuario: " & Session("SegUsuario") & " </TD>"
        Whtml = Whtml & "</TR>"
        Whtml = Whtml & "</TABLE>"
        Whtml = Whtml & "<hr color=" & Session("CMFMENU") & " size=1>"


        CorFormHeader = Whtml

    End Function

    Function CorFormHeader2(NForm, Nopt)

        Dim Whtml

        sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#005500 SIZE=3>1</FONT></B>"
        sImagenx1 = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>&#0128;</FONT></B>"

        Whtml = ""
        Whtml1 = ""
        Whtml2 = ""
        Whtml3 = ""
        Whtml4 = ""

        xsscpnm = request.servervariables("script_name")
        xsscpnmi = InStrRev(xsscpnm, "/")
        xsscpnmy = Right(UCase(xsscpnm), Len(xsscpnm) - xsscpnmi)


        whtcls = "TTT"
        whtclr = session("CMFMEFT") ' session("CMFMNBOT")
        If session("NUEVOMENU") = "S" Then
            Whtml = Whtml & "<TABLE class='tbh' id='TBH' width=100%  marginwidth=0  marginheight=0 "
        Else
            Whtml = Whtml & "<TABLE class='tbh' id='TBH' width=100%  marginwidth=0  marginheight=0 "
        End If
        Whtml = Whtml & " cellspacing=0  cellpadding=0  xborder=1"

        If Session("FULLGUI") <> "S" Then
            Whtml = Whtml & " style='position:fixed;z-index:999999;left:0px;' "
           ' Whtml = Whtml & " style='position:relative;left:-240px;margin-bottom:20px;' "
            'Whtml = Whtml & " style='display:none' "
        End If

        Whtml = Whtml & "  >"
        Whtml = Whtml & "<TR height=30 >"



        If Session("FULLGUI") = "xS" Then

            Whtml1 = Whtml1 & " <v:line ID='TBH0VML' NAME='TBH0VML' from='0,0' to='1,0' title='Mostrar/esconder encabezado.' onclick='shwhdemenu();' style='cursor:pointer;align:left;z-index:1;position:relative;top:-1;width:10px;height:10px;' > "
            Whtml1 = Whtml1 & "   <v:fill on='True' color='" & session("CMFMNBOT") & "' color2='" & session("CMFMNBOT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
            Whtml1 = Whtml1 & "   <v:stroke on='False' /> "
            Whtml1 = Whtml1 & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
            Whtml1 = Whtml1 & "   <v:path textpathok='True' /> "
            Whtml1 = Whtml1 & "     <v:textpath on='True' string='¿' style=""font-size:'16pt';font-family:'corflaticon';font-weight:bold;""  /> "
            Whtml1 = Whtml1 & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
            Whtml1 = Whtml1 & " </v:line>" & "&nbsp;"

            Whtml1 = Whtml1 & Replace(GenTitulo(NForm, 0), "XXXIDXXX", "TBH1VML")

            '  Whtml1 = ""

            Whtml4 = Whtml4 & " <v:line ID='TBHXVML' NAME='TBHXVML' from='0,0' to='1,0' title='Maximizar menu.' onclick='maximenu();' style='cursor:pointer;align:left;z-index:1;position:relative;top:-1;width:10px;height:10px;' > "
            Whtml4 = Whtml4 & "   <v:fill on='True' color='" & session("CMFMNBOT") & "' color2='" & session("CMFMNBOT") & "' method='sigma' type='gradient' angle='180' opacity='95%'  /> "
            Whtml4 = Whtml4 & "   <v:stroke on='False' /> "
            Whtml4 = Whtml4 & "   <v:shadow on='False' color='#ffffff' weight='15px' /> "
            Whtml4 = Whtml4 & "   <v:path textpathok='True' /> "
            Whtml4 = Whtml4 & "     <v:textpath on='True' string='2' style=""font-size:'16pt';font-family:'corflaticon';font-weight:bold;""  /> "
            Whtml4 = Whtml4 & "   <v:extrusion on='True' backdepth='2' color='#ffffff' metal='True'  rotationcenter='0,0,0'  rotationangle = '0,0' /> "
            Whtml4 = Whtml4 & " </v:line>"

            ''   Whtml4 = ""

            Whtml3 = Whtml3 & " if (PCorType == 'D' || PCorType == 'K') {" & Chr(13)
            Whtml3 = Whtml3 & "     document.getElementById('TBH0VML').textpath.string='¥';" & Chr(13)
            Whtml3 = Whtml3 & "     document.getElementById('TBH1VML').textpath.string=window.top.CorSol;" & Chr(13)
            Whtml3 = Whtml3 & " }" & Chr(13)
            Whtml3 = Whtml3 & " else {" & Chr(13)
            Whtml3 = Whtml3 & "     document.getElementById('TBH0VML').textpath.string='¿';" & Chr(13)
            Whtml3 = Whtml3 & " }" & Chr(13)


        Else

            If Session("FULLGUI") = "M" Then
                whtcls = "TTT"
                whtclr = session("CMFMEFT") 'session("CMFMNBOT")
            Else
                whtcls = "TTB"
                whtclr = Session("CMFPNTITFT")
            End If

            whtcls = "vti"
            whtclr = session("CMFMEFT") 'session("CMFMNBOT")

            Whtml1 = Whtml1 & "&nbsp;"

            Whtml2 = Whtml2 & "<TABLE  class='TBHST' id='TBHST' width=45%  xborder  marginwidth=0  marginheight=0"
            Whtml2 = Whtml2 & " cellspacing=0  cellpadding=0 xstyle='z-index:2;position:absolute;top:2;left:2;' >"
            If session("WEBDESK") = "S" Then
                Whtml2 = Whtml2 & "<TD  WIDTH=8%  style='display:none;' >"
                Whtml2 = Whtml2 & "&nbsp;"
                Whtml2 = Whtml2 & "</TD>"
            Else
                Whtml2 = Whtml2 & "<TD CLASS='" & whtcls & "' WIDTH=8% title='Mostrar/esconder encabezado.' onclick='shwhdemenu();' style='display:none;cursor:pointer;' >"
                Whtml2 = Whtml2 & "<FONT  ID='TBH0DV' NAME='TBH0DV' FACE='CorFlaticon' SIZE=5 COLOR='" & whtclr & "' >&#0191;</FONT>"
                Whtml2 = Whtml2 & "</TD>"
            End If


            If session("WEBDESK") = "S" Then
                Whtml2 = Whtml2 & "<TD CLASS='" & whtcls & "' WIDTH=90% style='display:none;'>"
            Else

                Whtml2 = Whtml2 & "<TD CLASS='" & whtcls & "' WIDTH=90% >"

            End If

            Whtml2 = Whtml2 & "<BIG><DIV ID='TBH1DV' >"
            If session("WEBDESK") = "S" Then
                Whtml2 = Whtml2 & "&nbsp;"
            Else
                Whtml2 = Whtml2 & NForm
            End If


            Whtml2 = Whtml2 & "</DIV></BIG>"
            Whtml2 = Whtml2 & "</TD>"
            Whtml2 = Whtml2 & "</TR>"
            Whtml2 = Whtml2 & "</TABLE><span class=sdelim>&nbsp;</span>"


            Whtml4 = Whtml4 & "<TABLE     class='TBHST' id='TBHXDV'     width=100%  marginwidth=0  marginheight=0 "
            Whtml4 = Whtml4 & " cellspacing=0  cellpadding=0  >"
          '  for each x in Session.Contents
			'  Response.Write(x + " =" & Session.Contents(x) & "<br>")
			'  next
            If session("WEBDESK") = "S" Then
                Whtml4 = Whtml4 & "<TD xCLASS='ut'     marginwidth=0  marginheight=0 cellspacing=0  cellpadding=0 >"
                mmns = Trim(Replace(request.querystring("mnn"), "_", " "))
                whtclr = session("CMFMEFT")


                'If mmns = "" Then mmns = Replace(request.querystring("nombre"), "_", "&nbsp;")
                If mmns = "" Then mmns = Replace(session("FormN"), "_", "&nbsp;")

                if SESSION("FormN") = "NOVEDADES" then mmns=""

                If InStr(UCase(xsscpnmy & ""), "NOVE") > 0  Then
                    Whtml4 = Whtml4 & "<FONT  ID='TBHX0DV' NAME='TBHX0DV'  COLOR='" & whtclr & "' ><imgx src='" & Session("CliLogo") & "'></FONT>"
                Else
                    'Whtml4 = Whtml4 & "<FONT  size=2 ID='TBHX0DV' NAME='TBHX0DV'  COLOR='" & whtclr & "' ><img src='" & Session("CliLogo") & "'><BIG>&nbsp;&nbsp;" & Replace(mmns, " ", "&nbsp;")  & "</BIG></FONT>"
                    Whtml4 = Whtml4 & "<FONT  size=2 ID='TBHX0DV' NAME='TBHX0DV'  COLOR='" & whtclr & "' ></FONT>"
                End If
                '     Whtml2 = Whtml2 & "&nbsp;"
            Else
                Whtml4 = Whtml4 & "<TD CLASS='ut'  title='Maximizar menu.' onclick='maximenu();' style='cursor:pointer;'  >"

                'Whtml4 = Whtml4 & "<FONT  ID='TBHX0DV' NAME='TBHX0DV' FACE='CorFlaticon' SIZE=4 xCOLOR='" & whtclr & "' ><img src='" & Session("CliLogo") & "'></FONT>2</FONT>"
                Whtml4 = Whtml4 & "<FONT  ID='TBHX0DV' NAME='TBHX0DV' FACE='CorFlaticon' SIZE=4 xCOLOR='" & whtclr & "' ></FONT></FONT>"
            End If
            sInputT = "<input type='text' class='dt dtx' name='buscar' placeholder='Buscar documento...' id='buscar' size='25'>"

            celdacc = ""'"<td></td>"

			formx=""'.GenRow("", "", "", "", "", "", "", celda, "si")
		    celdacc = celdacc & " <td >" &  sInputT & "</td>"

			sInput = "<td><input type='submit' value=Buscar class='bietx'   xSTYLE='background:url(IMAGENES/LUPA.GIF) no-repeat;color:#ffffff;font-weight:bold;Xbackground-color:#990000;border-style:NONE;border-width:1;Xborder-color:#dddddd;height:30px;WIDTH:30px;cursor:hand;'></input></td>" 'ofv.GenerarInput("", "", "submit", "", "5", "at STYLE='background:url(IMAGENES/LUPA.GIF) no-repeat;color:#ffffff;font-weight:bold;Xbackground-color:#990000;border-style:NONE;border-width:1;Xborder-color:#dddddd;height:30px;WIDTH:30px;cursor:hand;' ", "")
                	'if(request='buscar'="") then
                	'response.write "<script type='text/javascript'>"
                	'Response.Write "alert('No se encontraron los documentos buscados.');"
    			'Response.Write "</script>"('no se encontraron los documentos ')
                	'Else
                sAccion = "../coremanager/SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX&FORMULARIOS=D"
				formx ="<form target=info method=post action='"& sAccion & "' name=Subconsulta id=Subconsulta>"

			sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX&FORMULARIOS=D"
			'End If
		'	response.write xsscpnmy
         If InStr(UCase(xsscpnmy & ""), "NOVEDMENU") > 0 Then
            Novedform = "CLIMN16317206150930885546227"
			ofv.ObtenerAtributosNOdo Novedform, "", "Tipo", "7"
			ConsultarDocvencer = ofv.AConsultar
			docvencer=""
			if ConsultarDocvencer then
             docvencer="<div class='icono-contenedor'><i ID=docvencer title='Documentos a vencer' class='las la-calendar-day la-2x'></i><span id='contador-burbuja3' class='burbuja'>0</span></div>"
            end if
             s =  "<div style='display: flex; max-width: 500px; width: 70%;height:30px;'><div class=barrafs>" & docvencer & "<div class='icono-contenedor'><i ID=suscripciones title='Suscripciones' class='las la-bell la-2x'></i><span id='contador-burbuja2' class='burbuja'>0</span></div><div class='icono-contenedor'><i id=favoritos title='Autorizaciones' class='las la-user-check la-2x'></i><span id='contador-burbuja' class='burbuja'>0</span></div></div>"
          end if
			 'Botón de configuración
			 's = s & "<button type='button' id='btnConf' title='Configurar' style='xtransform:translateY(-50%); background:none; border:none; cursor:pointer; border-top-left-radius:14px;border-bottom-left-radius: 14px; border: solid 1px white;'>"
			 's = s & "<i class='la la-cog' alt='Mic' style=' color:white; font-size:20px; width:20px; height:20px;'></i>"
			 's = s & "</button>"

             rutaActual = Server.MapPath(Request.ServerVariables("SCRIPT_NAME"))
             rutaActual=ucase(rutaActual)
             if Instr(rutaActual,"AGREGAR.ASP")=0 AND instr(rutaActual,"CORESS20") = 0 AND instr(rutaActual,"COESTDLISDET")=0 AND instr(rutaActual,"CONSULTARCMD")=0 AND instr(rutaActual,"REPORTES.ASP")=0 THEN
			 ' Input de texto (con espacio a la derecha para los íconos)
			  s = s & "<input type='text' id='buscar' name=buscar placeholder='Buscar documento...' class='dtx' xstyle='padding-right:70px; border-left-radius: 24px;border-top-left-radius:0px;border-bottom-left-radius: 0px; flex: 1;'>"


			 ' Botón de búsqueda
			  s = s & "<button xonclick=""document.getElementById('Subconsulta').submit();"" type='button' id='btnBusc' title='Buscar' class='inputbutton' xstyle='position:absolute; right:32px; top:53%; transform:translateY(-50%); background:none; border:none; cursor:pointer;'>"
			  s = s & "<i class='la la-search' alt='Buscar' style='font-size:18px; line-height:1; vertical-align:middle;color:#8a8a8a;'></i>"
			  s = s & "</button>"
             END IF
			 			 ' Botón de micrófono
			 's = s & "<button type='button' id='btnMic' title='Hablar' class='inputmic' xstyle='position:absolute; right:10px; top:53%; transform:translateY(-50%); background:none; border:none;border-left: 0.5px solid grey; cursor:pointer;padding-left:0.5px;'>"
			 's = s & "<i class='la la-microphone' alt='Mic' style='font-size:18px; width:10px; height:15px;'></i>"
			 's = s & "</button>"

			s = s & "</div>"
			formx ="<form style='margin:0' target=info method=post action='"& sAccion & "' name=Subconsulta id=Subconsulta>" & s & "</form>"
			'tablabs =   ofv.GenTabla("", "aaa  STYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "98%", "", "0", "", "-2", "0", "", "", formx, "si")
            'Filad = ofv.genform("Subconsulta", "", sAccion, "", "post", "", tablabs, "si")
            'aca


            Whtml4 = Whtml4  & "</TD>" '<td>"& formx &"</td>"

            Whtml4 = Whtml4 & "</TR>"
            Whtml4 = Whtml4 & "</TABLE>"


            If session("WEBDESK") = "S" Then
               Whtml3 = Whtml3 & " ubimenu();  var StrNOmDoc=''; if (typeof(window.top.CorEsq) != 'undefined' && window.top.CorEsq != null && window.top.CorEsq.length > 60) { " & Chr(13)
                Whtml3 = Whtml3 & " StrNOmDoc=window.top.CorEsq.substring(0,60) + '..+' ;" & Chr(13)
                Whtml3 = Whtml3 & " }else{" & Chr(13)
                Whtml3 = Whtml3 & " StrNOmDoc=window.top.CorEsq;" & Chr(13)
                Whtml3 = Whtml3 & " }; " & Chr(13)



                Whtml3 = Whtml3 & " var datof=''; if( window.top.CorEsq !=undefined) { datof='<br><i title=\""' + window.top.CorEsq.replace(/ /g,'&nbsp;')  +'\"">' + StrNOmDoc.replace(/ /g,'&nbsp;') + '</i>'; };"
                Whtml3 = Whtml3 & "  if( window.top.CorSol !=undefined) {datof=datof + '<br>SOLAPA : ' + window.top.CorSol.replace(/ /g,'&nbsp;') };"
               ' Whtml3 = Whtml3 & " alert('" & xsscpnmy & "');"

                If InStr(UCase(xsscpnmy & ""), "NOVE") > 0 Then
                P1=request.querystring("tab")
                sdatof=TitMenuNovedades(P1)
                 Whtml3 = Whtml3 & " document.getElementById('TBHX0DV').innerHTML=document.getElementById('TBHX0DV').innerHTML   + '" & sdatof  &"';" & Chr(13)
                else
                 Whtml3 = Whtml3 & " document.getElementById('TBHX0DV').innerHTML=document.getElementById('TBHX0DV').innerHTML  + datof  ;" & Chr(13)
                end if

                '	; document.getElementById('TBHX0DV').innerHTML=document.getElementById('TBHX0DV').innerHTML  + if  x === undefined window.top.CorEsq + '<br>Solapa:' + window.top.CorSol;"
            Else
                Whtml3 = Whtml3 & " if (PCorType == 'D' || PCorType == 'K') {" & Chr(13)
                Whtml3 = Whtml3 & "     document.getElementById('TBH0DV').innerHTML='&#0165;';" & Chr(13)
                ''   Whtml3 = Whtml3 & "     document.all.TBH1DV.innerHTML=window.top.CorSol;" & chr(13)
                Whtml3 = Whtml3 & " }" & Chr(13)
                Whtml3 = Whtml3 & " else {" & Chr(13)
                Whtml3 = Whtml3 & "     document.getElementById('TBH0DV').innerHTML='&#0191;';" & Chr(13)
                Whtml3 = Whtml3 & " }" & Chr(13)
                'Whtml3 = Whtml3 & " ubimenu(); " & Chr(13)
            End If




        End If

        StyL2 = " STYLE='cursor:pointer;xdisplay:none;"
        ''   StyL2  = StyL2 & "border-style:solid;border-width:1;border-color:#eeeeee;"
        StyL2 = StyL2 & "' "

'ampliado de 3 a 50
        Whtml = Whtml & "<td id='tbhMAX'  title='abre/cierra' class='" & whtcls & "' xwidth=5% marginwidth=0  marginheight=0 cellspacing=0  cellpadding=0  style='width:30px;padding:0px;'"

         mmns = Trim(Replace(request.querystring("mnn"), "_", " "))
         If mmns = "" Then mmns = Replace(session("FormN"), "_", "&nbsp;")

         if SESSION("FormN") = "NOVEDADES" then mmns=""

         XMESAURL = ""

		 		If session("URLTYP") = "V" Then
		 			XMESAURL = Left(request.servervariables("SCRIPT_NAME"), InStr(Right(request.servervariables("SCRIPT_NAME"), Len(request.servervariables("SCRIPT_NAME")) - 1), "/"))
		 		End If


		 		If session("URLPORT") = "N" Then
		 			XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & XMESAURL
		 		Else
		 			XMESAURL = tiposite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
		 		End If



		 	surx=XMESAURL & "/coremanager/novedmenu.asp"


        SINICIO="<a class=burguer title='INICIO' href='" &  surx & "' target=info>"
        IF instr(rutaActual,"AGREGAR.ASP")=0 AND instr(rutaActual,"COESTDLISDET")=0 AND instr(rutaActual,"CONSULTARCMD")=0 AND instr(rutaActual,"REPORTES.ASP")=0 AND instr(rutaActual,"edobjwfldetPEND")=0   THEN
         SINICIO=SINICIO & "<i class='la la-home la-2x' title='Inicio' ></a></i>"
        END IF
        IF Instr(rutaActual,"AGREGAR.ASP")=0 AND instr(rutaActual,"COESTDLISDET")=0 and instr(rutaActual,"CONSULTARCMD")=0 AND instr(rutaActual,"REPORTES.ASP")=0 AND instr(rutaActual,"edobjwfldetPEND")=0  THEN
        Whtml = Whtml & StyL2 & " align=left xstyle='vertical-align:top;' ><div class=Dinicio style='cursor:pointer;' xstyle='display:none' onclick='cerrarFrame();'><i class='la la-bars la-3x'></i></div>"
        ELSE
         Whtml = Whtml & StyL2 & " align=left xstyle='vertical-align:top;' ><div class=Dinicio style='cursor:pointer;' xstyle='vertical-align: text-top;' onclick='cerrarFrame();'></div>"

        END IF

  '    xx=getinfouser(Session("Usuario"))
	'			suserinf= "&nbsp;&nbsp;Email :" & capitalize(UCase(Session("sgu_mail")& "" ))
	'			suserinf=suserinf & "&nbsp;&nbsp;Grupos:"

	'			wktgrupos = Split(session("SegGrpdsc"), Chr(9))

	'			For wkti = 0 To UBound(wktgrupos)
	'				suserinf=suserinf & " " & capitalize(wktgrupos(wkti))
	'			Next


        'Whtml = Whtml & "<i  class='la la-user la-3x'></i>" &  Session("SegUsuario")  & suserinf & "</div>" & Chr(13)

        Whtml = Whtml & Whtml4

        Whtml = Whtml & "</td><td xvaling=top id=datov  name=datov marginwidth=0  marginheight=0 cellspacing=0  cellpadding=0 class=ntmnn>" & sinicio  & Xmmns &"</td>"
        Whtml = Whtml & "<td valing=middle  id=datobsq name=datobsq marginwidth=0  marginheight=0 cellspacing=0  cellpadding=0 align=right >"  & formx &"</td>"

        Whtml = Whtml & "<TD id='TBH1'  class='tbh1'  >" & Whtml1
        Whtml = Whtml & "</TD>"

        Whtml = Whtml & "<TD  id='TBH2' width=50% class='tbh2' align=right valing=bottom marginwidth=0  marginheight=0 cellspacing=0  cellpadding=0>"
        if session("SUBLOGO") <>"" then
        Whtml = Whtml & " <img id='sublogoid' src='../coress20/empresa/" & session("SUBLOGO") & "'>"
        else
        'Whtml = Whtml & "&nbsp;"
        end if
        Whtml = Whtml & "</TD>"

        Whtml = Whtml & "</TR>"
        Whtml = Whtml & "</TABLE>"



         xx=getinfouser(Session("Usuario"))
	     suserinf=suserinf & "<i xstyle=' font-size: 18px;'>Mail: " & capitalize(UCase(Session("sgu_mail")& "" ))
	     suserinf=suserinf & "<BR>Grupos: "

	     wktgrupos = Split(session("SegGrpdsc"), Chr(9))

	     For wkti = 0 To UBound(wktgrupos)
		suserinf=suserinf & " " & capitalize(wktgrupos(wkti))
	    Next

	      suserinf=suserinf & "</i>"


       ' For Each i in Session.Contents
		'  Response.Write(i & "=" & session(i) &"<br>")
       ' Next
	    suserinf="<div style='display:none;z-index:99999' id=infouser class=userinfo>Usuario:" & session("SGU_nombrew") & "<br>" & xsuserinf& "</div>"
       ' Whtml = Whtml &  suserinf & "<div id=userc style='width:100%;xborder-style:solid;'><div  CLASS=userc style='right:5px;position:fixed;top:50px;z-index:9999'><i  class='la la-user la-2x'></i></div>"  & "<br><br><br><br></div>"
        Whtml = Whtml &  suserinf & "<div id=userc style='width:100%;xborder-style:solid;'><div  CLASS=userc style='right:5px;position:fixed;top:50px;z-index:9999'></div>"  & "<br><br></div>"
        Whtml = Whtml & Whtml2

        Whtml = Whtml & "<script language=javascript >"
        Whtml = Whtml & " if (window.name=='Principal') {"
        Whtml = Whtml & " var StrNOmDoc=''; if (typeof(window.top.CorEsq) != 'undefined' && window.top.CorEsq != null && window.top.CorEsq.length > 60) { " & Chr(13)
		Whtml = Whtml & " StrNOmDoc=window.top.CorEsq.substring(0,60) + '..+' ;" & Chr(13)
		Whtml = Whtml & " }else{" & Chr(13)
		Whtml = Whtml & " StrNOmDoc=window.top.CorEsq;" & Chr(13)
		Whtml = Whtml & " }; " & Chr(13)
		Whtml = Whtml & " var datof=''; if( window.top.CorEsq !=undefined) { datof='<i title=\""' + window.top.CorEsq.replace(/ /g,'&nbsp;')  +'\"">' + StrNOmDoc.replace(/ /g,'&nbsp;') + '</i>'; };"
	    Whtml = Whtml & "  if( window.top.CorSol !=undefined) {datof=datof + '<br>SOLAPA : ' + window.top.CorSol.replace(/ /g,'&nbsp;') };"
	    Whtml = Whtml & "document.getElementById('datov').innerHTML=datof;"
	    Whtml = Whtml & "document.getElementById('datobsq').innerHTML='';"
	    Whtml = Whtml & "document.getElementById('userc').innerHTML='<br><br><br>';"
	    Whtml = Whtml & "}"
	    Whtml = Whtml & " function cerrarFrame(){"
		Whtml = Whtml & "var iframeBody = window.top.document;"
		Whtml = Whtml & " if (window.name=='Principal') {" ' edicion de documentos


		Whtml = Whtml & " var iframe = iframeBody.getElementById('fxtdoc');"
		Whtml = Whtml & " var dato=new String(iframe.cols);"
		Whtml = Whtml & "  if (dato.indexOf('250')==0){"
		Whtml = Whtml & "  iframe.cols='50,0,0,*,0,0,0';"
	    Whtml = Whtml & "   window.top.CorMoc='C';"
	    Whtml = Whtml & "var iframeBody = window.top.document;var iframe = iframeBody.getElementById('WK1');"
	    Whtml = Whtml & "var innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
	    Whtml = Whtml & "var iframeBody = innerDoc;"
	   	Whtml = Whtml & "iframeBody.getElementById('conteM').style.visibility = 'hidden' ;"

	    Whtml = Whtml & "  }"
		Whtml = Whtml & "  else"
		Whtml = Whtml & "  {"
	    Whtml = Whtml & "  iframe.cols='250,0,0,*,0,0,0';"
	    Whtml = Whtml & "var iframeBody = window.top.document;var iframe = iframeBody.getElementById('WK1');"
		Whtml = Whtml & "var innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
		Whtml = Whtml & "var iframeBody = innerDoc;"
		Whtml = Whtml & "iframeBody.getElementById('conteM').style.visibility = '' ;"

		Whtml = Whtml & "   window.top.CorMoc='O';"
		Whtml = Whtml & "  }"
		Whtml = Whtml & "}else{"

        Whtml = Whtml & " var iframe = iframeBody.getElementById('fxt');"
		Whtml = Whtml & " var dato=new String(iframe.cols);"

		Whtml = Whtml & "  if (dato.indexOf('250')==0){"
        Whtml = Whtml & "  iframe.cols='50,*,0,0,0,0,0,0';"
        Whtml = Whtml & "   window.top.CorMoc='C';"
        Whtml = Whtml & "var iframeBody = window.top.document;var iframe = iframeBody.getElementById('estilo');"
		Whtml = Whtml & "var innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
		Whtml = Whtml & "var iframeBody = innerDoc;"
	    Whtml = Whtml & "iframeBody.getElementById('conteM').style.visibility = 'hidden' ;"
		Whtml = Whtml & "  }"
		Whtml = Whtml & "  else"
		Whtml = Whtml & "  {"
		Whtml = Whtml & "  iframe.cols='250,*,0,0,0,0,0,0';"
		Whtml = Whtml & "var iframeBody = window.top.document;var iframe = iframeBody.getElementById('estilo');"
		Whtml = Whtml & "var innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
		Whtml = Whtml & "var iframeBody = innerDoc;"
		Whtml = Whtml & "iframeBody.getElementById('conteM').style.visibility = '' ;"

		Whtml = Whtml & "   window.top.CorMoc='O';"
		Whtml = Whtml & "  } }"



	    Whtml = Whtml & "} "

     '   Whtml = Whtml & " var iframeBody = window.top.document;var iframe = iframeBody.getElementById('xheader');" & Chr(13)
     '   Whtml = Whtml & "var iframeBody = window.top.document;"
     '   Whtml = Whtml &  "var innerDoc = iframe.contentDocument || iframe.contentWindow.document;"
	'    Whtml = Whtml & "var iframeBody = innerDoc;"
 	'	Whtml = Whtml & "var cc=iframeBody.getElementById('XHDM');"
     '   Whtml = Whtml & "cc.innerHTML=document.getElementById('MDPTOP1').outerHTML;"

        Whtml = Whtml & Whtml3
        Whtml = Whtml & " if (window.top.CorMMnu=='H')                                 {" & Chr(13)
        Whtml = Whtml & "     shwhdemenu();" & Chr(13)
        Whtml = Whtml & "                                                              }" & Chr(13)
        Whtml = Whtml & "</script>" & Chr(13)



        CorFormHeader2 = Whtml

    End Function


    Function BrowserName()    'as String
        Dim strContent 'as String

        strContent = Request.ServerVariables("HTTP_USER_AGENT")
        strContent = LCase(strContent)


        'Test for: ms edge
        If InStr(strContent, "edg/") > 0 Then
            BrowserName = "CHROME"
            Exit Function
        End If
        'Test for: Chrome
        If InStr(strContent, "chrome") > 0 Then
            BrowserName = "CHROME"
            Exit Function
        End If

        'Test for: "Firefox"
        If InStr(strContent, "firefox") > 0 Then
            BrowserName = "FIREFOX"
            Exit Function
        End If

        'Test for: Explorer
        If InStr(strContent, "msie") > 0 Then
            BrowserName = "IE"
            Exit Function
        End If


        If InStr(strContent, " trident/") > 0 Then
            BrowserName = "IE"
            Exit Function
        End If


        'If still here ... other


    End Function




    Function BtnReportes(Indice, titulo, xform, sAccion)

        Indice = 16

        If titulo = "" Then
            titulo = "Reporte"
        End If


        xxxx = CorGenBotonXp(Indice, "NO", "NO", "NO", "NO", "algo", "NO", titulo)
        BtnReportes = ofv.GenBoton(CInt(Indice), "esp1Bot", "BT", "S", "", "S", "A", "N", xform & "", sAccion & "")

    End Function

    Function ImgDefault(sCod)







		If wIndDef > wtopeDef Then
			wIndDef = 1
		End If

        if IsArray(aDefCnt) then
		 ImgDefault = aDefCnt(wIndDef)
		else
		 ImgDefault ="ModeloIconos/imagenesV5/defAULT.jpg"
		end if
		wIndDef = wIndDef + 1

	End Function


Function ListarArchivos(rutaCarpetaFisica, filtroExtensiones,rutaw)

    Dim objFSO, objCarpeta, objArchivo, arrFiltros
    Dim arrArchivos(), contador
    Dim bModoFiltro

    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

    contador = 0
    ReDim arrArchivos(0)

    ' Procesamos la lista de extensiones
    If Trim(filtroExtensiones) <> "" Then
        bModoFiltro = True
        ' Limpiamos el string: quitamos espacios, lo ponemos en minúsculas y lo dividimos en un array
        arrFiltros = Split(Replace(LCase(filtroExtensiones), " ", ""), ",")
    Else
        bModoFiltro = False
    End If

    If objFSO.FolderExists(rutaCarpetaFisica) Then
        Set objCarpeta = objFSO.GetFolder(rutaCarpetaFisica)

        For Each objArchivo In objCarpeta.Files
            Dim agregarArchivo
            agregarArchivo = False

            If Not bModoFiltro Then
                ' Si no estamos en modo filtro, agregamos todos los archivos
                agregarArchivo = True
            Else
                ' Modo filtro: verificamos si la extensión del archivo está en nuestra lista
                Dim posPunto, extensionArchivo
                ' Buscamos la posición del último punto en el nombre del archivo
                posPunto = InStrRev(objArchivo.Name, ".")

                If posPunto > 0 Then
                    ' Extraemos la extensión y la pasamos a minúsculas
                    extensionArchivo = LCase(Mid(objArchivo.Name, posPunto + 1))

                    ' Recorremos nuestro array de filtros para ver si hay coincidencia
                    Dim filtroActual
                    For Each filtroActual In arrFiltros
                        If extensionArchivo = filtroActual Then
                            agregarArchivo = True
                            Exit For ' Si la encontramos, no hace falta seguir buscando
                        End If
                    Next
                End If
            End If

            ' Si la bandera es verdadera, agregamos el archivo al resultado
            If agregarArchivo Then
                ReDim Preserve arrArchivos(contador)
                arrArchivos(contador) = rutaw & objArchivo.Name
                contador = contador + 1
            End If
        Next

        Set objArchivo = Nothing
        Set objCarpeta = Nothing
    End If

    Set objFSO = Nothing

    ListarArchivos = arrArchivos

End Function

%>