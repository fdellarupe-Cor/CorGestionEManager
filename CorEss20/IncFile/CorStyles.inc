<%

 Function ConUsrset()


           set XCcn = ofv.conectar(ofv.strconn2)


           StrSql = "Select * From SYSDEFDSKTOPDET WHERE CODSET = 0" & Session("USRSET")
           Set XCrs = OFV.crearconsultaEx(StrSql, XCcn, 1, "")
           IF NOT XCRS.EOF THEN
              DO UNTIL XCRS.EOF

                 SESSION(XCRS(2)) = XCRS(3)

                 XCRS.MOVENEXT
                 LOOP
           END IF

           call ofv.cerrarconn(XCcn)


 END FUNCTION


 Function Conset(Tipo)

  Select Case Tipo

    Case "0"

''' Definido por el usuario


 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#800000"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"


 Session("CMIMENU")        = "back10.gif"

 Session("CMFMENU")        = "#77aadd"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#FFFFFF"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMFMNBOT")       = "#6699cc"


 Session("CMFPNBOT")       = "#6699cc"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#6699cc"
 Session("CMFPNTITFTFC")   = "Tahoma"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#fbeedd"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#888888"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#888888"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#888888"
 Session("CMFPNGRFTSZ")    = "7"

'' Session("CMFPNIP")        = "#f7eedd"
 Session("CMFPNIP")        = "fbeedd"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#888888"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"	'1-encabezado tablas paneles - seccion indice normativo - novedades - font
 Session("CMIETTFT")       = "#888888"  ' 1-encabezado tablas paneles - seccion indice normativo - novedades - color de font
 Session("CMIETTFTSZ")     = "8"	'1-encabezado tablas paneles - seccion indice normativo - novedades - tamaño de font
 Session("CMIETT")         = "#ffffff" ' 1-encabezado tablas paneles - seccion indice normativo - novedades - color de fondo (132510)
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#888888"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#888888"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#f7eedd"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#888888"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#888888"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#888888"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#888888"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#888888"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#888888"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        =  XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#888888"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#888888"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#888888"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#888888"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#888888"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"


    Case "1"

''' Set 1 - rojo/rojo oscuro/verdana

 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#888888"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"

 Session("CMIMENU")        = "Xback10.gif"

 Session("CMFMENU")        = "#d68811"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#FFFFFF"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNBOT")       = "#880000"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"

 Session("CMFPNBOT")       = "#aa0000"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#aa0000"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#eeeeee"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#888888"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#888888"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#DDDDDD"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#888888"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#888888"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#888888"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#888888"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#DDDDDD"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#888888"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#AA0000"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#888888"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#888888"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#888888"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#888888"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#888888"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#888888"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#888888"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#888888"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"

    Case "2"

 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"

 Session("CMIMENU")        = "Xback10.gif"

 Session("CMFMENU")        = "#6699cc"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#FFFFFF"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMFMNBOT")       = "#13B44D"

 Session("CMFPNBOT")       = "#13B44D"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#6699cc"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#F2F2F2"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#000080"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#000080"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#F2F2F2"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#0000aa"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#F2F2F2"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#137ACF"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#000080"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#000080"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#000080"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#000080"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#000080"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"

    Case "3"

''' Set 3 - Celeste/Azul/arial

 Session("APLBGD")         = "#ffffff"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "f7eedd"


 Session("CMFMENU")        = "#993333"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMIMENU")        = "xcorf5.jpg"
 Session("CMFMNBOT")       = "#f7aa22"

 Session("CMFPNBOT")       = Session("CMFMNBOT")
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = Session("CMFMNBOT")
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = Session("APLITBGD")
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = Session("CMFMENU")
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = Session("CMFMENU")
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = Session("CMFMENU")
 Session("CMFPNGRFTSZ")    = "7"

 Session("CMFPNIP")        = Session("APLITBGD")
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = Session("CMFMENU")
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = ""
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "8"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#ffffff"
 Session("CMISTTBL")       = ""
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#000080"
 Session("CMIBTTFTSZ")     = "8"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = ""
 Session("CMIBTTIT")       = ""
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#000080"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = ""
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = ""
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspDShw("2","2",Session("APLSHW")) & ";"
''' Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = Session("CMFMNBOT")
 Session("CORGGFMFTSZ")     = "7"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = Session("CMFMNBOT")
 Session("CORGGFSFTSZ")     = "7"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = Session("CMFMNBOT")
 Session("CORGGFNFTSZ")     = "7"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = Session("CMFMNBOT")
 Session("CORGRFTFTSZ")     = "7"
 Session("CORGRFT")         = ""

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = Session("CMFMNBOT")
 Session("CORGRFCFTSZ")     = "7"
 Session("CORGRFC")         = ""

    Case "4"

''' Set 4 - Celeste/Azul/tahoma

 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#EEEEEE"
 Session("APLITBGD")       = "#99CCEE"

 Session("CMFMENU")        = "#6495ED"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"
 Session("CMIMENU")        = "Xback10.gif"
 Session("CMFMNBOT")       = "#DAA520"

 Session("CMFPNBOT")       = "#DAA520"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#DAA520"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#F7EEDD"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#000080"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#000080"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#F7EEDD"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#0000aa"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#DDDDDD"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#DAA520"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#000080"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#000080"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#000080"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#000080"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#000080"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"

    Case "5"

''' Definido por el usuario


 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#000080"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "eeeeee"

 Session("CMIMENU")        = "xback10.gif"

 Session("CMFMENU")        = "#e0dbd5"
 Session("CMFMEFTFC")      = "verdana"
 Session("CMFMEFT")        = "#1e93a7"
 Session("CMFMEFTSZ")      = "7"

 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#ffffff"
 Session("CMFMNFTSZ")      = "7"
 Session("CMFMNBOT")       = "#1e93a7"

 Session("CMFPNBOT")       = "#1e93a7"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#1e93a7"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#eeeeee"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#000080"
 Session("CMFPNITFTSZ")    = "7"

 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#000080"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"

 Session("CMFPNIP")        = "eeeeee"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#0000aa"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#000080"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#000080"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#000080"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#f7eedd"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#000080"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#993333"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#000080"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#000080"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#000080"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#000080"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#000080"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#000080"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#000080"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#000080"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"


    Case "6"

''' Definido por el usuario


 Session("APLBGD")         = "#FFFFFF"
 Session("APLFTFC")        = "verdana"
 Session("APLFT")          = "#4D4B3D"
 Session("APLFTSZ")        = "7"
 Session("APLSHW")         = "#dddddd"
 Session("APLITBGD")       = "fbeedd"

 Session("CMIMENU")        = ""

 Session("CMFMENU")        = "#B3B2A0"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"

 Session("CMFMNBOT")       = "#828278"
 Session("CMFMNFTFC")      = "verdana"
 Session("CMFMNFT")        = "#FFFFFF"
 Session("CMFMNFTSZ")      = "7"

 Session("CMFPNBOT")       = "#828278"
 Session("CMFPNBOTFTFC")   = "verdana"
 Session("CMFPNBOTFT")     = "#FFFFFF"
 Session("CMFPNBOTFTSZ")   = "7"

 Session("CMFPNTIT")       = "#828278"
 Session("CMFPNTITFTFC")   = "verdana"
 Session("CMFPNTITFT")     = "#FFFFFF"
 Session("CMFPNTITFTSZ")   = "7"

 Session("CMFPNIT")        = "#E7E7E0"
 Session("CMFPNITFTFC")    = "verdana"
 Session("CMFPNITFT")      = "#4D4B3D"
 Session("CMFPNITFTSZ")    = "7"



 Session("CMFPNGN")        = "#ffffff"
 Session("CMFPNGNFTFC")    = "verdana"
 Session("CMFPNGNFT")      = "#4D4B3D"
 Session("CMFPNGNFTSZ")    = "7"

 Session("CMFPNGR")        = "#ffffff"
 Session("CMFPNGRFTFC")    = "verdana"
 Session("CMFPNGRFT")      = "#993333"
 Session("CMFPNGRFTSZ")    = "7"


 Session("CMFPNIP")        = "#E7E7E0"
 Session("CMFPNIPFTFC")    = "verdana"
 Session("CMFPNIPFT")      = "#4D4B3D"
 Session("CMFPNIPFTSZ")    = "7"

 Session("CMIETTFTFC")     = "verdana"
 Session("CMIETTFT")       = "#4D4B3D"
 Session("CMIETTFTSZ")     = "8"
 Session("CMIETT")         = "#ffffff"
 Session("CMIETTBL")       = "1"
 Session("CMIETTIT")       = ""
 Session("CMIETTUL")       = ""

 Session("CMIETXFTFC")     = "verdana"
 Session("CMIETXFT")       = "#4D4B3D"
 Session("CMIETXFTSZ")     = "7"
 Session("CMIETX")         = "#ffffff"
 Session("CMIETXBL")       = ""
 Session("CMIETXIT")       = ""
 Session("CMIETXUL")       = ""

 Session("CMISTTFTFC")     = "verdana"
 Session("CMISTTFT")       = "#4D4B3D"
 Session("CMISTTFTSZ")     = "8"
 Session("CMISTT")         = "#f7eedd"
 Session("CMISTTBL")       = "1"
 Session("CMISTTIT")       = ""
 Session("CMISTTUL")       = ""

 Session("CMISTXFTFC")     = "verdana"
 Session("CMISTXFT")       = "#4D4B3D"
 Session("CMISTXFTSZ")     = "8"
 Session("CMISTX")         = "#ffffff"
 Session("CMISTXBL")       = ""
 Session("CMISTXIT")       = ""
 Session("CMISTXUL")       = ""

 Session("CMIBTTFTFC")     = "verdana"
 Session("CMIBTTFT")       = "#993333"
 Session("CMIBTTFTSZ")     = "9"
 Session("CMIBTT")         = "#ffffff"
 Session("CMIBTTBL")       = "1"
 Session("CMIBTTIT")       = "1"
 Session("CMIBTTUL")       = ""


 Session("CMICTTFTFC")     = "verdana"
 Session("CMICTTFT")       = "#4D4B3D"
 Session("CMICTTFTSZ")     = "8"
 Session("CMICTT")         = "#ffffff"
 Session("CMICTTBL")       = ""
 Session("CMICTTIT")       = ""
 Session("CMICTTUL")       = ""

 Session("CMICTXFTFC")     = "verdana"
 Session("CMICTXFT")       = "#4D4B3D"
 Session("CMICTXFTSZ")     = "8"
 Session("CMICTX")         = "#ffffff"
 Session("CMICTXBL")       = ""
 Session("CMICTXIT")       = ""
 Session("CMICTXUL")       = ""

 Session("CMICTRFTFC")     = "verdana"
 Session("CMICTRFT")       = "#993333"
 Session("CMICTRFTSZ")     = "8"
 Session("CMICTR")         = "#ffffff"
 Session("CMICTRBL")       = ""
 Session("CMICTRIT")       = "1"
 Session("CMICTRUL")       = ""

 Session("CMICXRFTFC")     = "verdana"
 Session("CMICXRFT")       = "#4D4B3D"
 Session("CMICXRFTSZ")     = "8"
 Session("CMICXR")         = "#ffffff"
 Session("CMICXRBL")       = "1"
 Session("CMICXRIT")       = ""
 Session("CMICXRUL")       = ""

' Arbol de Estructura .....................................

 Session("CBRDBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CBRDLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CBRDRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CBRDTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CBRDICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CBRDABL")        = XspSShw(Session("APLSHW"),"125") & ";"

' Organigrama Grafico .....................................

 Session("CORGBOT")        = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFMENU") & ";"
 Session("CORGLFT")        = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("CMFMENU") & ";"
 Session("CORGRGT")        = "border-right-style:solid;border-right-width:1;border-right-color:" & Session("CMFMENU") & ";"
 Session("CORGTOP")        = "border-top-style:solid;border-top-width:1;border-top-color:" & Session("CMFMENU") & ";"
 Session("CORGICO")        = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"
 Session("CORGTABBRD")     = "border-style:ridge;border-width:1;border-Color:#dddddd;" & XspSShw(Session("APLSHW"),"125") & ";"
 Session("CORGREFBRD")     = "border-style:solid;border-width:1;border-color:" & Session("CMFMENU") & ";"

 Session("CORGGFMFTFC")     = "verdana"
 Session("CORGGFMFT")       = "#4D4B3D"
 Session("CORGGFMFTSZ")     = "8"
 Session("CORGGFM")         = "#ffffff"

 Session("CORGGFSFTFC")     = "verdana"
 Session("CORGGFSFT")       = "#4D4B3D"
 Session("CORGGFSFTSZ")     = "8"
 Session("CORGGFS")         = "#ffffff"

 Session("CORGGFNFTFC")     = "verdana"
 Session("CORGGFNFT")       = "#4D4B3D"
 Session("CORGGFNFTSZ")     = "8"
 Session("CORGGFN")         = "#ffffff"

 Session("CORGRFTFTFC")     = "verdana"
 Session("CORGRFTFT")       = "#4D4B3D"
 Session("CORGRFTFTSZ")     = "8"
 Session("CORGRFT")         = "#ffffff"

 Session("CORGRFCFTFC")     = "verdana"
 Session("CORGRFCFT")       = "#4D4B3D"
 Session("CORGRFCFTSZ")     = "8"
 Session("CORGRFC")         = "#ffffff"



    Case Else

        Conset("1")


  End Select




 End Function

 Function XspEx(XspCl1,XspCl2,XspTyp)

     if   Session("FULLGUI") = "S" then

          XspEx = "filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & XspCl1 & ", endColorStr=" & XspCl2 & ", gradientType=" & XspTyp & ")"

     elseif  Session("FULLGUI") = "M" then

          XspEx = "filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & XspCl1 & ", endColorStr=" & XspCl2 & ", gradientType=" & XspTyp & ")"

     else

          XspEx = "background-color:" & XspCl2 & ";" & XspAlpha(0,100,1) & ""

     end if

 End Function

 Function XspAlpha(XspCl1,XspCl2,XspTyp)

     if   Session("FULLGUI") = "S" then

          XspAlpha = "filter: alpha(opacity=" & XspCl1 & ",finishOpacity=" & XspCl2 & ",STYLE=" & XspTyp & ")"

     elseif  Session("FULLGUI") = "M" then

          XspAlpha = "filter: alpha(opacity=" & XspCl1 & ",finishOpacity=" & XspCl2 & ",STYLE=" & XspTyp & ")"

     else

        '''  XspAlpha = "filter: alpha(opacity=100,finishOpacity=90,STYLE=" & XspTyp & ")"

          XspAlpha = ""

     end if

 End Function

 Function XspSShw(XspCl1,XspAng)

     if   Session("FULLGUI") = "S" then

          XspSShw = "filter:shadow(color=" & XspCl1 & ",direction=" & XspAng & ")"

     elseif  Session("FULLGUI") = "M" then

          XspSShw = "filter:shadow(color=" & XspCl1 & ",direction=" & XspAng & ")"

     else

        ''  XspSShw = "border-right-style:solid;border-right-width:1;border-right-color:" & XspCl1 & ";border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & XspCl1 & ""
          XspSShw = ""

     end if

 End Function

 Function XspDShw(XspOFX,XspOFY,XspCl1)

     if   Session("FULLGUI") = "S" then

          XspDShw = "filter: dropShadow(Color=" & XspCl1 & ",OFFX=" & XspOFX & ",OFFY=" & XspOFY & ")"

     elseif  Session("FULLGUI") = "M" then

          XspDShw = "filter: dropShadow(Color=" & XspCl1 & ",OFFX=" & XspOFX & ",OFFY=" & XspOFY & ")"

     else

        ''  XspDShw = "border-right-style:solid;border-right-width:1;border-right-color:" & XspCl1 & ";border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & XspCl1 & ""
          XspDShw = ""
     end if

 End Function

 Function CorStyler()

   txsscpnm  = request.servervariables("script_name")
   txsscpnmi = instrrev(txsscpnm,"/")
   txsscpnmy = right(ucase(txsscpnm),len(txsscpnm)-txsscpnmi)

   if  txsscpnmy = "INICIO.ASP" then

       CorStyler = " "

   else
       wcorstyl = "<script language='javascript' > "

         wcorstyl = wcorstyl &" function ie_ver(){  "
         wcorstyl = wcorstyl &" msie = parseInt((/msie (\d+)/.exec(navigator.userAgent.toLowerCase()) || [])[1]);"
	       wcorstyl = wcorstyl &" if (isNaN(msie)) {"
	       wcorstyl = wcorstyl &" msie = parseInt((/trident\/.*; rv:(\d+)/.exec(navigator.userAgent.toLowerCase()) || [])[1]);if (isNaN(msie)) {msie=-1;}"
          wcorstyl = wcorstyl &"}else {msie=-1;"
         wcorstyl = wcorstyl  &"} return msie;}"
'window.parent.opener.location //posible para reemplazar window.opener   alert(ie_ver); '
'IE no Reconoce al objeto CorMtopm'
             wcorstyl = wcorstyl & "  if (ie_ver() < 19 )  {  "
             wcorstyl = wcorstyl & "  if (window.top.CorMtopm) {  "
                 wcorstyl = wcorstyl & "      var wcortopf = window.top.CorMtopf;} "
	         wcorstyl = wcorstyl & "  else {"
	         wcorstyl = wcorstyl & "         if (  window.opener.top.CorMtopf) { "
	         wcorstyl = wcorstyl & "             var wcortopf = window.opener.top.CorMtopf; } "
	         wcorstyl = wcorstyl & "         else { ;"
	         wcorstyl = wcorstyl & "             var wcortopf = window.opener; "
	 ''  sHTML = sHTML & " alert(wcortopf.name);" & chr(13)
	         wcorstyl = wcorstyl & "            for (ii=0;ii<10;ii++) {" & chr(13)
	  ' sHTML = sHTML & " alert(CorMtopm.top.name + ' ' + ii);" & chr(13)
	         wcorstyl = wcorstyl & "               if (wcortopf.top.name=='superm') {" & chr(13)
	         wcorstyl = wcorstyl & "                   wcortopf = wcortopf.top;" & chr(13)
	  'sHTML = sHTML & " alert(CorMtopf.name);" & chr(13)
	   ''sHTML = sHTML & " alert(verfuncion());" & chr(13)
	         wcorstyl = wcorstyl & "               break; " & chr(13)
	         wcorstyl = wcorstyl & "               } else {" & chr(13)
	         wcorstyl = wcorstyl & "                 wcortopf = wcortopf.top.opener; " & chr(13)
	         wcorstyl = wcorstyl & "               }" & chr(13)
	         wcorstyl = wcorstyl & "             }" & chr(13)
	         wcorstyl = wcorstyl & "          } "
	         wcorstyl = wcorstyl & "   } "
	         wcorstyl = wcorstyl & "document.write ('<style>'); "
	         wcorstyl = wcorstyl & "document.write (wcortopf.verestilo()); "
	         wcorstyl = wcorstyl & "document.write ('</style>');}"
	         wcorstyl = wcorstyl & " else {"
	         wcorstyl = wcorstyl & " var storage = window.localStorage;"
             wcorstyl = wcorstyl & " var stilo = storage.getItem(" &chr(34) &"estil" &chr(34)&");"
	         wcorstyl = wcorstyl & "document.write ('<style>');  "
			 	wcorstyl = wcorstyl & "document.write (stilo); "
	         wcorstyl = wcorstyl & "document.write ('</style>');"
	         wcorstyl = wcorstyl & "}"


          '// por la transicion nueva para ordenar elementos y despues mostrar en jquery
          wcorstyl = wcorstyl & "window.onload = function() {document.body.style.opacity = '1';};</script>"

         CorStyler = wcorstyl
   end if

 End Function

 Function CorStyler2()

         if Session("CMFMEFTFC") = "" then
            Session("CMFMEFTFC")      = Session("CMFMNFTFC")
            Session("CMFMEFT")        = Session("CMFMNFT")
            Session("CMFMEFTSZ")      = Session("CMFMNFTSZ")
         end if

'''filter: gray;
'''filter: xray;

         s = "<META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " CONTENT=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & ">"

         's= " <link rel=""stylesheet"" type=""text/css"" href=""../v48/line-awesome/css/menu.asp"">"
         s =s &"<style id='estil' >"

        s = s & "BODY.menu {background:" & Session("CMFMENU") & ";"
s = s & "color:" & Session("CMFMENU") & ";"
s = s & "min-width:200px;"
s = s & "box-shadow:0 4px 12px rgba(0,0,0,0.2);"
s = s & "transition:opacity 0.2s ease;"
s = s & "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;}"
s = s & "}"

's = s & "BODY {SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;}"


         s = s & "p{margin-top :0px;margin-bottom:0px;font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
         s = s & "a{text-decoration:none;}"
         s = s & "a.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;}"
         s = s & "a.link{color:#ffffff;text-decoration:none;}"
         s = s & "a.istt{font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:" & Session("CMISTTFTSZ") & "pt;" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & ";}"
		 'background-color: " & Session("CMISTT") & "'
		 'color: " & Session("CMISTTFT") & "'
        ' s = s & "a.istt{color:#000000;text-decoration:none;}"
         s = s & "a.nn{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:#ffffff;text-transform:capitalize;font-weight:500;text-decoration:none;}"

s = s & "a.nt{cursor:pointer;transition: opacity 0.2s ease-in-out;padding-right:5px;border-radius:10px;padding-left:5px;font-weight:500;font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNGNFT") & ";text-decoration:none;}"

s = s & "a.docs{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNBOT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:500;}"
s = s & "a.docs:hover{text-decoration:underline;}"

s = s & "a.selec{font-family:" & Session("CMFPNBOTFTFC") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "color:" & Session("CMFPNBOT") & ";}"
s = s & "a.selec:hover{text-decoration:underline;}"

s = s & "* {box-sizing:border-box;}"

's = s & "li.documento-link {"
's = s & "display: block;"
's = s & "padding: 10px 12px;"
's = s & "margin-bottom: 6px;"
's = s & "border: 1px solid #ddd;"
's = s & "border-left: 5px solid #ccc;"  línea gris por defecto
's = s & "border-radius: 8px;"
's = s & "background-color: #f9f9f9;"
's = s & "box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);"
's = s & "color:" & Session("CMFPNIPFT") & ";"
's = s & "font-size:" & Session("CMFPNIPFTSZ") & "pt;"
's = s & "font-family:" & Session("CMFPNIPFTFC") & ";"
's = s & "font-weight: 500;"
's = s & "text-decoration: none;"
's = s & "cursor:pointer;"
's = s & "transition: all 0.2s ease-in-out;"
's = s & "}"

s = s & "a.nt:hover {text-decoration: underline;opacity: 0.7;}"
s = s & "a.nt:active {text-decoration: underline;opacity: 0.6;}"



         s = s & "a.ntslp{font-size:" & Session("CMFMEFTSZ") & "pt;font-family:" & Session("CMFMEFTFC") & ";color:" &  Session("CMFMEFT")& ";font-weight:500;xtext-decoration:none;}"
         s = s & "a.ntmnn2{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & session("CMIBTTFT") & ";font-weight:500;text-decoration:none;}"

         'originals = s & "a.ntmnn{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & session("CMFMNFT") & ";text-transform: capitalize !important;}"
s = s & "a.ntmnn{"
s = s & "font-size:" & Session("CMFMNFTSZ") & "pt;"
s = s & "font-family:" & Session("CMFMNFTFC") & ";!important"
s = s & "font-weight:500;"
s = s & "color:" & Session("CMFMNFT") & "!important;"
s = s & "text-transform: capitalize !important;"
s = s & "display: block;"
s = s & "padding: 10px 14px;"
s = s & "margin: 4px 0;"
s = s & "border-radius: 6px;"
s = s & "text-decoration: none;"
s = s & "cursor: pointer;"
s = s & "transition: all 0.2s ease-in-out;"
s = s & "}"

s = s & ".item-text2{"
s = s & "flex-grow: 1;"
s = s & "transition: padding 0.2s ease;"
s = s & "background-color: lightgray;"
s = s & "border-radius: 0px 10px 6px 0px;"
s = s & "padding-bottom: 5px;"
s = s & "}"

s = s & "a.ntmnn3{"
s = s & "font-size:" & Session("CMFMNFTSZ") & "pt;"
s = s & "font-family:" & Session("CMFMNFTFC") & ";!important"
s = s & "font-weight:500;"
s = s & "color:" & Session("CMFMNFT") & "!important;"
s = s & "text-transform: capitalize !important;"
s = s & "display: block;"
s = s & "padding: 5px 14px;"
s = s & "margin-bottom: 4px;"
s = s & "border-radius: 0px 6px 6px 0px;"
s = s & "text-decoration: none;"
s = s & "background-color: slategray;"
s = s & "cursor: default;"
s = s & "}"

s = s & "a.ntmnn:hover{"
's = s & "background-color:" & Session("CMFMNFT") & ";"
's = s & "color:" & Session("CMFMENU") & ";"
s = s & "transform: translateY(-3px)" & ";"
s = s & "}"

s = s & "td.ntmnn{"
s = s & "font-size:" & Session("CMFMNFTSZ") & "pt;"
s = s & "font-family:" & Session("CMFMNFTFC") & ";"
s = s & "color:" & Session("CMFMNFT") & ";"
s = s & "text-transform: capitalize !important;"
s = s & "display: block;"
s = s & "padding: 10px 14px;"
s = s & "margin: 4px 0;"
s = s & "border-radius: 6px;"
s = s & "text-decoration: none;"
s = s & "transition: background-color 0.2s ease, color 0.2s ease;"
s = s & "}"


s = s & "a.ntmnn.spcex {"
s = s & "font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif;"
s = s & "color:#dcdde1;"
s = s & "}"

s = s & "a.ntmnn.spcex2 {"
s = s & "font-family: Segoe UI, Tahoma, Geneva, Verdana, sans-serif;"
s = s & "color:#dcdde1;"
s = s & "}"

	  's = s & "div.item-text.spcapp {"
	  's = s & "font-family: Verdana, Geneva, sans-serif;"
      's = s & "padding: 10px 15px;"
	  's = s & "font-size: 11px;"
      's = s & "font-weight: 600;"
      's = s & "color: #b0b5bd;"
      's = s & "text-transform: uppercase;"
      's = s & "letter-spacing: 0.5px;"
      's = s & "margin-top: 15px;"
      's = s & "margin-bottom: 5px;"
	  's = s & "border-top: 1px solid #ccc;"
      's = s & "border-bottom: 1px solid #ccc;"
      's = s & "}"

	  s = s & "div.item-text.spcapp2 {"
	  s = s & "font-size: 11px;"
      s = s & "font-weight: bold;"
	  s = s & "font-family: Segoe UI;"
      s = s & "color: " & Session("CMFMEFT") &";"
      s = s & "text-transform: uppercase;"
      s = s & "letter-spacing: 0.5px;"
      s = s & "padding: 10px 20px;"
      's = s & "border-top: 1px solid #555;"
      's = s & "border-bottom: 1px solid #555;"
      s = s & "margin: 20px 0 10px;"
	  s = s & "}"

	  	  s = s & "div.item-text.spcapp {"
	  s = s & "font-size: 11px;"
      s = s & "font-weight: bold;"
	  s = s & "font-family: Segoe UI;"
      s = s & "color: " & Session("CMFMEFT") &";"
      s = s & "text-transform: uppercase;"
      s = s & "letter-spacing: 0.5px;"
      s = s & "padding: 10px 20px;"
      's = s & "border-top: 1px solid #555;"
      's = s & "border-bottom: 1px solid #555;"
      s = s & "margin: 20px 0 10px;"
	  s = s & "}"

	  s = s & "div.userm {"
	  s = s & "font-size: 11px;"
      s = s & "font-weight: bold;"
	  s = s & "font-family: Segoe UI;"
      s = s & "color: #b0b5bd;"
      s = s & "text-transform: uppercase;"
      s = s & "letter-spacing: 0.5px;"
      s = s & "border-top: 1px solid " & Session("CMFMNFT") & ";"
      s = s & "border-bottom: 1px solid " & Session("CMFMNFT") & ";"
      s = s & "margin: 20px 0 10px;"
	  s = s & "}"

	  s = s & "ul.ulmenu {"
	  s = s & "scrollbar-width: thin;"
	  s = s & "scrollbar-color: rgba(0, 0, 0, 0.2) transparent;"
	  s = s & "}"

	  s = s & "ul {"
	  s = s & "padding:0px;"
	  s = s & "margin:0px;"
	  s = s & "}"

	  s = s & "ulvencer {"
	  s = s & "border-left-color: #fbbf24;"
	  s = s & "background-color: #fffbea;}"


	  s = s & "div.userc {"
	  s = s & "cursor: pointer;"
	  s = s & "}"



	  s = s & "a.nti {"
s = s & "font-size:" & Session("CMFPNIPFTSZ") & "pt;"
s = s & "display: block;"
s = s & "padding: 10px 8px;"
s = s & "margin-bottom: 5px;"
s = s & "border: 1px solid #ddd;"
s = s & "border-radius: 8px;"
s = s & "background-color: #f9f9f9;"
s = s & "color:" & Session("CMIETT") & ";"
s = s & "text-decoration: none;"
s = s & "box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);"
s = s & "transition: all 0.2s ease-in-out;"
s = s & "font-weight: 700;"
s = s & "text-overflow: ellipsis;"
s = s & "}"

s = s & "a.nti:hover {"
s = s & "background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);"
s = s & "border-color:" & Session("CMIETT") & ";"
s = s & "box-shadow: 0 4px 10px rgba(0, 0, 0, 0.08);"
s = s & "transform: translateY(-2px);"
s = s & "}"

s = s & "a.nti:active {"
s = s & "background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 70%);"
s = s & "}"

'iconos del Sistema
s = s & ".la-play {position: absolute;top: 50%;left: 50%;transform: translate(-50%, -50%);font-size: 6rem;color: white;pointer-events:none;text-shadow:0 0 20px rgba(0,0,0,0.5);transition: all 0.2s ease-in-out;z-index:10;}"
s = s & ".la-play:hover {opacity:0.7;}"
s = s & ".contenedor i:hover{opacity: 1;}"
s = s & ".tbh i:hover{filter: brightness(1.15) !important;}"
s = s & ".tbh i:active{filter: brightness(1.3) !important;}"
s = s & ".aaa i:hover{opacity:0.8;}"
s = s & ".aaa i:active{opacity:0.6;}"
s = s & ".aaa la-info-circle{font-size:1rem !important;color: #ff0000 !important;}"
s = s & ".aaa la-info-circle:hover{opacity:1;}"
s = s & ".aaa .la-arrow-right{font-size: 11pt;color:" & Session("CMFPNBOT") &";border-radius: 100px;padding: 0px 4px;border: 1px solid "& Session("CMFPNBOT") &";text-indent: 0px !important;}"
s = s & ".aaa .la-arrow-right:hover{background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 50%);opacity:1;}"
s = s & ".aaa .la-arrow-right:active{background-color:" & Session("CMFPNBOT") & ";color:white;}"
s = s & ".aaa .la-times{font-size: 11pt;color:" & Session("CMFPNBOT") &";border-radius: 100px;padding: 0px 4px;border: 1px solid "& Session("CMFPNBOT") &";text-indent: 0px !important;}"
s = s & ".aaa .la-times:hover{background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 50%);opacity:1;}"
s = s & ".aaa .la-times:active{background-color:" & Session("CMFPNBOT") & ";color:white;}"
s = s & ".aaa .la-upload{font-size: 11pt;color:" & Session("CMFPNBOT") &";border-radius: 100px;padding: 0px 4px;border: 1px solid "& Session("CMFPNBOT") &";text-indent: 0px !important;}"
s = s & ".aaa .la-upload:hover{background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 50%);opacity:1;}"
s = s & ".aaa .la-upload:active{background-color:" & Session("CMFPNBOT") & ";color:white;}"
s = s & ".aaa .la-plus{font-size: 11pt;color:" & Session("CMFPNBOT") &";border-radius: 100px;padding: 0px 4px;border: 1px solid "& Session("CMFPNBOT") &";text-indent: 0px !important;}"
s = s & ".aaa .la-plus:hover{background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 50%);opacity:1;}"
s = s & ".aaa .la-plus:active{background-color:" & Session("CMFPNBOT") & ";color:white;opacity:1;}"
s = s & ".aaa a .la-user{font-size: 11pt;color:" & Session("CMFPNBOT") &";border-radius: 100px;padding: 0px 4px;border: 1px solid "& Session("CMFPNBOT") &";text-indent: 0px !important;}"
s = s & ".aaa a .la-user:hover{background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 50%);opacity:1;}"
s = s & ".aaa a .la-user:active{background-color:" & Session("CMFPNBOT") & ";color:white;}"
s = s & ".aaa .la-history{font-size: 11pt;color:" & Session("CMFPNBOT") &";border-radius: 100px;padding: 0px 4px;border: 1px solid "& Session("CMFPNBOT") &";text-indent: 0px !important;}"
s = s & ".aaa .la-history:hover{background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 50%);opacity:1;}"
s = s & ".aaa .la-history:active{background-color:" & Session("CMFPNBOT") & ";color:white;}"


'Cards Noved
s = s & "div.grid-item {border: 1px solid" & Session("CMIETT") & ";background: white;border-radius: 6px;position: absolute;display: flex;align-items: center;justify-content: center;font-weight: bold;box-shadow: 0 2px 5px rgba(0,0,0,0.2);}"
s = s & "div.header {color:" & Session("CMIETTFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;" & "background-color:" & Session("CMIETT") & ";text-align:center;padding: 12px 16px;font-weight:700;border-radius: 5px 5px 0px 0px;margin-bottom: 5px;border-bottom: 1px solid rgba(0, 0, 0, 0.05);font-family:" & Session("CMFPNTITFTFC") & "}"
's = s & "div.contenido {font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;" & "background-color:" & Session("CMFPNTIT") & ";padding: 12px 16px;font-weight:600;border-radius: 5px 5px 0px 0px;margin-bottom: 5px;border-bottom: 1px solid rgba(0, 0, 0, 0.05);}"
s = s & "div.noved-doc-body{flex:1;display:flex;flex-direction:column;gap:4px;min-width:0;cursor: pointer;}"
s = s & "div.noved-doc-icon {color:" & Session("CMIETTFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;" & "background-color:" & Session("CMIETT") & ";border-radius:7px;padding:10px;}"
s = s & "div.noved-doc-head{display:flex;align-items:center;gap:8px;min-width:0;}"
s = s & "span.noved-doc-info {color:" & Session("CMIETT") & ";position:relative;cursor:default;font-size:14pt;flex-shrink:0;}"
s = s & "span.noved-doc-info::before{content:'';position:absolute;right:8px;top:115%;border:6px solid transparent;border-bottom-color:#111827;opacity:0;transition:.15s ease;}"
s = s & "span.noved-doc-info::after{content:attr(data-info);position:absolute;right:0;top:130%;background:#111827;color:#fff;font-size:0.72rem;line-height:1.35;padding:6px 8px;border-radius:6px;width:180px;opacity:0;pointer-events:none;transform:translateY(4px);transition:.15s ease;box-shadow:0 8px 20px rgba(0,0,0,.25);z-index:10;}"
s = s & ".noved-doc-info:hover::after, .noved-doc-info:hover::before{opacity:1;transform:translateY(0);color:" & Session("CMFPNTITFT") & "}"
s = s & "div.noved-doc-meta{display:flex;gap:8px;flex-wrap:wrap;align-items:center;}"
s = s & ".noved-doc-tag{color:" & Session("CMIETT") & ";background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);display:inline-flex;align-items:center;font-weight:700;padding:2px 8px;border-radius:999px;font-size:0.70rem;border-style: solid;border-width: 1px;width: fit-content;margin-left: 5px;}"
s = s & ".lista-docs{margin-bottom:8px;}"
s = s & ".documento-cards {display: flex;padding: 10px 7px;xmargin-bottom: 8px;border: 1px solid #ddd;/* l?nea gris por defecto */border-radius: 8px;background-color: #f9f9f9;box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);font-weight: 500;text-decoration: none;font-size: 12pt;transition: all 0.2s ease-in-out;align-items: center;flex: 1;cursor: pointer;gap: 5px;}"
s = s & ".documento-cards:hover {background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);border-color: " & Session("CMIETT") & ";text-decoration: none;transform: translateY(-2px);}"
s = s & ".documento-cards:hover ~ .documento-cards-footer {transform: translateY(-2px);}"
s = s & ".documento-cards:active {background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 70%);}"
s = s & ".documento-cards-footer {position:relative;z-index:-1;transition: all 0.2s ease-in-out;padding:10px 7px;font-weight:500;font-size: 12pt;align-items: center;border: 1px solid #ddd;background: color-mix(in srgb, #132574, transparent 90%);COLOR: #132574;border-radius: 0px 0px 8px 8px;}"
s = s & ".notas-card {display:flex;font-size:7pt;align-items:center;}"
s = s & "div.noved-doc-title{white-space:nowrap;overflow:hidden;text-overflow:ellipsis;flex:1;}"
s = s & "span.swiper-pagination-bullet-active{background-color: " & Session("CMFPNBOT") & " !important;}"
s = s & ".swiper-pagination-bullet {border: 1px solid " & Session("CMFPNBOT") & ";}"
s = s & ".icono-contenedor {position: relative;display: inline-block;padding: 10px;cursor: pointer;}"
s = s & ".burbuja {position: absolute;top: 6px;right: -4px;background: #c10218;color: #fff;width: 18px;height: 18px;border-radius: 50%;font-size: 10px;font-weight: 600;display: flex;align-items: center;justify-content: center;padding: 1px;border: 1px solid white;}"
s = s & ".toggle-drag-button {position: fixed; top: 11px; right: 16px; z-index: 9999999;}"
s = s & ".toggle-drag-button input {width: 44px; height: 24px; appearance: none; background: #cbd5e1; border-radius: 999px; position: relative; cursor: pointer; outline: none; transition: background 0.25s ease, box-shadow 0.25s ease; border: 2px solid #e2e8f0;}"
s = s & ".toggle-drag-button input::before {content: ''; position: absolute; top: 2px;left: 2px;width: 18px; height: 18px; background: #fff; border-radius: 50%; transition: transform 0.25s ease;}"
s = s & ".toggle-drag-button input:checked {background: #22c55e; border-color: #16a34a;}"
s = s & ".toggle-drag-button input:checked::before {transform: translateX(20px);}"
s = s & ".toggle-drag-button input:hover {box-shadow: 0 0 0 4px #22c55e40;}"
s = s & ".nvermas {color:" & Session("CMFPNBOTFT") & ";background-color:" & Session("CMFPNBOT") & ";width:fit-content;font-weight:bold;padding: 5px 10px;border-radius: 6px;margin-top: 1rem;font-size:1rem;cursor:pointer;transition: all 0.2s ease-in-out;}"
s = s & ".nvermas:hover {filter:brightness(1.15);}"
s = s & ".nvermas:active {filter:brightness(1.3);}"

'Publicacion de Documento con Multimedia
s = s & ".mostrarnovedad {font-family: SEGOE UI !important;color: #ffffff !important;font-size: 8pt !important;background-color: #7995ff !important;padding: 4px 8px !important;border-radius: 15px !important;border: none !important;cursor: pointer !important;margin-left:auto;margin-right:5px;}"
s = s & ".documento-cards:hover ~ .publi-card {transform: translateY(-2px);}"
s = s & ".publi-card {font-family:" & Session("CMFPNTITFTFC") & ";background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 70%);position:relative;z-index:-1;transition: all 0.2s ease-in-out;display:flex;flex-wrap:wrap;gap:5px;}"
s = s & ".publi-imagenovideo {flex:1 1 100px;max-width:100%;}"
s = s & ".publi-imagenovideo img {width:100%;height:100%;}"
s = s & ".publi-contenido {flex:2 1 200px;display:flex;flex-direction:column;gap:5px;justify-content:center;padding:10px}"
s = s & ".publi-titulo {font-size:0.7rem;color:" & Session("CMIETT") & ";font-weight:800;text-transform:uppercase;}"
s = s & ".publi-cuerpo {font-size:0.7rem;color:" & Session("CMIETT") & ";line-height:1.5;margin:0;word-wrap:break-word;}"

'POP-UP PUBLICACIONES
s = s & ".noticia-container {font-family:" & Session("CMFPNTITFTFC") & ";background: color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);max-height: 32rem;overflow-y:auto;overflow-x:hidden;border-radius: 1rem;position: relative;animation: slideUp 0.3s ease-out;padding: 1.5rem;}"
s = s & ".fechaycerrar {display: flex;flex-direction: row-reverse;}"
s = s & ".fechapubli {font-size: 13px;color: black;font-weight: bold;text-transform: uppercase;letter-spacing: 1px;}"
s = s & ".titulopubli {font-weight: bold;font-size: 20pt;line-height: 1.2;color:black;margin: 0;margin-bottom: 0.5rem;}"
s = s & ".copetepubli {font-size: 12pt;line-height: 1.5;color:#333333;margin-bottom: 1rem;}"
s = s & ".imagenovideo {position:relative;width: 100%;xbackground-color: color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);border-radius:5px;margin-bottom:1rem;}"
s = s & ".cuerpopubli {background: white;padding: 5px;border-radius: 5px;}"
s = s & ".footerpubli {xbackground: color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);border-radius:5px;font-weight:bold;padding:10px;font-size: 13px;color:black;margin-top: 1rem;}"

s = s & ".pill-video {padding: 0px 5px;background: darkred;width: fit-content;border-radius: 20px;border: 1px solid red;margin-bottom: 0 !important;font-size: 10pt;}"
s = s & ".fecvigen{background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 50%);font-size:10pt;font-weight:100;position: absolute;top: 15px;right: 20px;padding: 3px;border-radius: 5px;}"

'Ruta compatida susc y noved
s = s & "span.noved-doc-path{font-size:0.72rem;color:#6b7280;width:50%;xdisplay:flex;xgap:6px;align-items:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}"
s = s & "span.separator {color: #6b7280;font-size: 0.7rem;}"

'Cards autorizaciones
s = s & "div.authb-item{ background:var(--card); border:1px solid #ddd; border-radius:var(--radius); box-shadow:var(--shadow); position:relative; cursor:pointer; transition:all 0.2s ease-in-out; overflow:hidden;}"
s = s & "div.authb-item::before{background:" & Session("CMIETT") & ";content:'';position:absolute;left:0;top:0;bottom:0;width:6px;}"
s = s & "div.authb-item.is-selected{border-color:rgba(30,58,138,.45);box-shadow:0 14px 36px rgba(30,58,138,.14), var(--shadow);}"
s = s & "div.authb-list{display:flex;flex-direction:column;gap:10px;}"
s = s & "div.authb-item:hover{transform:translateY(-2px);background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);}"
s = s & "div.authb-item:active {background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 70%);}"
s = s & "div.authb-grid{display:grid;gap:10px;}"
s = s & "div.authb-wrap{margin:0px auto;padding:0px auto;}"
's = s & "@media(max-width:960px){.authb-grid{grid-template-columns:1fr;}}"
s = s & "div.authb-item-inner{padding:12px 14px 12px 18px;display:grid;grid-template-columns:1fr auto;gap:16px;}"
s = s & "div.authb-main{display:flex;flex-direction:column;gap:6px;min-width:0;}"
s = s & "div.authb-sub{font-size:9pt;color:var(--muted);padding-left: 5px;}"
s = s & "div.authb-tag{display:inline-flex;align-items:center;gap:6px;font-size:12px;border:1px solid var(--border);background:#f8fafc;padding:6px 10px;border-radius:999px;width:fit-content;}"
s = s & "div.authb-meta{display:flex;flex-direction:column;gap:6px;align-items:flex-end;min-width:170px;}"
's = s & "div.authb-preview{background:#fff;border:1px solid var(--border);border-radius:var(--radius);box-shadow:var(--shadow);padding:14px;position:sticky;top:16px;}"
's = s & "div.authb-preview-sub{font-size:12.5px;color:var(--muted);margin-bottom:10px;display:flex;gap:6px;align-items:center;}"

'Pill compatida aut y susc
s = s & "span.authb-pill{color:" & Session("CMFPNGNFT") & ";font-size:0.70rem;background:#ffffff;padding:4px 10px;border-radius:999px;display:flex;gap:6px;align-items:center;white-space:nowrap;width: fit-content;}"

'Header general de susc/doc a vencer/aut
s = s & ".ui-widget-header {font-family:" & Session("CMFPNTITFTFC") & " !important;background-color:" & Session("CMIETT") & " !important;color:" & Session("CMIETTFT") & " !important;font-weight: bold !important;border-radius: 5px !important;border: none !important;}"

'Cards suscripciones
s = s & "div.susc-item{background: white;border-radius: 10px;padding: 20px;box-shadow: 0 2px 5px rgba(0, 0, 0, 0.05);border: 1px solid #e0e0e0;display: flex;justify-content: space-between;align-items: center;gap: 20px;margin-bottom: 15px;position: relative;cursor: pointer;transition: all 0.2s ease-in-out;overflow: hidden;}"
s = s & "div.susc-item:hover{transform:translateY(-2px);background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);}"
s = s & "div.susc-item:active {background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 70%);}"
s = s & "div.susc-item::before{background:" & Session("CMIETT") & ";content:'';position:absolute;left:0; top:0; bottom:0;width:6px;}"

'Cards doc a vencer
s = s & "li.documento-vencer {margin-bottom:8px;display: flex;padding: 10px 7px;border: 1px solid #ddd;border-radius: 8px;background-color: #F9F9F9;box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);color: #333;font-weight: 500;text-decoration: none;font-size: 12pt;transition: all 0.2s ease-in-out;align-items: center;flex: 1;cursor: pointer;gap:5px;}"
s = s & "span.noved-doc-tag-vencer{color:#C7BF2E;background-color: rgba(235, 226, 61, 0.1);display:inline-flex;align-items:center;transition: all 0.2s ease-in-out;font-weight:700;padding:2px 8px;border-radius:999px;font-size:0.70rem;border-style: solid;border-width: 1px;width: fit-content;margin-left: 5px;}"
s = s & "li.documento-vencer:hover{background-color: #ebe23d1a !important;transform: translateY(-2px);}"
s = s & "li.documento-vencer:active{background-color: #faf8d4 !important;}"
s = s & "div.noved-doc-icon-vencer{color:" & Session("CMIETTFT") & ";background-color:#C7BF2E;font-size:14pt;transition: all 0.2s ease-in-out;border-radius:7px;padding:10px;}"

'Mensaje sin informacion
s = s & ".Sininfo {font-family:" & Session("CMFPNTITFTFC") & ";font-weight: bold;font-size: 1rem;}"

'iconos

	  s = s & "div.barrafs { display: flex; align-items: center; gap: 14px; margin-right: 24px; cursor: pointer; color:" & Session("CMFMNFT") & ";transition: color 0.2s ease;}"
	  s = s & "div.dinicio { cursor: pointer; color:" & Session("CMFMNFT") & ";transition: color 0.2s ease;}"
	  s = s & "a.burguer { cursor: pointer; color:" & Session("CMFMNFT") & ";transition: color 0.2s ease;margin-left:15px;}"
	  s = s & "type-chip.i { font-size: " & Session("CMFPNIPFTSZ") & "pt;}"
}



'''''''''indice
s = s & "div.footer{padding: 6px 12px;font-size: 0.7rem;color: #444;text-align: right;background: transparent;border-top: 1px solid rgba(0, 0, 0, 0.05);}"
s = s & "div.contenedor {font-family:" & Session("CMIETTFTFC") & ";height: 100%;width: 100%;background-size: cover;border-radius: 5px;display: flex;flex-direction: column;overflow: hidden;transition: all 0.3s ease;background-color: FFFFFF;box-sizing: border-box;}"
s = s & "div.drag-handle {background: rgba(51, 51, 51, 0.5); backdrop-filter: blur(5px); color: white; padding: 5px;cursor: move;text-align: center;font-size: 12px;margin-bottom: 10px;display:none;position: absolute;z-index: 20;width:100%;box-shadow: 0 2px 5px rgba(0,0,0,0.2);}"
s = s & "div.grid-canvas {margin: 0 auto; min-height: 100vh; width: 100%; height: 200%; position: relative; padding: 0;}"
s = s & "div.grid-canvasON { background: #e0e0e0; background-image: linear-gradient(to right, #ccc 1px, transparent 1px),linear-gradient(to bottom, #ccc 1px, transparent 1px); background-size: 20px 20px; margin: 0 auto; min-height: 100vh; width: 100%; height: 200%; position: relative; padding: 0;}"
s = s & "div.contenido {flex: 1;  padding: 0px 4px; overflow-y: auto;overflow-x: hidden;display: flex;flex-direction: column;gap: 10px;transform: translateY(-7px);margin-top:10px;}"
s = s & "div.attachments .chip {display: inline-flex; align-items: center; gap: 8px; max-width: 100%; padding: 6px 10px; border-radius: 8px; background: #f9fafb;   border: 1px solid #e5e7eb; font-size: 12.5px; color:" & Session("CMIETT") & "; text-decoration: none; line-height: 1; transition: border-color .15s ease, box-shadow .15s ease;}"
''s = s & "div.title {font-weight: 700; font-size: 16px; margin: 0; display: flex; gap: 8px;width: 80%; align-items: center;color:" & Session("CMFPNTIT") & ";}"
s = s & ".attachments .tag {font-size: 10px; padding: 2px 6px; border-radius: 6px; background: #eef2ff;color:" & Session("CMFPNTIT") & ";}"
s = s & ".em-filtro-all {background: var(--azul);color: #fff;font-weight: 600;font-size: 13px;border: none;border-radius: 8px;padding: 8px 14px;cursor: pointer;transition: .2s;}"
's = s & "div.title { display: flex; min-width: 0; flex: 1;line-height: 1.4; transition: all 0.2s ease in out; padding: 6px 10px; overflow: hidden; text-overflow: ellipsis;white-space: nowrap; }"
s = s & "#sidebarResults .em-title {font-weight: 700; font-size: 18px;line-height: 1.25; margin: 0; color:" & Session("CMFPNTIT") & "; letter-spacing: .2px; -webkit-font-smoothing: antialiased; text-rendering: optimizeLegibility;font-family:" & Session("CMFPNTITFTFC") & "}"
s = s & "div.info-row .label {color:" & Session("CMIETT") & "; font-weight: 700; min-width: 72px;}"
s = s & ".contg {height: 200px;overflow-x: hidden;margin-bottom:20px}"
s = s & ".em-facet-title {font-size: " & Session("CMFPNIPFTSZ") & "pt; color:" & Session("CMIETT") & "; font-weight: 600; margin: 0; letter-spacing: .2px;font-family:" & Session("CMFPNTITFTFC") & "}"
s = s & "#sidebarResults .em-count {color:" & Session("CMIETT") & ";font-size: " & Session("CMFPNIPFTSZ") & "pt; margin-top: 4px;}"
s = s & ".em-caret {display: inline-block;transition: transform .2s ease;}"
s = s & ".em-chk {display: flex;align-items: center;gap: 10px;font-size: 14px;text-overflow: ellipsis;white-space: nowrap;cursor: hand;}"
s = s & ".em-chip {display: inline-flex;align-items: center;gap: 6px;padding: 4px 8px;border-radius: 12px;font-size: 12px;font-weight: 500;color:" & Session("CMIETT") & ";background: #fff;border: 1px solid #e6e9ef;}" '''color: #334155
s = s & ".em-chips-label {font-weight: 600;color: var(--muted);white-space: nowrap;margin-right: 4px;font-size: 12px;}"
s = s & "span.excerpt {display: inline;}"
s = s & "a.more { color: var(--muted); cursor: pointer; text-decoration: none; margin-left: 2px;}"
s = s & "i {transition: all 0.2s ease-in-out;}"
s = s & ".em-clear-all {margin-left: auto;background: transparent;color: #64748b;border: none;padding: 0;font-size: 12px;font-weight: 600;cursor: pointer;text-decoration: underline;text-underline-offset: 2px;}"
s = s & "div.em-facet {border-top: 1px dashed #e5e7ee;padding-top: 12px;margin-top: 12px;}"
s = s & "div.em-facet:first-of-type {border-top: none;padding-top: 0;margin-top: 0;}"
s = s & "div.em-facet-head {display: flex;align-items: center;justify-content: space-between;margin: 0 0 10px;gap: 8px;}"
s = s & ".em-facet-toggle {border: 1px solid #d6d6d6;background: #fff;color: #334155;border-radius: 8px;height: 28px;padding: 0 8px;cursor: pointer;font-weight: 700;display: inline-flex;align-items: center;gap: 6px;}"
s = s & ".em-side {font-family: " & Session("CMIETTFTFC") & ";background: #fff;border: 1px solid var(--borde);border-radius: 12px; padding: 16px; box-shadow: 0 4px 12px rgba(0, 0, 0, .06); top: 12px; height: auto; overflow: auto; transition: width .25s ease, padding .25s ease;}"



'''''''consultas por tema
s = s & "div.footercons {display:flex;justify-content:center;gap:10px;background-color:" & Session("CMFPNTIT") & ";padding: 7px 0px;margin:0 90px;border-radius: 0px 0px 5px 5px;}"
s = s & "font.paginador {font-family:" & Session("CMFPNTITFTFC") & ";font-size: 12px;font-weight: 500;color: " & Session("CMIETTFT") & ";padding: 2px 6px;border-radius: 5px;text-align: center;cursor: default;}"
s = s & "div.contenedorcons {xheight: 100%;background-size: cover;border-radius: 5px 5px 0px 0px;display: flex;flex-direction: column;overflow: hidden;background-color: FFFFFF;box-sizing: border-box;margin: 0px 90px;border: 1px solid " & Session("CMFPNTIT") & ";}"
s = s & "div.headertema {color:" & Session("CMIETTFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;" & "background-color:" & Session("CMFPNTIT") & ";text-align:center;padding: 12px 16px;font-weight:600;border-radius: 5px 5px 0px 0px;margin-bottom: 5px;border-bottom: 1px solid rgba(0, 0, 0, 0.05);font-family:" & Session("CMFPNTITFTFC") & "}"
s = s & "div.headertema2 {color:" & Session("CMIETTFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;" & "background-color:" & Session("CMFPNTIT") & ";padding: 12px 16px;font-weight:600;border-radius: 5px 5px 0px 0px;margin-bottom: 5px;border-bottom: 1px solid rgba(0, 0, 0, 0.05);font-family:" & Session("CMFPNTITFTFC") & "}"
s = s & "div.headerpath {color:" & Session("CMIETTFT") & ";font-family:" & Session("CMFPNIPFTFC") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;" & "background-color:" & Session("CMFPNTIT") & ";display: flex;align-items: center;gap: 5px;padding: 10px 16px;font-weight: 600;border-bottom: 1px solid rgba(0, 0, 0, 0.05);border-radius: 5px 5px 0 0;box-shadow: 0 2px 4px rgba(0, 0, 0, 0.08);width: 100%;}"
's = s & ".header a:hover{text-decoration: underline;}"
s = s & ".headerpath:last-child:hover {text-decoration: none;cursor: default;}"
s = s & ".dialog-form-header-filtrado{background-color: " & Session("CMFPNTIT") & ";}"
s = s & ".ui-dialog .ui-dialog-title:{background-color: " & Session("CMFPNTIT") & ";}"
s = s & "div.title-row-tema {pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNGNFT") & ";display: flex;align-items: center;gap: 8px;min-width: 0;width: 100%;}"
s = s & "div.idxv-item {display: flex;align-items: center;justify-content: space-between;border: 1px solid #eef2f7;padding: 12px;transition: all 0.2s ease;border-radius: 5px;}"
s = s & "div.idxv-item:hover {border-color:" & Session("CMFPNTIT") & "}"
s = s & "span.type-chip {display: inline-flex;align-items: center;gap: 8px;border-radius: 5px;padding: 6px 12px;border: 1px solid " & Session("CMFPNTIT") & ";background: #f6f6f6;transition: all 0.2s ease-in-out;}"
s = s & "span.type-chip:hover {background-color: color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 30%);}"
s = s & "span.chip-name {white-space: nowrap;overflow: hidden;text-overflow: ellipsis;}"
s = s & "button.child-inline {border: 1px solid " & Session("CMFPNTIT") & ";background: #f6f6f6;color: " & Session("CMFPNTIT") & ";padding: 3px 8px;border-radius: 999px;font-size: .78rem;cursor: pointer;line-height: 1;transition: all 0.2s ease-in-out;}"
s = s & "button.child-inline:hover {background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 30%);}"
s = s & "button.child-inline:active {background-color:" & Session("CMFPNBOT") & "color:" & Session("CMFPNBOTFT") & ";}"
s = s & "button.icon-btn {background-color:" & Session("CMFPNTIT") & ";border: none;padding: 6px;display: inline-flex;align-items: center;border-radius: 5px;}"
s = s & ".icon-btn i {color:" & Session("CMIETTFT") & ";font-size: 1rem;}"
s = s & "div.right-ctrl{gap: 5px;display: flex;flex-direction: column;}"
's = s & ".contenedorcons {height: 100%;background-size: cover; border-radius: 5px;font-family:" & Session("CMFPNTITFTFC") & ";display: flex;flex-direction: column; overflow: hidden; transition: all 0.3s ease;background-color: FFFFFF; box-sizing: border-box; margin: 0px 90px; border: 1px solid #ff7f07;}"

''''''''administracion de novedades
''s = s & ""

''''''''consulta / grafico
s = s & ".date-row {display: flex; gap: 20px;flex-wrap: wrap; }"
s = s & "input.form-group {  margin-bottom: 1rem; }"
's = s & "input.form-group label {display: block; font-weight: bold; margin-bottom: 0.5rem;}"
s = s & "div.norma-card2 {max-width: 960px;margin: 30 auto;border-radius: 10px;padding: 10px 40px;line-height: 1.7;background-color: " & Session("CMIETT") & ";color: #ffffff;font-family:" & Session("CMFPNTITFTFC") & "}"
s = s & "#graficar {width: 100%;padding: 12px 20px;font-size: 11pt;font-weight: bold;color: #fff;background-color:" & Session("CMFPNBOT") & ";border: none;border-radius: 25px;cursor: pointer;transition: background-color 0.2s;font-family:" & Session("CMFPNBOTFTFC") & ";}"
s = s & "#graficar:hover {opacity: 0.9;}"
s = s & ".chart-container {position: relative; height: 60vh; width: 100%; margin: auto;}"
s = s & ".form-group input[type='date'],"
s = s & ".form-group input[type='text'],"
s = s & ".form-group select {width: 100%;padding: 0.5rem;font-size: 9pt;border-radius: 20px;cursor: pointer;}"
s = s & ".form-group { margin-bottom: 1rem;}"

'contenedor del

'contenedor del e




     	  s = s & "a.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;text-decoration:none;}"
         s = s & "a.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;text-decoration:none;}"
         s = s & "a.cc{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";text-decoration:none;}"
         s = s & "a.tt{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFPNTITFT") & ";transition:all 0.2s ease-in-out;font-weight:800;text-decoration:none;}"
		 s = s & "a.tt:hover{text-decoration:none;opacity:0.8;}"
		 s = s & "a.tt:active{text-decoration:none;opacity:0.6;}"
		 s = s & "a.ts{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";;font-weight:800;text-decoration:none;}"
         s = s & "a.cmb{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-weight:normal;text-decoration:none;}"

         s = s & "a.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:500;text-decoration:none;}"

         s = s & "font.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"

         s = s & "font.bt2{font-family:Corflaticon;color:#ffffff;font-size:10pt;border-radius:3px;font-weight:bold;background-color:#dd0000;}"
         s = s & "font.bt1{font-family:Corflaticon;color:#ffffff;font-size:10pt;border-radius:3px;font-weight:bold;background-color:#00bb00;}"

          s = s & "font.btc1{font-family:Corflaticon;color:white;font-size:8pt;font-weight:bold;display: flex}"


'''timeline
s = s & "div.right-panel {font-family:" & Session("CMIETTFTFC") & ";flex:1;border-left: 1px solid #e5e7eb;padding-left: 24px;}"
s = s & "div.timeline {position: relative;margin-left: 20px;padding-left: 20px;border-left: 3px solid #d1d5db;}"
s = s & "div.timeline-title{font-size: " & Session("CMFPNIPFTSZ") & "pt;font-weight: 700;margin-bottom: 16px;border-bottom: 1px solid #d1d5db; padding-bottom: 4px;color:" & Session("CMIETT") & ";}"
s = s & ".timeline-event { position: relative; margin-bottom: 28px; padding-left: 16px;}"
s = s & ".timeline-event::before {content: ''; position: absolute; left: -32px; top: -2px; width: 16px; height: 16px; background-color:" & Session("CMIETT") & "; border-radius: 50%; border: 3px solid #fff; box-shadow: 0 0 0 2px " & Session("CMIETT") & ";}"
s = s & ".timeline-event:last-child { margin-bottom: 0;}"
s = s & ".info-item {background: #f9fafb;padding: 5px 10px;border-radius: 6px;border: 1px solid #e5e7eb;width: 95%;box-sizing: border-box;}"
s = s & ".breadcrumb {font-size:11pt;margin-bottom:10px;}"
s = s & ".info-grid {display: flex;flex-direction: column;gap: 10px;margin-bottom: 24px;}"
s = s & ".event-date { font-size: " & Session("CMFPNIPFTSZ") & "pt; font-weight: 500;color:" & Session("CMIETT") & ";}"
s = s & ".event-label {font-size: 0.9rem;color: #6b7280;margin-bottom: 4px;font-weight: 600;}"

    ' botones elim
         's = s & "input{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("aplbgd") & ";}"
		 s = s & "input{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background:none;}"
         s = s & "input.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:#ffffff;border-radius:15px;}"

		 's = s & "input.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";" & "color:" & Session("CMFPNBOTFT") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "background-color:" & Session("CMFPNBOT") & ";" & "padding: 4px 8px;" & "border-radius: 8px;" & "border: none;" & "cursor: pointer;" & "transition: background-color 0.3s, transform 0.2s;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"

         s = s & "input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#8a8a8a",session("CMFPNBOT"),0) & ";transition: all 0.2s ease-in-out;border-style:solid;border-width:1;border-color:#8a8a8a;height:30px;padding: 0px 10px; border-radius: 15px; border: none; cursor: pointer;}"
         s = s & "input.bt:hover{filter: brightness(1.15);}"
         s = s & "input.bt:active{filter: brightness(1.3);}"
         ' s = s & "input.bte{font-family:'Line Awesome Free';font-weight:900;margin-left:5px;font-size:11pt;color:red;padding:0px 4px;border:1px solid red;cursor:pointer;border-radius:100px;text-indent:0px;}"
		 ' s = s & "input.bte:hover{background-color:color-mix(in srgb, red, transparent 50%);opacity:1;}"
		 ' s = s & "input.bte:active{background-color:red;color:white;}"
		 ' Xinput7 = "&#xf00d;" (opcion X)

		 s = s & "input.bte{font-family:'Line Awesome Free';font-weight:900;margin-left:1px;border:none;font-size:14pt;color:darkgray;cursor:pointer;text-indent:0px;}"
		 s = s & "input.bte:hover{color:color-mix(in srgb, red, transparent 50%);opacity:1;}"
		 s = s & "input.bte:active{color:red;}"

		 s = s & "i.bte{font-family:'Line Awesome Free';font-style:normal;margin-left:5px;font-size:11pt;color:red;padding:0px 4px;border:1px solid red;cursor:pointer;border-radius:100px;text-indent:0px;}"
		 s = s & "i.bte:before{content:'\f00d';font-family: 'Line Awesome Free';font-weight: 900;}"
		 s = s & "i.bte:hover{background-color:color-mix(in srgb, red, transparent 50%);opacity:1;}"
		 s = s & "i.bte:active{background-color:red;color:white;}"
         s = s & "input.btG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#eeeeee;height:20px;cursor:pointer;}"
         s = s & "input.btR{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#550000",0) & ";border-style:solid;border-width:1;border-color:#550000;height:20px;cursor:pointer;}"


         s = s & "input.dtb{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dtn{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-radius:10px;text-align=right;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
		 s = s & "input.dtn:focus {border-color: " & Session("CMFPNBOT") & ";box-shadow: 0 0 0 2px " & Session("CMFPNBOT") & "40;transition: all 0.2s ease-in-out;}"
         s = s & "input.dts{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
         s = s & "input.dtbs{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;padding-left:5px;border-radius:10px;font-weight:bold;border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
         s = s & "input.dtns{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"


         s = s & "input.it{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";border-radius:15px;}"
         s = s & "input.rd{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"
         s = s & "input.tt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;border-radius:15px;cursor:pointer;font-weight:600;background-color:" & session("CMFPNBOT") & ";}"
		 s = s & "input.tt:hover{filter: brightness(1.15);}"
		 s = s & "input.tt:active{filter: brightness(1.3);}"
         s = s & "input.tt3{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;border-radius:15px;cursor:pointer;font-weight:600;background-color:gray;}"
		 s = s & "input.tt3:hover{filter: brightness(1.15);}"
		 s = s & "input.tt3:active{filter: brightness(1.3);}"

         s = s & "option{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"
         s = s & "option.OCM{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"

         s = s & "select{width:100%;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("APLBGD") & ";border-style:solid;border-radius:15px;border-width:1;cursor:pointer;border-color:" & Session("CMFPNTIT") & ";}"
         s = s & "select.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";}"

         s = s & "span.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"

   '------------------
         s = s & "span.XMN{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";background-color:"&SESSION("CMICTT")&";border-style:solid;border-width:1pt;border-color:#000000;color:"&session("CMICTTFT")&";}"

         s = s & "span.XSL{font-size:" & Session("CMFMEFTSZ") & "pt;font-family:" & Session("CMFMEFTFC") & ";background-color:"&SESSION("CMIBTT")&";border-style:solid;border-width:1pt;border-color:#FFFFFF;color:"&session("CMIBTTFT")&";}"

         s = s & "span.ntslp{font-size:" & Session("CMFMEFTSZ") & "pt;font-family:" & Session("CMFMEFTFC") & ";color:" &  Session("CMFMEFT")& ";font-weight:500;}"

         s = s & "span.ntmnn{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & session("CMFMNFT") & ";font-weight:500;}"
         's = s & "a.ntmnn{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & session("CMFMNFT") & ";font-weight:500;text-decoration:none;}"

   '-----------------------------
         s = s & "span.mt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;;font-weight:600;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNIP") & ";width:80%;}"
		 s = s & "span.sdelim{font-size: 15pt;}"

         s = s & "th{border-radius: 20px; font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

         s = s & "tr.st{background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "tr.cc{background-color:" & Session("CMFPNIT") & ";border-style:ridge;border-width:1pt;}"
         s = s & "tr.XFK{height:20px;}"
         s = s & "td.round{"
		 s = s & "    border-collapse: collapse;"
		 s = s & "    border-style: hidden;"
		 s = s & "    box-shadow: 0 0 0 1px #000; "
        s = s & "border-radius: 4px;}"

         s = s & "td.aa{text-transform: capitalize;border-radius: 5px;}"
		 s = s & "td.aaa{border-radius: 5px;}"
         s = s & "td{font-family:" & Session("CMFPNITFTFC") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;font-weight:500;}"
         s = s & "td.at{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;}"
         s = s & "td.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:" & session("CMFPNBOT") & ";text-decoration:underline;border-style:outset;border-width:1pt;}"
         s = s & "td.btG{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#dddddd;border-style:inset;border-width:1pt;}"
         s = s & "td.cc{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-radius:5px;text-align:center;border-style:solid;border-width:1;border-color:" & Session("CMFPNITFT") & ";}"
         s = s & "td.ee{font-size:" & Session("CMFPNITFTSZ") & "pt;font-family:" & Session("CMFPNITFTFC") & ";border-width:1pt;border-color:" & Session("CMFPNTIT") & ";color:" & Session("CMFPNITFT") & ";font-weight:500;}"
         s = s & "td.lt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;text-indent:10;}"
         s = s & "td.dt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.dt2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.et{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.mt{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;;font-weight:600;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNIT") & ";}"
         s = s & "td.qt{font-family:" & Session("CMFPNGRFTFC") & ";color:" & Session("CMFPNGRFT") & ";font-size:" & Session("CMFPNGRFTSZ") & "pt;font-weight:600;}"
         s = s & "td.rr{font-size:" & Session("CMFPNTITFTSZ") & "pt;font-family:" & Session("CMFPNTITFTFC") & ";border-width:1pt;border-color:" & Session("CMFMNBOT") & ";color:" & Session("CMFPNTITFT") & ";font-weight:500;}"
         s = s & "td.st{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

         s = s & "td.tdcm1{cursor:pointer;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-radius:10px;border-style:solid;border-width:1pt;border-color:#dddddd;padding:0 5px;background-color:#f0f0f0;min-width: 100px;}"
         s = s & "td.tdcm2{cursor:pointer;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-radius:10px;border-style:solid;background-color:#0000dd;transition: all 0.2s ease-in-out;}"
		 s = s & "td.tdcm2:hover{filter: brightness(1.15);}"
		 s = s & "td.tdcm2:active{filter: brightness(1.3);}"


         s = s & "td.btbs{color:" & Session("CMFPNTIT") & ";}"
         s = s & "td.btbn{color:" & Session("CMFPNTIT") & ";text-align:center;}"
         s = s & "td.btba{color:" & Session("CMFPNTIT") & ";cursor:pointer;}"

         s = s & "td.btbsr{background-color:#6699cc;color:#ffffff;border-style:solid;border-width:1;}"
         s = s & "td.btbnr{background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;}"
         s = s & "td.btbar{background-color:#005500;color:#ffffff;border-style:solid;border-width:1;}"

         s = s & "div.ut  {border-style:none;border-width:1;border-color:#ffffff;}"
         s = s & "div.btbs{}"
         s = s & "div.btbn{color:" & Session("CMFPNBOT") & ";}"
         s = s & "div.btba{}"

         s = s & "a.btbs{color:" & Session("CMFPNTIT") & ";text-decoration:none;}"
         s = s & "a.btbn{color:" & Session("CMFPNTIT") & ";text-decoration:none;}"
         s = s & "a.btba{color:" & Session("CMFPNTIT") & ";text-decoration:none;}"



          s = s & "font.btic{font-family:Corflaticon;font-size:10pt;line-height:12pt;font-weight:normal;text-decoration:none}"




        s = s & "font.btiw{font-family:Corflaticon;font-size:10pt;line-height:12pt;font-weight:normal;}"
         s = s & "font.btic2{font-family:" & Session("CMFPNIPFTFC") & ";font-size:9pt;line-height:12pt;font-weight:normal;}"


         s = s & "a.ft1n{font-family:Corflaticon;font-size:10pt;padding:2px;text-decoration:none;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";border-radius:15px;}"
         s = s & "a.ft1o{font-family:Corflaticon;font-size:10pt;padding:3px;text-decoration:none;background-color:" & Session("CMFPNBOT") & ";border-radius:15px;}"

         s = s & "a.ft2n{font-family:Corflaticon;font-size:10pt;line-height:12pt;padding:2px;color:" & Session("CMFPNBOT") & ";font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";border-radius:15px;}"
         s = s & "a.ft2o{font-family:Corflaticon;font-size:10pt;line-height:12pt;padding:2px;text-indent:7pt;color:#ffffff;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";background-color:" & Session("CMFPNBOT") & ";border-radius:15px;}"

         s = s & "a.ft3n{font-family:Corflaticon;font-size:10pt;line-height:12pt;border-radius:15px;padding:2px;color:" & Session("CMFPNBOT") & ";font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";}"
         s = s & "a.ft3o{font-family:Corflaticon;font-size:10pt;line-height:12pt;border-radius:15px;padding:2px;text-indent:5pt;color:#ffffff;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";background-color:" & Session("CMFPNBOT") & ";}"

         s = s & "textarea{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#F9F9F9;border-radius:5px;}"
         s = s & "textarea.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;width:100%;border-radius:5px;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#F9F9F9;}"

''' Estilos de Impresion......


         s = s & "div{page-break-after:always;}"
         s = s & "div.DINDM{background-color:" & Session("APLBGD") & ";border-style:SOLID;padding:3px;border-radius:5px;border-width:2;border-color:" & session("CMFMNBOT") & ";}"
         s = s & "div.DINDC{border-style:solid;border-width:1;border-color:#AAAAAA;background-color:#f9f9f9;}"


		 s = s & "td.iettprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;text-decoration:none;}"




		  s = s & "td.ietxprn{Xbackground-color:" & Session("CMFMNFT") & ";color:" & Session("CMFMENU") & ";font-family:" & Session("CMFMNFTFC") & ";border-radius:10px 0px 0px 10px;font-size:8pt;font-weight:bold;font-style:italic;text-decoration:none;}"

		  s = s & "td.isttprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"


		  s = s & "td.istxzprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"

		  s = s & "td.ibttprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.icttprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.ictxprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weifght:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.ictrprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.icxrprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"



''''Session("CMFPNBOTFTFC")

         'Botones de panel
	s = s & "#panelM { background-color: transparent !important; background-image: none !important;box-shadow:none !important;border: none;}"
        ' s = s & "input.bietx{font-weight:bold;border-radius:00px;cursor:pointer;font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;background-color:" & Session("CMFPNBOT") & ";xborder-style:solid;xborder-width:1;xborder-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
        ' s = s & "input.bietm{border:none white;background-color:#787878;font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:12pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
        ' s = s & "input.bietmG{border:none red;background-color:#787878;font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:12pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
        ' originals = s & "input.bietx{border-radius:10px;font-weight:bold;xborder-radius:10px;cursor:pointer;font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;background-color:" & Session("CMFPNBOT") & ";xborder-style:solid;xborder-width:1;xborder-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		' opcion 1  s = s & "input.bietx{" & "border-radius: 10px;" & "font-weight: bold;" & "cursor: pointer;" & "font-family:" & Session("CMFPNBOTFTFC") & ";" & "color:" & Session("CMFPNBOTFT") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "background-color:" & Session("CMFPNBOT") & ";" & "padding: 8px 16px;" & "border: none;" & "transition: background-color 0.3s, transform 0.2s;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		' opcion 1 s = s & "input.bietx:hover{" & "background-color:" & Session("CMFPNBOTHD") & ";" & "color:white;" & "transform: scale(1.05);" & "}"
		s = s & ".bietx{font-weight:bold !important;transition:all 0.2s ease-in-out;}"
		s = s & ".bietx:hover{filter: brightness(1.15);}"
		s = s & ".bietx:active{filter: brightness(1.3);}"
		s = s & ".bietxC:hover{filter: brightness(1.15);}"
		s = s & ".bietxC:active{filter: brightness(1.3);}"
		s = s & "input.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";" & "color:" & Session("CMFPNBOTFT") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "background-color:" & Session("CMFPNBOT") & " !important;" & "padding: 4px 8px;" & "border-radius: 15px;" & "border: none;" & "cursor: pointer;" & ";" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		s = s & "button.bietx{font-family:" & Session("CMFPNBOTFTFC") & " !important;" & "color:" & Session("CMFPNBOTFT") & " !important;" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt !important;" & "background-color:" & Session("CMFPNBOT") & " !important;" & "padding: 4px 8px !important;" & "border-radius: 15px !important;" & "border: none !important;" & "cursor: pointer !important;" & ";" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
        s = s & "button.bietxC{font-family:" & Session("CMFPNBOTFTFC") & " !important;" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt !important;color:#ffffff !important;background-color:#c10218 !important;padding:4px 8px !important;border-radius:15px !important;border:none !important;cursor:pointer;transition:all 0.3s ease-in-out;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		s = s & "button.bietxApr{font-family:" & Session("CMFPNBOTFTFC") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;color:#ffffff;background-color:08c102 !important;padding:4px 8px;border-radius:15px;border:none;cursor:pointer;transition:all 0.3s ease-in-out;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		s = s & "button.bietxApr:hover{opacity:0.6;}"
		s = s & "button.bietxApr:active{opacity:0.8;}"
		's = s & "input.bietx:hover{" & "background-color:white;" & "color:" & Session("CMFPNBOT") & "}"
		 's = s & "input.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "color:" & Session("CMFPNBOTFT") & ";" & "background-color:" & Session("CMFPNBOT") & ";" & "border:1px solid " & Session("CMFPNBOT") & ";" & "padding:10px 16px;" & "border-radius:24px;" & "cursor:pointer;" & "transition:background-color 0.3s, box-shadow 0.2s;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		 's = s & "input.bietx:hover{" & "background-color:" & Session("CMFPNBOTHD") & ";" & "color:white;" & "box-shadow:0 2px 8px rgba(0,0,0,0.15);" & "}"



		         s = s & "input.bietm{border:none white;background-color:#787878;font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:12pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		         s = s & "input.bietmG{border:none red;background-color:#787878;font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:12pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"


          'original s = s & "a.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		 s = s & "a.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";" & "color:" & Session("CMFPNBOTFT") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "background-color:" & Session("CMFPNBOT") & ";" & "padding: 4px 8px;" & "border-radius: 20px;" & "border: none;" & "cursor: pointer;" & "transition: all 0.2s ease-in-out" & ";" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		 's = s & "a.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "; padding: 10px 16px; margin-right:6px; display:inline-block; background-color:#f1f1f1; border-radius:5px 5px 0 0; text-decoration:none; transition:all 0.2s ease;}"
         's = s & "a.bietx:hover{background-color:#e0e0e0; color:#1e3a8a;}"
         's = s & "a.bietx.active{background-color:#ffffff; border-bottom:2px solid transparent; color:#1e3a8a;}"
		 s = s & "a.bietm{font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:10pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
         ' original s = s & "td.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;background-color:" & Session("CMFPNBOT") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		 s = s & "td.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";" & "color:" & Session("CMFPNBOTFT") & ";" & "font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & "background-color:" & Session("CMFPNBOT") & ";" & "padding: 4px 8px;" & "border-radius: 20px;" & "border: none;" & "cursor: pointer;" & "transition: all 0.3s ease-in-out;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
		 s = s & "td.bietx:hover{filter: brightness(1.15);}"
		 s = s & "td.bietx:active{filter: brightness(1.3);}"
		 s = s & "tr.bietm{font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:10pt;background-color:#787878;xborder-style:solid;xborder-width:1;xborder-color:#666666;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"

		 s = s & "td.bietm{font-family:" & Session("CMFPNBOTFTFC") & ";color:white;font-size:10pt;padding:10px;background-color:#787878;xborder-style:solid;xborder-width:1;xborder-color:#666666;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"

         s = s & "input.bietxG{border-radius:15px;font-family:" & Session("CMFPNBOTFTFC") & ";color:black;font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & ";" & "padding: 4px 8px;" & "pt;border-color:" & Session("CMFPNBOT") & ";opacity: 0.8;background-color:#e8e9eb;border-style:solid;border-width:2px;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"

		 ' INPUT BUSCADOR PANTALLA DE BUSQUEDA

         s = s & "input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;padding-left:5px;border-radius:10px;border-style:solid;border-width:1pt;border-color:#8a8a8a;background-color:#f0f0f0;}"
		 s = s & "input.dt:focus {border-color: " & Session("CMFPNBOT") & ";box-shadow: 0 0 0 2px " & Session("CMFPNBOT") & "40;transition: all 0.2s ease-in-out;}"

	     ' INPUT BUSCADOR TOPBAR

		 's = s & "input.dtx {font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:#f9f9f9;border: 1px solid #ccc;border-radius: 8px;padding: 12px 16px;width: 100%;max-width: 500px;box-shadow: inset 0 1px 2px rgba(0,0,0,0.1);transition: all 0.3s ease;outline: none;box-sizing: border-box;}"
s = s & "input.dtx{font-family:" & Session("CMFPNIPFTFC") & ";" & _
             "color:" & Session("CMFPNIPFT") & ";" & _
             "font-size:" & Session("CMFPNIPFTSZ") & "pt;" & _
             "background-color:#ffffff;" & _
             "border:1px solid " & Session("CMFPNBOT") & ";" & _
             "padding:10px 16px;" & _
             "width:400px;" & _
             "display:inline-block;" & _
             "vertical-align:middle;" & _
             "border-radius:28px 0 0 28px;" & _
             "box-shadow:0 1px 4px rgba(0,0,0,0.1);" & _
             "transition:all 0.3s ease;" & _
             "outline:none;" & _
			 "height: 34px;" & _
             "box-sizing:border-box;}"


		 'BOTONES DEL BUSCADOR

		 'BOTON MICROFONO

		 s = s & "button.inputmic{position:absolute; right:8px; top:53%; transform:translateY(-50%); background:none; border:none;border-left: 0.5px solid grey; cursor:pointer;height:100%;}"
		 s = s & "button.inputmic:hover{background-color:" & Session("CMFMENU") & ";color:" & Session("CMFMNFT") &";}"

         'BOTON DE BUSCAR

		s = s & "button.inputbutton{display:inline-block;" & _
					 "vertical-align:middle;" & _
					 "background-color:#ffffff;" & _
					 "border:1px solid " & Session("CMFPNBOT") & ";" & _
					 "cursor:pointer;" & _
					 "padding:8px;" & _
					 "height: 34px;" & _
					 "box-shadow:0 1px 4px rgba(0,0,0,0.1);" & _
					 "border-radius:0 28px 28px 0;" & _
					 "box-sizing:border-box;}"

		's = s & "button.inputbutton:hover{background-color:" & Session("CMFMENU") & ";color:" & Session("CMFMNFT") & ";}"

		s = s & "button.inputbutton:focus{border-color:#007BFF;box-shadow:0 0 6px rgba(0,123,255,0.3);}"





'----------------------------------------------------------------------------------------------------------------------

         'seccion ,indice normativo --- novedades
         s = s & "td.iett{ border-radius:10px; xbox-shadow: 4px 4px 12px rgba(0,0,0,0.4);font-family:" & Session("CMIETTFTFC") & ";color:" & Session("CMIETTFT") & ";font-size:" & Session("CMIETTFTSZ") & "pt;background-color:" & Session("CMIETT") & ";" & replace(Session("CMIETTBL"),"1","font-weight:bold;") & replace(Session("CMIETTIT"),"1","font-style:italic;") & replace(Session("CMIETTUL"),"1","text-decoration:underline;") & "}"

         'Publicaciones y recomendaciones y Mis autorizaciones y observaciones NOVEDADES

	 s = s & "li.documento-publi {display: flex;padding: 10px 5px;margin-bottom: 6px;border: 1px solid #ddd;border-left: 5px solid #ccc;border-radius: 8px;background-color: #f9f9f9;box-shadow: 0 2px 6px rgba(0, 0, 0, 0.05);color: #333;font-family: 'Segoe UI', sans-serif;font-weight: 500;text-decoration: none;font-size: 12pt;transition: all 0.2s ease-in-out;align-items: center;flex: 1;cursor: auto;}"
	 s = s & "li.documento-publi:hover {background-color: #E5F2E0;border-left-color: #52B927;text-decoration: none;transform: translateY(-3px);}"

         s = s & "input.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:" & Session("CMIETX") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"
         s = s & "a.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"

         s = s & "td.ietx{ box-shadow: 5px 5px #cccccc; width:20px;font-weight:bold;color:" & Session("CMIETXFT") &"; background-color:" & Session("CMIETX") & ";height:40px;text-align:center;border-radius:20px;xbox-shadow: 4px 4px 12px rgba(0,0,0,0.4);font-family:" & Session("CMIETXFTFC") & ";xcolor:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;xbackground-color:" & Session("CMIETX") & ";xborder-style:solid;xborder-width:1;xborder-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"

       ' s = s & "td.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:" & Session("CMIETX") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"

         s = s & "td.ietxBG{height:45px;background-color:" & SESSION("CMICTTFTULTCS") & ";font-family:" & Session("CMICTTFTFCULTCS") & ";color:" & Session("CMIBTTFTULTCS") & ";font-size:" & Session("CMICTTFTSZULTCS") & "pt;}"

         s = s & "input.ietxG{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:#bbbbbb;border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"
          s = s & " option.OCMB { width:300px }"


 'Publicaciones y recomendaciones y Mis autorizaciones y observaciones NOVEDADES
         'TITULO DEL DOCUMENTO
         s = s & "td.istt{ border-radius:10px;xbox-shadow: 4px 4px 12px rgba(0,0,0,0.4);font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:" & Session("CMISTTFTSZ") & "pt;background-color:" & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.isttBF{ xbox-shadow: 4px 4px 12px rgba(0,0,0,0.4);font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:26px;background-color:" & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & ";xborder-style:solid;xborder-size:2px;border-color:red;Xborder-radius:10px;}"
         s = s & "div.isttBF{color:black;font-family:" & Session("CMISTTFTFC") & ";xcolor:" & Session("CMISTTFT") & ";font-size:26px;" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & ";}"
         s = s & "div.isttBF2{font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:24px;background-color:" & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & ";border-style:solid;border-size:2px;border-color:red;}"
	 's = s & "td.header.rh{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt; padding: 12px 16px; color: #ffffff; font-weight: 600; border-bottom: 1px solid rgba(0, 0, 0, 0.05); background-color:" & Session("CMFPNTITFT") & "; border-radius: 5px 5px 0px 0px; margin-bottom:5px; }" 'agrego td.header para color dinamico de cards NOVEDMENU.asp 13-01-26 CG

     'Publicaciones y recomendaciones y Mis autorizaciones y observaciones NOVEDADES
         'TITULO DEL DOCUMENTO
         s = s & "td.istt{ height:20px; text-align:center; xbox-shadow: 4px 4px 12px rgba(0,0,0,0.4);font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:" & Session("CMISTTFTSZ") & "pt;background-color: " & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"

				'****** NUEVA Para consulta de Formularios *****'
		          s = s & "td.isttf{ border-color:#ffffff;border-width:1px;xborder-radius:5px;border-style:solid;height:35px; text-align:left; xbox-shadow: 6px 6px 6px #dddddd;font-family:" & Session("CMISTTFTFC") & ";color:black;font-size:" & Session("CMISTTFTSZ") & "pt;background-color:#eeeeee;" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"
				  'color:" & Session("CMISTTFT")'
				  'background-color: " & Session("CMISTT") & "'
				  s = s & "td.isttfEst{ border-radius:0px 0px 5px 5px;border-style:none;height:35px; text-align:left; box-shadow: 6px 6px 6px #dddddd;font-family:" & Session("CMISTTFTFC") & ";color:black;font-size:" & Session("CMISTTFTSZ") & "pt;background-color:#eeeeee;" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"
				  'color:" & Session("CMISTTFT")'
				  'background-color: " & Session("CMISTT") & "'
  				  s = s & "td.isttfNom{ border-radius:0px 0px 5px 5px;border-style:none;height:35px; text-align:left; box-shadow: 6px 0px 12px #dddddd;font-family:" & Session("CMISTTFTFC") & ";color:black;font-size:" & Session("CMISTTFTSZ") & "pt;background-color:#eeeeee;" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"
				  'color:" & Session("CMISTTFT")'
				  'background-color: " & Session("CMISTT") & "'

      s = s &".block1 {"
      xs = s & " width: 50%; "
      s = s & " height: 200px; "
      s = s & "overflow:auto;border-style:solid;border-color:#6699CC;border-width:2;"
      s = s & "}"
      s = s & ".block2 {height: 90%;width:100%;xborder-color:#6699CC;border-style:none "
      s = s & "}"
      s = s & ".block3 {height: 90%;xborder-color:#6699CC;border-style:none "

      s = s &	 "}"

 '& Session("CMISTT") -- background-color del td
         'TITULO CONSULTAS
         s = s & "td.istxz{font-family:" & Session("CMISTXFTFC") & ";color:" & Session("CMISTXFT") & ";font-size:" & Session("CMISTXFTSZ") & "pt;background-color:" & Session("CMISTX") & ";" & replace(Session("CMISTXBL"),"1","font-weight:bold;") & replace(Session("CMISTXIT"),"1","font-style:italic;") & replace(Session("CMISTXUL"),"1","text-decoration:underline;border-radius:5px;") & "}"

         s = s & "td.ibtt{font-family:" & Session("CMIBTTFTFC") & ";color:" & Session("CMIBTTFT") & ";font-size:" & Session("CMIBTTFTSZ") & "pt;background-color:" & Session("CMIBTT") & ";" & replace(Session("CMIBTTBL"),"1","font-weight:bold;") & replace(Session("CMIBTTIT"),"1","font-style:italic;") & replace(Session("CMIBTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictt{font-family:" & Session("CMICTTFTFC") & ";color:" & Session("CMICTTFT") & ";font-size:" & Session("CMICTTFTSZ") & "pt;background-color:" & Session("CMICTT") & ";" & replace(Session("CMICTTBL"),"1","font-weight:bold;") & replace(Session("CMICTTIT"),"1","font-style:italic;") & replace(Session("CMICTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictx{font-family:" & Session("CMICTXFTFC") & ";color:" & Session("CMICTXFT") & ";font-size:" & Session("CMICTXFTSZ") & "pt;background-color:" & Session("CMICTX") & ";" & replace(Session("CMICTXBL"),"1","font-weight:bold;") & replace(Session("CMICTXIT"),"1","font-style:italic;") & replace(Session("CMICTXUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictr{font-family:" & Session("CMICTRFTFC") & ";color:" & Session("CMICTRFT") & ";font-size:" & Session("CMICTRFTSZ") & "pt;background-color:" & Session("CMICTR") & ";" & replace(Session("CMICTRBL"),"1","font-weight:bold;") & replace(Session("CMICTRIT"),"1","font-style:italic;") & replace(Session("CMICTRUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.icxr{font-family:" & Session("CMICXRFTFC") & ";color:" & Session("CMICXRFT") & ";font-size:" & Session("CMICXRFTSZ") & "pt;background-color:" & Session("CMICXR") & ";" & replace(Session("CMICXRBL"),"1","font-weight:bold;") & replace(Session("CMICXRIT"),"1","font-style:italic;") & replace(Session("CMICXRUL"),"1","text-decoration:underline;") & "}"


     if session("WEBDESK") = "S" then

'''background-color:#f0eedd;

     ''    s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#888888;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;" & XspEx("#f0f0f0",Session("CMFMENU"),0) & ";border-style:solid;border-width:1;border-color:#888888;height:20px;cursor:pointer;}"
     ''    s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#bbbbbb;height:20px;cursor:pointer;}"

      ''   s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#888888;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;" & XspEx("#f0f0f0",Session("CMFMENU"),0) & ";border-style:none;border-left-style:solid;border-left-width:2;border-left-color:#990000;border-right-style:solid;border-right-width:2;border-right-color:#990000;height:20px;cursor:pointer;}"

      ''   s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#555555;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:none;border-left-style:solid;border-left-width:1;border-left-color:#ff0000;border-right-style:solid;border-right-width:1;border-right-color:#ff0000;height:20px;cursor:pointer;}"



	  '-------------------------------------------------------------------------------------------------------
		'TARJETA DE NOVEDADES
	  '-------------------------------------------------------------------------------------------------------
	s = s & ".pnoved {white-space:nowrap;overflow:hidden;text-overflow: ellipsis;font-family:" & Session("CMFPNBOTFTFC") & ";position: absolute;bottom: 8px;left: 20px;right: 20px;font-size: 0.9rem;font-weight: 400;color: #f2f2f2;line-height: 1.4;background: rgba(0, 0, 0, 0.35);padding: 8px 14px;border-radius: 0 0 8px 8px;text-shadow: 0 1px 3px rgba(0, 0, 0, 0.4);backdrop-filter: blur(3px);max-width: 95%;z-index: 2;}" ''''''carrousel/slider de novedades/YT
	s = s & "div.contenedor2 {font-family:" & Session("CMIETTFTFC") & ";width: 100%;height: 100%;background-size: cover;border-radius: 5px;box-shadow: 0 8px 20px rgba(0, 0, 0, 0.05);backdrop-filter: blur(12px);-webkit-backdrop-filter: blur(12px);display: flex;flex-direction: column;overflow: hidden;transition: all 0.3s ease;background-color: FFFFFF;margin-botton: 5px;}"
	  '----------TEXTO DEL SLIDER--------
	  's = s & "span.txtsld{background-color:red;}"
	' SLIDER--------
	s = s & ".bx-wrapper ul li {height:220px;}"
    '-------------------------------------------------------------
      ' botones sig,anterior , agregar , buscar ,consultar
      '-------------------------------------------------------------


''''" & XspEx("#dddddd","#f9f9f9",0) & "

'''        s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#666666;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;;border-style:solid;border-width:1;border-color:#dddddd;border-left-style:solid;border-left-width:1;border-left-color:#ff0000;border-right-style:solid;border-right-width:1;border-right-color:#ff0000;height:20px;cursor:pointer;}"
'''         s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#cccccc;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#f0f0f0;height:20px;cursor:pointer;}"
       ' s = s & " SPAN.searchword { z-index:999;background-color:gray;color:#ffffff;cursor:pointer; }"
        if session("APLMDCLI")<> "BGAL" then
	       s = s & " SPAN.searchword { z-index:999;background-color:gray;color:#ffffff;cursor:hand; }"
	    else
	       s = s & " SPAN.searchword {z-index:999; background-color:#191ef3;color:#ffffff;cursor:hand; }"
        end if


        s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#666666;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;border-radius:15px;font-weight:bold;;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;}"
         s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#cccccc;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#f0f0f0;height:20px;cursor:pointer;}"

         s = s & "td.ts{font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";font-size:" & Session("CMFMNFTSZ") & "pt;;font-weight:600;border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";" & XspAlpha(100,60,1) & ";}"
         s = s & "td.ttb{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-indent:4px;;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

s=s & "#dhtmltooltip{"
s=s & "position: absolute;"
s=s & "left: -300px;"
s=s & "width: 250px;"
s=s & "border: 1px solid black;"
s=s & "padding: 2px;"
s=s & "background-color: lightyellow;"
s=s & "font-family: arial;"
s=s & "font-size:9pt;"
s=s & "font-style:italic;"
s=s & "visibility: hidden;"
s=s & "z-index: 100;"

s=s & "filter: progid:DXImageTransform.Microsoft.Shadow(color=gray,direction=135);"
s=s & "}"

s=s & "#dhtmlpointer{"
s=s & "position:absolute;"
s=s & "left: -300px;"
s=s & "z-index: 101;"
s=s & "visibility: hidden;"
s=s & "}"

s=s & "mark{"
s=s &    "background: orange;"
s=s &   " color: black;"
s=s & "}"


     '-------------------------------------------------------------




     ''  s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE","#888888",0) & ";}"
      ''   s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
       ''  s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:#BCBCBC;}"

         'originals = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
         'originals = s &    "border-style:solid;border-width:1;border-color:#000000;background-color:"& Session("CMFPNTIT") &";}"
			s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius: 5px;font-weight:600;" & "line-height:12pt;" & "text-align: center;" & "padding:6px 15px;" & "border-right:1px solid #cccccc;" & "border-bottom:2px solid #bbbbbb;" & "background-color:" & Session("CMFPNTIT") & ";" & "box-shadow:0 2px 4px rgba(0,0,0,0.1);" & " text-transform: uppercase;" & "letter-spacing: 0.5px;" & "}"
'			s = s & "td.ttx:last-child{border-right:none;}"



     ''    s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
      ''   s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-left-style:solid;border-left-width:2;border-left-color:#aa0000;background-color:#BCBCBC;}"
     ''    s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;border-bottom-style:solid;border-bottom-width:1pt;border-bottom-color:#eeeeee;" & XspSShw("#dddddd","125") & ";}"

        '     session("CMFPNGNFT") = "#000000"


    '--------------------PAGINA DEL BUSCADOR----------------------------

	'---------------- Documentos -------------
	's = s & "div.infodoc{display: flex; flex-direction:column; gap 3px; padding 8px 10 px; margin-bottom: 6px; border-radius:5px; border-left:4px solid; color:" & Session("CMIETT") & "; text-decoration: none; font-weight:500; font-size: 0.85rem; transition: background-color 0.2s ease, border-left-color 0.2s ease, color 0.2s ease; cursor:pointer;}"
	s = s & "div.infodoc{font-family:" & Session("CMIETTFTFC") & ";display: flex;position:relative; flex-direction:column; gap: 3px; padding: 10px 15px; margin-bottom: 6px; background-color: #f9f9f9; border-radius:10px; font-weight:500; transition: all 0.2s ease-in-out; cursor:pointer;}"
	s = s & "div.infodoc:hover{transform:translateY(-2px);z-index:100;background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%);scroll-margin-top: 100px; }"
	s = s & "div.infodoc::before{background:" & Session("CMIETT") & ";content:'';border-top-left-radius:5px;border-bottom-left-radius:5px;position:absolute;left:0; top:0; bottom:0;width:6px;}"
	s = s & "breadcrumb { margin: 6px 0 8px;display: flex; gap: 6px; flex-wrap: wrap;color:" & Session("CMFPNTIT") & ";font-size: 12.5px;}"
	s = s & "div.breadcrumb.current {color: var(--azul); font-weight: 600;}"
	s = s & "div.text-box {background: var(--badge-bg);border: 1px solid var(--borde);border-radius: 6px;padding: 8px 10px;font-size: 12.5px;line-height: 1.45;}"
	s = s & "div.attachments {margin-top: 10px;display: flex;flex-wrap: wrap;gap: 10px;}"
	s = s & "div.content2 {grid-area: content; min-width: 0;}"
	s = s & "div.doc-info {color:" & Session("CMIETT") & ";font-size:15pt;position: relative;display: inline-flex;align-items: center;justify-content: center;}"
	s = s & "nav.noved-doc-path {color:gray;font-size:0.8rem;}"
	s = s & "div.info-tip {position: absolute;z-index:9999;top: 24px;right: 25px;background: var(--badge-bg);border: 1px solid var(--borde);border-radius: 10px;box-shadow: var(--shadow);padding: 10px 12px;color:" & Session("CMFPNGNFT") & ";font-size: 12.5px;line-height: 1.45;opacity: 0;transform: translateY(-4px);pointer-events: none;transition: all 0.2s ease-in-out, transform .15s ease;width: max-content;max-width: 240px;z-index: 3;}"
	s = s & "div.excerpt {display: inline;color:" & Session("CMFPNTIT") & ";}"
	s = s & "div.pill {font-size: 10px;font-weight: 700;color: #808080;background: #f2f2f2;border: 1px solid #808080;padding: 2px 6px;border-radius: 15px;text-align: center;cursor: default;}"
	s = s & "div.title-row {font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNGNFT") & ";display: flex;align-items: center;gap: 8px;min-width: 0;width: 100%;justify-content:space-between;}"
	s = s & "a.enlace-con-ellipsis {width: 90%; display: inline-block; overflow: hidden;white-space:nowrap; text-overflow: ellipsis;}"
	s = s & "div.title-search {margin: 0; display: flex;flex: 1;min-width: 0; gap: 8px; align-items: center;}"
	s = s & "#loader-wrapper { position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); padding: 30px; background-color: #ffffff; border-radius: 10px; box-shadow: 0 5px 15px rgba(0, 0, 0, 0.2); display: flex;  flex-direction: column; align-items: center;  justify-content: center;  z-index: 9999; display: none; }"
        s = s & ".loader { border: 10px solid #f3f3f3; border-top: 10px solid #3498db; border-radius: 50%; width: 60px; height: 60px; animation: spin 1s linear infinite;}"
	s = s & "#loader-wrapper p {font-family: sans-serif;margin-top: 20px;color: #333;}"

	'---------------- Visor de detalles--------
s = s & "div.viewdoc{"
s = s & "  background-color:#fff;"
's = s & "  padding:14px 16px;"
s = s & "  border-radius:6px;"
s = s & "  box-shadow:0 1px 4px rgba(0,0,0,0.1);"
s = s & "  color:#333;"
s = s & "  font-size:0.9rem;"
's = s & "  margin-bottom:16px;"
s = s & "  border-left:4px solid #1e3a8a;"
s = s & "  height:100%;"
s = s & "  width:40%;"
s = s & "  display:flex;"
s = s & "  flex-direction:column;"
s = s & "  justify-content:space-between;"
s = s & "  position:absolute;"
s = s & "  top:65px;"
s = s & "}"
s = s & "div.viewdoc .info-line{margin-bottom:6px;}"
s = s & "div.viewdoc .info-line b{color:#111; font-weight:600; margin-right:4px;}"
s = s & "div.viewdoc .acciones{margin-top:12px; display:flex; gap:8px; flex-wrap:wrap;}"
s = s & "div.viewdoc .btn {"
s = s & "  width: 50%;"
's = s & "  background-color: #1e3a8a;"
s = s & "  color: #fff;"
s = s & "  border: none;"
s = s & "  padding: 10px 16px;"
s = s & "  border-radius: 6px;"
s = s & "  font-weight: 600;"
s = s & "  font-size: 0.9rem;"
s = s & "  box-shadow: 0 1px 2px rgba(0,0,0,0.1);"
s = s & "  cursor: pointer;"
s = s & "  transition: background-color 0.2s ease, transform 0.1s ease;"
s = s & "}"
s = s & "div.viewdoc .btn:hover {"
s = s & "  background-color: #152c66;"
s = s & "  transform: translateY(-1px);"
s = s & "}"
s = s & "div.viewdoc .btn:active {"
s = s & "  transform: scale(0.98);"
s = s & "}"

's = s & "td.tdview{border: 1px solid red;height:100%;}"

'TITULO DE BUSQUEDA DE DOCUMENTO

s = s & "div.titbsq {width:95%;overflow:hidden; white-space:nowrap;text-overflow:ellipsis;font-size: 1.1rem; color: #1E3A8A; font-weight: 700; cursor: pointer; text-decoration: none; margin-bottom: 4px; display: inline-block; transition: color 0.2s ease; }"
s = s & "div.titbsq:hover { color: #6A0DAD; text-decoration: underline; }"

	'--------------- Paginas ------------------
	's = s & "li.lipag{background-color:yellow important!;}"

	'-------------------
	'----------- SOLAPA GENERAL-----------
	'----------- tarjeta--------

s = s & "td.tt2{"
's = s & "font-family:" & Session("CMFPNTITFTFC") & ";"
s = s & "padding: 5px 10px;"
s = s & "margin-bottom:12px;"
s = s & "border-radius:6px;"
's = s & "background-color:" & Session("CMFPNTIT") & ";"
s = s & "background-color: #f9fafb;"
's = s & "box-shadow: 0 2px 6px rgba(0,0,0,0.1);"
s = s & "border: 1px solid #e5e7eb;"
s = s & "margin-right: 16px;"
s = s & "max-width: 25rem;"
s = s & "}"

'-------------- titulo
s = s & ".gentit{"
s = s & "font-family:" & Session("CMFPNTITFTFC") & ";"
's = s & "color:" & Session("CMFPNTITFT") & ";"
s = s & "color: #6b7280;"
s = s & "font-size:" & Session("CMFPNTITFTSZ") & "pt;"
s = s & "font-weight:600;"
s = s & "line-height:1.5;"
s = s & "text-indent:0;"
's = s & "text-align:center;"
s = s & "text-transform: uppercase;"
s = s & "margin-bottom: 4px;}"

'-------------- contenido
s = s & ".gencontent{"
s = s & "font-family:" & Session("CMFPNTITFTFC") & ";"
s = s & "color:" & Session("CMFPNIPFT") & ";"
s = s & "font-size:" & Session("CMFPNTITFTSZ") & "pt;"
s = s & "text-transform: capitalize !important;"
s = s & "font-weight:500;}"

'Separador hr de consultas
s = s & "hr{margin:2rem 0;border-color:" & Session("CMFMNFT") & "}"


      'tit paneles CMFPNTITFTFC
      '------------------------------------------------
         s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:5px;font-weight:600;line-height:10pt;padding:8px;border-style:solid;border-width:1;border-color:#ffffff;background-color:"& session("CMFPNTIT")  &";}"
'CMFPNTITFT
         s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:8px;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;padding-top: 10px;padding-bottom: 10px;"
         s = s &    "border-style:solid;border-width:1;border-color:#ffffff;background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "td.tc{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;width:14rem;border-radius:8px;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;padding-top: 10px;padding-bottom: 10px;"
         s = s &    "border-style:solid;border-width:1;border-color:#ffffff;background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:5px;text-align:center;font-weight:600;border-style:solid;border-width:1pt;border-color:white;line-height:20pt;" 'text-indent:5px
         s = s & "background-color:" & Session("CMFPNTIT") & ";}"   '#e0e0e0

         s = s & "td.mbt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & session("CMFMNBOT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;" & XspDShw("1","1","#f9f9f9") & ";padding-right:5px;}"
         s = s & "td.tu{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") + 2 & "px;font-weight:bold;}"
         s = s & "td.tx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNTITFT") & ";}"

         '---------------------
		 ' tit paneles
		 '--------------------

         s = s & "td.za{font-family:" & Session("CMFPNTITFTFC") & ";color:#ff0000;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;;}"
         s = s & "td.zb{font-family:" & Session("CMFPNTITFTFC") & ";color:#0000ff;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;;}"
         s = s & "td.zc{font-family:" & Session("CMFPNTITFTFC") & ";color:#000000;font-size:" & Session("CMFPNTITFTSZ") & "pt;;}"
         s = s & "td.zd{font-family:" & Session("CMFPNTITFTFC") & ";color:#000099;font-size:" & Session("CMFPNTITFTSZ") & "pt;;}"
         s = s & "td.ze{font-family:" & Session("CMFPNTITFTFC") & ";color:#000033;font-size:" & Session("CMFPNTITFTSZ") & "pt;;}"
         s = s & "td.tbh1{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;}"

       '----------------------------------------------------------------

         s = s & "td.ft{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;text-align: center;" & XspDShw("2","3","#cccccc") & ";}"
         s = s & "td.fti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";padding-left:20px;}"
         s = s & "td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;padding:0px 0px;font-weight:500;}"

		 s = s & "body.stbody{overflow: auto; Height: 100%;}"
		 s = s & "body.stbody{scrollbar-width: thin; scrollbar-color: rgba(0, 0, 0, 0.2) transparent;}"
		 s = s & "body.stbody{background:" & Session("CMFMENU")& ";}"
         '--------------------------------------
         '/* Ocultamos el body inicialmente */
         s = s & "body {"
		 s = s & "    opacity: 0;"
		 s = s & "    transition: opacity 0.2s ease-out" 'opacity 0.2s ease-out  0.3s ease-in;
		 s = s & "}"

		 '/* Clase que aplicaremos con JS */
			s = s & "body.content-ready {"
		 	s = s & "    opacity: 1;"
	 		s = s & "}"

		's = s & "body {"
		's = s & "    opacity: 0;"
		's = s & "    transform: translateY(10px);" ' /* Empieza 10px abajo */
		's = s & "    transition: "
		's = s & "        opacity 0.2s ease-out,"
		's = s & "        transform 0.25s ease-out;"
		's = s & "    will-change: opacity, transform;"
		's = s & "}"

		s = s & "body.content-ready {"
		s = s & "    opacity: 1;"
		s = s & "    transform: translateY(0);" '/* Sube a su posición original */
		s = s & "}"

         '---------------------------------------
		 s = s & "table.tnav{overflow: auto;}"
		 s = s & "table.tnav{list-style: none; margin: 0; background: #22388C; overflow: scroll; border-bottom-right-radius: 8px;}"
		 s = s & "table.tnav{font-size:" & Session("CMFMNFTSZ") & "pt;}"
		 s = s & "table.tnav{font-family:" & Session("CMFMNFTFC") & ";}"
		 's = s & "table.tnav{display: block;"
		 's = s & "table.tnav{margin: 4px 0;"
		 's = s & "table.tnav{border-radius: 6px;"
		 's = s & "table.tnav{text-decoration: none;"
		 's = s & "table.tnav{cursor: pointer;"
		 's = s & "table.tnav{transition: background-color 0.2s ease, color 0.2s ease;"
		 's = s & "}"



		 s = s & "td.utbrr{padding: 10px 14px;}"
		 's = s & "td.utbrr{display: flex; cursor: pointer;transition: background-color 0.2s ease;}"
		 's = s & "td.utbrr{border-left: 4px solid" & Session("APLITBGD") & ";}"


		 s = s & "elmclass{background-color:red;}"

		 s   = s & "li.utx{background-color:" &  Session("CMFMENU") & ";font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
        ' s = s & "li.utx{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"

		 '***** NUEVA para consulta de Formularios ******** '
		 s = s & "div.left-panel {flex: 2;font-family:" & Session("CMFPNGNFTFC") & "}"
		 s = s & "div.form-title {color:" & Session("CMFPNTIT") & ";font-size: 1.3rem;font-weight: 700;margin-bottom: 8px;}"
		 s = s & ".breadcrumb strong {color: #6b7280;}" 'ver de si se deja color: " & Session("CMFPNTIT") & "
		 s = s & "div.info-value {font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight: 500;color: " & Session("CMFPNTIT") & ";}"
		 s = s & "div.estado-label {display: inline-block; background-color:color-mix(in srgb, " & Session("CMIETT") & ", transparent 90%); color: " & Session("CMFPNTIT") & "; font-weight: 600; font-size: 0.75rem;padding: 4px 12px; border-radius: 20px;text-transform: uppercase; margin-bottom: 20px;}"
		 s = s & "td.utf{border-style:none; font-family:" & Session("CMFPNGNFTFC") & ";color:#ff7b3e; font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
		 'color:" & Session("CMFPNGNFT") & "'

         s = s & "td.uts{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;border-radius:5px;text-indent:5px;}"
         s = s & "td.utis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
		 s = s & "td.utiaa{border-radius:5px;}"
		 s = s & "td.utizz{border-radius:5px;}"
         s = s & "td.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;}"
         s = s & "td.utgi{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utgs{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utgis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.ut2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;font-weight:none;}"
         s = s & "td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;font-weight:600;}"
         s = s & "td.vtg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;border-radius:5px;font-weight:600;}"
         s = s & "td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;font-weight:600;}"
         s = s & "td.vtig{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.wt{font-family:" & Session("CMFPNGNFTFC") & ";color:#550000;font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.xt{font-family:" & Session("CMFPNGNFTFC") & ";color:#dddddd;font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;;font-weight:600;background-color:#336699;}"


         s = s & "table.tbb{xborder-style:none;xvisibility:hidden;}"

        s = s & "table.tbh{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";background-color:" & Session("CMFMNBOT") & ";box-shadow:0 0.125rem 0.25rem rgba(0,0,0,0.075);" & XspAlpha(0,100,1) & ";}"
		s = s & "table.tbh2{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";}"
     '  s = s &    "left:0px;"
      ' s = s &    "position:fixed;}"

         s = s & "table.tbp{border-radius:5px;" & XspEx(Session("APLBGD"),"",0) & ";}"
         s = s & "table.TDCM{border-style:outset;border-width:1;border-color:#ffffff;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"
         s = s & "table.TWIND{border-radius:5px;" & XspEx(Session("APLBGD"),"transparent",0) & ";}"
         s = s & "table.TWINE{border-radius:5px;border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"
         s = s & "table.TWINE2{background-color:#ffffff;}"
		 's = s & "table.aaa{margin-top: 15px;border-radius:5px;}"
s = s & _
"table.aaa{" & _
"max-width:90%;" & _
"margin:0 auto;" & _
"border:1px solid #dddddd;" & _
"border-radius:10px;" & _
"background:#ffffff;" & _
"box-shadow:0 4px 4px rgba(2,6,23,.08);" & _
"padding: 10px;" & _
"}"

s = s & _
"div.norma-card{" & _
"max-width:90%;" & _
"margin:auto;" & _
"border:1px solid #dddddd;" & _
"border-radius:10px;" & _
"background:#ffffff;" & _
"box-shadow:0 4px 4px rgba(2,6,23,.08);" & _
"padding: 10px;" & _
"}"

'"width:860px;"& _



     else


         s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;border-radius:15px;font-weight:bold;background-color:#f0eedd;border-style:solid;border-width:1;border-color:#888888;height:20px;cursor:pointer;}"
         s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#eeeeee;height:20px;cursor:pointer;}"

         s = s & "td.ts{font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";font-size:" & Session("CMFMNFTSZ") & "pt;;font-weight:600;border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";" & XspAlpha(100,60,1) & ";}"
        '---------------------
        ' tit paneles
        '--------------------

         s = s & "td.ttb{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-indent:4px;;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:5px;font-weight:600;border-style:solid;border-width:1pt;border-color:white;background-color:#e9e9e9;}"
         s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:5px;font-weight:600;line-height:10pt;padding:8px;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE","#888888",0) & ";}"
         s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;line-height:12pt;padding-right:4pt;margin: 5 15px;"
         s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:"& Session("CMFPNTIT") &";}"
         s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-radius:15px;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
         s = s &    "border-style:solid;border-width:1;border-color:#ffffff;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:#e9e9e9;}"
         s = s & "td.mbt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;" & XspDShw("1","1","#f9f9f9") & ";padding-right:5px;}"
         s = s & "td.tu{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") + 2 & "px;font-weight:bold;}"
         s = s & "td.tx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNTITFT") & ";}"
         's = s & "td.DialogTitlebar{ font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-style:outset; ;font-weight: 600;border-bottom: 1px solid rgba(0, 0, 0, 0.05);background-color:" & Session("CMFPNTITFT") & ;border-radius: 5px;'377396 margin-bottom:5px;";}"
         s = s & "td.DialogTitlebar{ font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;border-style:outset;font-weight: 600;border-bottom: 1px solid rgba(0, 0, 0, 0.05);background-color:" & Session("CMFPNTITFT") & ";border-radius: 5px;margin-bottom:5px;}"


}}
}
         '---------------------
		  ' tit paneles
		   '--------------------
		          s = s & "td.za{font-family:" & Session("CMFPNTITFTFC") & ";color:#ff0000;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;;}"
		          s = s & "td.zb{font-family:" & Session("CMFPNTITFTFC") & ";color:#0000ff;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;;}"
		          s = s & "td.zc{font-family:" & Session("CMFPNTITFTFC") & ";color:#000000;font-size:" & Session("CMFPNTITFTSZ") & "pt;;}"
		          s = s & "td.zd{font-family:" & Session("CMFPNTITFTFC") & ";color:#000099;font-size:" & Session("CMFPNTITFTSZ") & "pt;;}"
		          s = s & "td.ze{font-family:" & Session("CMFPNTITFTFC") & ";color:#000033;font-size:" & Session("CMFPNTITFTSZ") & "pt;;}"


         '------------------------------------------


         s = s & "td.ft{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";}"
         s = s & "td.fti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";padding-left:20px;}"
         s = s & "td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;padding:0px 0px;font-weight:500;}"
         s = s & "td.uts{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;}"
         s = s & "td.utgi{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utgs{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utgis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.ut2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;font-weight:none;}"
         s = s & "td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.vtg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;border-radius:5px;font-weight:600;}"
         s = s & "td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;border-radius:5px;font-weight:600;}"
         s = s & "td.vtig{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.wt{font-family:" & Session("CMFPNGNFTFC") & ";color:#550000;font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.xt{font-family:" & Session("CMFPNGNFTFC") & ";color:#dddddd;font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;;font-weight:600;background-color:#336699;}"

         s = s & "td.tbh1{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;;font-weight:600;}"


         s = s & "table.tbb{xborder-style:none;xvisibility:hidden;}"

         s = s & "table.tbh2{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";}"
         s = s & "table.tbh{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";background-color:" & Session("CMFMNBOT") & ";" & XspAlpha(0,100,1) & ";}"
         s = s & "table.tbp{border-radius:5px;" & XspEx(Session("APLBGD"),"",0) & ";}"
         s = s & "table.TDCM{border-style:outset;border-width:1;border-color:#ffffff;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"
         s = s & "table.TWIND{border-radius:5px;" & XspEx(Session("APLBGD"),"transparent",0) & ";}"
         s = s & "table.TWINE{border-radius:5px;border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"
         s = s & "table.TWINE2{background-color:#ffffff;}"


     end if

	 s = s & "input.btind{box-shadow: 5px 5px #cccccc;width:340px;color:black;background-color:#dddddd;font-weight:bold;height:40px;text-align:center;border-radius:20px;border-style:none;font-family:" & Session("CMIETXFTFC") & ";font-size:" & Session("CMIETXFTSZ") & "pt;}"


''' Fin Estilos ................

         s = s & "</style>"

      CorStyler2 = s


End Function
%>