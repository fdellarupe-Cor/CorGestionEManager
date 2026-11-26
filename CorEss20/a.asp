<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.write CorStyler3()

Function CorStyler3()
dim s
         if Session("CMFMEFTFC") = "" then
            Session("CMFMEFTFC")      = Session("CMFMNFTFC")
            Session("CMFMEFT")        = Session("CMFMNFT")
            Session("CMFMEFTSZ")      = Session("CMFMNFTSZ")
         end if

'''filter: gray;
'''filter: xray;

         s = "<META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " CONTENT=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & ">"
         s = "<style id='estil' type='text/css' >"

         s = s & "BODY {SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;}"


         s = s & "p{margin-top :0px;margin-bottom:0px;font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
         s = s & "a{text-decoration:none;}"
         s = s & "a.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;}"
         s = s & "a.link{color:#ffffff;text-decoration:none;}"
         s = s & "a.istt{font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:" & Session("CMISTTFTSZ") & "pt;background-color:" & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"
        ' s = s & "a.istt{color:#000000;text-decoration:none;}"
         s = s & "a.nn{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:#ffffff;text-transform:capitalize;font-weight:500;text-decoration:none;}"
         s = s & "a.nt{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-weight:500;text-decoration:none;}"

         s = s & "a.ntslp{font-size:" & Session("CMFMEFTSZ") & "pt;font-family:" & Session("CMFMEFTFC") & ";color:" &  Session("CMFMEFT")& ";font-weight:500;xtext-decoration:none;}"
         s = s & "a.ntmnn2{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & session("CMIBTTFT") & ";font-weight:500;text-decoration:none;}"

         s = s & "a.ntmnn{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & session("CMFMNFT") & ";font-weight:500;text-decoration:none;}"

     	  s = s & "a.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;text-decoration:none;}"
         s = s & "a.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;text-decoration:none;}"
         s = s & "a.cc{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";text-transform:uppercase;font-weight:800;text-decoration:none;}"
         s = s & "a.tt{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";text-transform:uppercase;font-weight:800;text-decoration:none;}"
		 s = s & "a.ts{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";text-transform:uppercase;font-weight:800;text-decoration:none;}"
         s = s & "a.cmb{font-size:" & Session("CMFPNIPFTSZ") & "pt;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-weight:normal;text-decoration:none;}"

         s = s & "a.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;text-decoration:none;}"

         s = s & "font.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1;border-color:#f9f9f9;}"

         s = s & "font.bt2{font-family:wingdings;color:#ffffff;font-size:11pt;font-weight:bold;background-color:#dd0000;border-style:inset;border-width:1;border-color:#ffffff;}"
         s = s & "font.bt1{font-family:wingdings;color:#ffffff;font-size:9pt;font-weight:bold;background-color:#00bb00;border-style:outset;border-width:1;border-color:#ffffff;}"
         s = s & "font.btc1{font-family:webdings;color:#ffffff;font-size:7pt;font-weight:bold;}"

    ' botones elim
         s = s & "input{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("aplbgd") & ";}"
         s = s & "input.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:#ffffff;}"
         s = s & "input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;}"
         s = s & "input.bte{font-family:'webdings';color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#990000","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#990000;height:20px;cursor:hand;}"
         s = s & "input.btG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#eeeeee;height:20px;cursor:hand;}"
         s = s & "input.btR{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#550000",0) & ";border-style:solid;border-width:1;border-color:#550000;height:20px;cursor:hand;}"

         s = s & "input.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dtb{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dtn{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#f0f0f0;}"
         s = s & "input.dts{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
         s = s & "input.dtbs{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;font-weight:bold;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"
         s = s & "input.dtns{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-align=right;border-top-style:none;border-left-style:none;border-right-style:none;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("CMFPNTIT") & ";background-color:#f0f0f0;}"


         s = s & "input.it{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"
         s = s & "input.rd{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"
         s = s & "input.tt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:600;background-color:" & session("CMFPNBOT") & ";}"
         s = s & "input.tt3{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:600;background-color:#993333;}"

         s = s & "option{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"
         s = s & "option.OCM{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;}"

         s = s & "select{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("APLBGD") & ";border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";}"
         s = s & "select.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1;border-color:" & Session("CMFPNTIT") & ";}"

         s = s & "span.at{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;background-color:" & Session("CMFPNIP") & ";}"

   '------------------
         s = s & "span.XMN{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";background-color:"&SESSION("CMICTT")&";border-style:solid;border-width:1pt;border-color:#000000;color:"&session("CMICTTFT")&";}"

         s = s & "span.XSL{font-size:" & Session("CMFMEFTSZ") & "pt;font-family:" & Session("CMFMEFTFC") & ";background-color:"&SESSION("CMIBTT")&";border-style:solid;border-width:1pt;border-color:#FFFFFF;color:"&session("CMIBTTFT")&";}"

         s = s & "span.ntslp{font-size:" & Session("CMFMEFTSZ") & "pt;font-family:" & Session("CMFMEFTFC") & ";color:" &  Session("CMFMEFT")& ";font-weight:500;}"

         s = s & "span.ntmnn{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & session("CMFMNFT") & ";font-weight:500;}"
         's = s & "a.ntmnn{font-size:" & Session("CMFMNFTSZ") & "pt;font-family:" & Session("CMFMNFTFC") & ";color:" & session("CMFMNFT") & ";font-weight:500;text-decoration:none;}"

   '-----------------------------
         s = s & "span.mt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNIP") & ";width:80%;}"

         s = s & "th{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

         s = s & "tr.st{background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "tr.cc{background-color:" & Session("CMFPNIT") & ";border-style:ridge;border-width:1pt;}"
         s = s & "tr.XFK{height:20px;}"

         s = s & "td.aa{text-transform: capitalize;}"
         s = s & "td{font-family:" & Session("CMFPNITFTFC") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;}"
         s = s & "td.at{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:500;}"
         s = s & "td.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:" & session("CMFPNBOT") & ";text-decoration:underline;border-style:outset;border-width:1pt;}"
         s = s & "td.btG{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#dddddd;border-style:inset;border-width:1pt;}"
         s = s & "td.cc{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
         s = s & "td.ee{font-size:" & Session("CMFPNITFTSZ") & "pt;font-family:" & Session("CMFPNITFTFC") & ";border-width:1pt;border-color:" & Session("CMFPNTIT") & ";color:" & Session("CMFPNITFT") & ";font-weight:500;}"
         s = s & "td.lt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;text-indent:10;}"
         s = s & "td.dt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.dt2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.et{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;}"
         s = s & "td.mt{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";font-size:" & Session("CMFPNITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNIT") & ";}"
         s = s & "td.qt{font-family:" & Session("CMFPNGRFTFC") & ";color:" & Session("CMFPNGRFT") & ";font-size:" & Session("CMFPNGRFTSZ") & "pt;font-weight:600;}"
         s = s & "td.rr{font-size:" & Session("CMFPNTITFTSZ") & "pt;font-family:" & Session("CMFPNTITFTFC") & ";border-width:1pt;border-color:" & Session("CMFMNBOT") & ";color:" & Session("CMFPNTITFT") & ";font-weight:500;}"
         s = s & "td.st{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"

         s = s & "td.tdcm1{cursor:hand;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;text-indent:5;background-color:#f0f0f0;}"
         s = s & "td.tdcm2{cursor:hand;font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNIPFTSZ") & "pt;border-style:solid;border-width:1pt;border-color:#0000dd;background-color:#0000dd;}"


         s = s & "td.btbs{color:#6699cc;cursor:hand;}"
         s = s & "td.btbn{color:#f7aa22;cursor:hand;}"
         s = s & "td.btba{color:#005500;cursor:hand;}"

         s = s & "td.btbsr{background-color:#6699cc;color:#ffffff;border-style:solid;border-width:1;}"
         s = s & "td.btbnr{background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;}"
         s = s & "td.btbar{background-color:#005500;color:#ffffff;border-style:solid;border-width:1;}"

         s = s & "div.ut  {border-style:none;border-width:1;border-color:#ffffff;}"
         s = s & "div.btbs{border-style:solid;border-width:1;border-color:#6699cc;}"
         s = s & "div.btbn{border-style:solid;border-width:1;border-color:#f7aa22;}"
         s = s & "div.btba{border-style:solid;border-width:1;border-color:#005500;}"

         s = s & "a.btbs{color:#6699cc;text-decoration:none;}"
         s = s & "a.btbn{color:#f7aa22;text-decoration:none;}"
         s = s & "a.btba{color:#005500;text-decoration:none;}"

         s = s & "font.btic{font-family:webdings;font-size:10pt;line-height:12pt;font-weight:normal;}"
         s = s & "font.btiw{font-family:wingdings;font-size:10pt;line-height:12pt;font-weight:normal;}"
         s = s & "font.btic2{font-family:" & Session("CMFPNIPFTFC") & ";font-size:9pt;line-height:12pt;font-weight:normal;}"


         s = s & "a.ft1n{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;}"
         s = s & "a.ft1o{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;text-indent:7pt;letter-spacing:3pt;color:#f7aa22;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#f7aa22;background-color:#ffffff;}"

         s = s & "a.ft2n{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#aa0000;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#aa0000;}"
         s = s & "a.ft2o{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;text-indent:7pt;letter-spacing:3pt;color:#aa0000;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#aa0000;background-color:#ffffff;}"

         s = s & "a.ft3n{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;letter-spacing:3pt;color:#00aa00;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#00aa00;}"
         s = s & "a.ft3o{font-family:wingdings;font-size:10pt;line-height:12pt;padding-left:3pt;text-indent:5pt;letter-spacing:3pt;color:#00aa00;font-weight:500;text-decoration:none;border-style:solid;border-width:1;border-color:#00aa00;background-color:#ffffff;}"

         s = s & "textarea{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#F9F9F9;}"
         s = s & "textarea.dt{font-family:" & Session("CMFPNIPFTFC") & ";color:" & Session("CMFPNIPFT") & ";font-size:" & Session("CMFPNIPFTSZ") + 1 & "pt;border-style:solid;border-width:1pt;border-color:#dddddd;background-color:#F9F9F9;}"

''' Estilos de Impresion......


         s = s & "div{page-break-after:always;}"
         s = s & "div.DINDM{background-color:" & Session("APLBGD") & ";border-style:SOLID;border-width:2;border-color:" & session("CMFMNBOT") & ";}"
         s = s & "div.DINDC{border-style:solid;border-width:1;border-color:#AAAAAA;background-color:#f9f9f9;}"


		 s = s & "td.iettprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;text-decoration:none;}"




		  s = s & "td.ietxprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"

		  s = s & "td.isttprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"


		  s = s & "td.istxzprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"

		  s = s & "td.ibttprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.icttprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.ictxprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.ictrprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"
		  s = s & "td.icxrprn{font-family:VERDANA;color:#000000;font-size:8pt;background-color:#ffffff;font-weight:bold;font-style:italic;text-decoration:none;}"



''''Session("CMFPNBOTFTFC")

         'Botones de panel

         s = s & "input.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;background-color:" & Session("CMFPNBOT") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
         s = s & "a.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.bietx{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;background-color:" & Session("CMFPNBOT") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"
         s = s & "input.bietxG{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;background-color:#bbbbbb;border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMFPNBOTBL"),"1","font-weight:bold;") & replace(Session("CMFPNBOTIT"),"1","font-style:italic;") & replace(Session("CMFPNBOTUL"),"1","text-decoration:underline;") & "}"



'----------------------------------------------------------------------------------------------------------------------

         'seccion ,indice normativo --- novedades
         s = s & "td.iett{font-family:" & Session("CMIETTFTFC") & ";color:" & Session("CMIETTFT") & ";font-size:" & Session("CMIETTFTSZ") & "pt;background-color:" & Session("CMIETT") & ";" & replace(Session("CMIETTBL"),"1","font-weight:bold;") & replace(Session("CMIETTIT"),"1","font-style:italic;") & replace(Session("CMIETTUL"),"1","text-decoration:underline;") & "}"







         'Publicaciones y recomendaciones y Mis autorizaciones y observaciones NOVEDADES


         s = s & "input.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:" & Session("CMIETX") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"
         s = s & "a.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ietx{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:" & Session("CMIETX") & ";border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"
         s = s & "input.ietxG{font-family:" & Session("CMIETXFTFC") & ";color:" & Session("CMIETXFT") & ";font-size:" & Session("CMIETXFTSZ") & "pt;background-color:#bbbbbb;border-style:solid;border-width:1;border-color:#666666;" & replace(Session("CMIETXBL"),"1","font-weight:bold;") & replace(Session("CMIETXIT"),"1","font-style:italic;") & replace(Session("CMIETXUL"),"1","text-decoration:underline;") & "}"






     'Publicaciones y recomendaciones y Mis autorizaciones y observaciones NOVEDADES
         'TITULO DEL DOCUMENTO
         s = s & "td.istt{font-family:" & Session("CMISTTFTFC") & ";color:" & Session("CMISTTFT") & ";font-size:" & Session("CMISTTFTSZ") & "pt;background-color:" & Session("CMISTT") & ";" & replace(Session("CMISTTBL"),"1","font-weight:bold;") & replace(Session("CMISTTIT"),"1","font-style:italic;") & replace(Session("CMISTTUL"),"1","text-decoration:underline;") & "}"

         'TITULO CONSULTAS
         s = s & "td.istxz{font-family:" & Session("CMISTXFTFC") & ";color:" & Session("CMISTXFT") & ";font-size:" & Session("CMISTXFTSZ") & "pt;background-color:" & Session("CMISTX") & ";" & replace(Session("CMISTXBL"),"1","font-weight:bold;") & replace(Session("CMISTXIT"),"1","font-style:italic;") & replace(Session("CMISTXUL"),"1","text-decoration:underline;") & "}"

         s = s & "td.ibtt{font-family:" & Session("CMIBTTFTFC") & ";color:" & Session("CMIBTTFT") & ";font-size:" & Session("CMIBTTFTSZ") & "pt;background-color:" & Session("CMIBTT") & ";" & replace(Session("CMIBTTBL"),"1","font-weight:bold;") & replace(Session("CMIBTTIT"),"1","font-style:italic;") & replace(Session("CMIBTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictt{font-family:" & Session("CMICTTFTFC") & ";color:" & Session("CMICTTFT") & ";font-size:" & Session("CMICTTFTSZ") & "pt;background-color:" & Session("CMICTT") & ";" & replace(Session("CMICTTBL"),"1","font-weight:bold;") & replace(Session("CMICTTIT"),"1","font-style:italic;") & replace(Session("CMICTTUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictx{font-family:" & Session("CMICTXFTFC") & ";color:" & Session("CMICTXFT") & ";font-size:" & Session("CMICTXFTSZ") & "pt;background-color:" & Session("CMICTX") & ";" & replace(Session("CMICTXBL"),"1","font-weight:bold;") & replace(Session("CMICTXIT"),"1","font-style:italic;") & replace(Session("CMICTXUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.ictr{font-family:" & Session("CMICTRFTFC") & ";color:" & Session("CMICTRFT") & ";font-size:" & Session("CMICTRFTSZ") & "pt;background-color:" & Session("CMICTR") & ";" & replace(Session("CMICTRBL"),"1","font-weight:bold;") & replace(Session("CMICTRIT"),"1","font-style:italic;") & replace(Session("CMICTRUL"),"1","text-decoration:underline;") & "}"
         s = s & "td.icxr{font-family:" & Session("CMICXRFTFC") & ";color:" & Session("CMICXRFT") & ";font-size:" & Session("CMICXRFTSZ") & "pt;background-color:" & Session("CMICXR") & ";" & replace(Session("CMICXRBL"),"1","font-weight:bold;") & replace(Session("CMICXRIT"),"1","font-style:italic;") & replace(Session("CMICXRUL"),"1","text-decoration:underline;") & "}"


     if session("WEBDESK") = "S" then

'''background-color:#f0eedd;

     ''    s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#888888;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;" & XspEx("#f0f0f0",Session("CMFMENU"),0) & ";border-style:solid;border-width:1;border-color:#888888;height:20px;cursor:hand;}"
     ''    s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#bbbbbb;height:20px;cursor:hand;}"

      ''   s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#888888;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;" & XspEx("#f0f0f0",Session("CMFMENU"),0) & ";border-style:none;border-left-style:solid;border-left-width:2;border-left-color:#990000;border-right-style:solid;border-right-width:2;border-right-color:#990000;height:20px;cursor:hand;}"

      ''   s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#555555;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:none;border-left-style:solid;border-left-width:1;border-left-color:#ff0000;border-right-style:solid;border-right-width:1;border-right-color:#ff0000;height:20px;cursor:hand;}"






    '-------------------------------------------------------------
      ' botones sig,anterior , agregar , buscar ,consultar
      '-------------------------------------------------------------

''''" & XspEx("#dddddd","#f9f9f9",0) & "

'''        s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#666666;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;;border-style:solid;border-width:1;border-color:#dddddd;border-left-style:solid;border-left-width:1;border-left-color:#ff0000;border-right-style:solid;border-right-width:1;border-right-color:#ff0000;height:20px;cursor:hand;}"
'''         s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#cccccc;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#f0f0f0;height:20px;cursor:hand;}"
        s = s & " SPAN.searchword { background-color:gray;color:#ffffff;cursor:hand; }"
        s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:#666666;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;}"
         s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#cccccc;font-size:" & Session("CMFPNBOTFTSZ") + 1 & "pt;font-weight:bold;background-color:#f0f0f0;border-style:solid;border-width:1;border-color:#f0f0f0;height:20px;cursor:hand;}"

         s = s & "td.ts{font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";font-size:" & Session("CMFMNFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";" & XspAlpha(100,60,1) & ";}"
         s = s & "td.ttb{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-indent:4px;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"
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
     '-------------------------------------------------------------




     ''  s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE","#888888",0) & ";}"
      ''   s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
       ''  s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:#BCBCBC;}"

         s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
         s = s &    "border-style:solid;border-width:1;border-color:#000000;background-color:"& Session("CMFPNTIT") &";}"


     ''    s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
      ''   s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-left-style:solid;border-left-width:2;border-left-color:#aa0000;background-color:#BCBCBC;}"
     ''    s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-bottom-style:solid;border-bottom-width:1pt;border-bottom-color:#eeeeee;" & XspSShw("#dddddd","125") & ";}"

        '     session("CMFPNGNFT") = "#000000"


    '------------------------------------------------

      'tit paneles CMFPNTITFTFC
      '------------------------------------------------
         s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:solid;border-width:1;border-color:#888888;background-color:"& session("CMFPNTIT")  &";}"
'CMFPNTITFT
         s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
         s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-left-style:solid;border-left-width:2;border-left-color:#aa0000;background-color:#e0e0e0;}"
         s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-bottom-style:solid;border-bottom-width:1pt;border-bottom-color:#000000;"
         s = s & "border-left-style:solid;border-left-width:2;border-left-color:#000000;background-color:#e0e0e0;}"

         s = s & "td.mbt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & session("CMFMNBOT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;" & XspDShw("1","1","#f9f9f9") & ";border-style:solid;border-width:1;border-color:#f9f9f9;cursor:hand;}"
         s = s & "td.tu{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") + 2 & "px;font-weight:bold;}"
         s = s & "td.tx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNTITFT") & ";}"

         '---------------------
		 ' tit paneles
		 '--------------------

         s = s & "td.za{font-family:" & Session("CMFPNTITFTFC") & ";color:#ff0000;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;text-transform:uppercase;}"
         s = s & "td.zb{font-family:" & Session("CMFPNTITFTFC") & ";color:#0000ff;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;text-transform:uppercase;}"
         s = s & "td.zc{font-family:" & Session("CMFPNTITFTFC") & ";color:#000000;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
         s = s & "td.zd{font-family:" & Session("CMFPNTITFTFC") & ";color:#000099;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
         s = s & "td.ze{font-family:" & Session("CMFPNTITFTFC") & ";color:#000033;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
         s = s & "td.tbh1{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;}"

       '----------------------------------------------------------------

         s = s & "td.ft{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";}"
         s = s & "td.fti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";padding-left:20px;}"
         s = s & "td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
         s = s & "td.uts{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;}"
         s = s & "td.utgi{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utgs{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utgis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.ut2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;font-weight:none;}"
         s = s & "td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.vtg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;}"
         s = s & "td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.vtig{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.wt{font-family:" & Session("CMFPNGNFTFC") & ";color:#550000;font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.xt{font-family:" & Session("CMFPNGNFTFC") & ";color:#dddddd;font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"


         s = s & "table.tbb{border-style:none;visibility:hidden;}"

         s = s & "table.tbh2{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";}"
         s = s & "table.tbh{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";background-color:" & Session("CMFMNBOT") & ";" & XspAlpha(0,100,1) & ";}"
         s = s & "table.tbp{border-style:SOLID;border-width:1;border-color:#EEEEEE;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"
         s = s & "table.TDCM{border-style:outset;border-width:1;border-color:#ffffff;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"
         s = s & "table.TWIND{border-style:solid;border-width:1;border-color:#eeeeee;" & XspEx(Session("APLBGD"),"#AAAAAA",0) & ";}"
         s = s & "table.TWINE{border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"
         s = s & "table.TWINE2{border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"


     else


         s = s & "input.btsb{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;background-color:#f0eedd;border-style:solid;border-width:1;border-color:#888888;height:20px;cursor:hand;}"
         s = s & "input.btsbG{font-family:" & Session("CMFPNBOTFTFC") & ";color:#bbbbbb;font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#eeeeee","#dddddd",0) & ";border-style:solid;border-width:1;border-color:#eeeeee;height:20px;cursor:hand;}"

         s = s & "td.ts{font-family:" & Session("CMFMNFTFC") & ";color:" & Session("CMFMNFT") & ";font-size:" & Session("CMFMNFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";" & XspAlpha(100,60,1) & ";}"
        '---------------------
        ' tit paneles
        '--------------------

         s = s & "td.ttb{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-indent:4px;text-transform:uppercase;font-weight:600;background-color:" & Session("CMFPNTIT") & ";}"
         s = s & "td.ttt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;border-bottom-style:solid;border-bottom-width:1pt;border-bottom-color:#000000;background-color:#e9e9e9;}"
         s = s & "td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE","#888888",0) & ";}"
         s = s & "td.ttx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
         s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:"& Session("CMFPNTIT") &";}"
         s = s & "td.tt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:12pt;text-indent:4pt;padding-right:4pt;"
         s = s &    "border-style:solid;border-width:1;border-color:#BBBBBB;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#990000;border-right-style:solid;border-right-width:1;border-right-color:#aa0000;background-color:#e9e9e9;}"
         s = s & "td.mbt{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;font-weight:600;" & XspDShw("1","1","#f9f9f9") & ";border-style:solid;border-width:1;border-color:#f9f9f9;cursor:hand;}"
         s = s & "td.tu{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") + 2 & "px;font-weight:bold;}"
         s = s & "td.tx{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;border-style:outset;border-width:1pt;background-color:" & Session("CMFPNTITFT") & ";}"
          '---------------------
		  ' tit paneles
		   '--------------------
		          s = s & "td.za{font-family:" & Session("CMFPNTITFTFC") & ";color:#ff0000;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;text-transform:uppercase;}"
		          s = s & "td.zb{font-family:" & Session("CMFPNTITFTFC") & ";color:#0000ff;font-size:" & Session("CMFPNTITFTSZ") + 2 & "pt;text-transform:uppercase;}"
		          s = s & "td.zc{font-family:" & Session("CMFPNTITFTFC") & ";color:#000000;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
		          s = s & "td.zd{font-family:" & Session("CMFPNTITFTFC") & ";color:#000099;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"
		          s = s & "td.ze{font-family:" & Session("CMFPNTITFTFC") & ";color:#000033;font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;}"


         '------------------------------------------


         s = s & "td.ft{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";}"
         s = s & "td.fti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;" & XspDShw("2","3","#cccccc") & ";padding-left:20px;}"
         s = s & "td.ut{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;}"
         s = s & "td.uts{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;}"
         s = s & "td.utgi{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;}"
         s = s & "td.utgs{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.utgis{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:500;padding-left:20px;" & XspSShw("#dddddd","125") & ";}"
         s = s & "td.ut2{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;font-weight:none;}"
         s = s & "td.vt{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.vtg{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;}"
         s = s & "td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.vtig{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:600;padding-left:20px;}"
         s = s & "td.wt{font-family:" & Session("CMFPNGNFTFC") & ";color:#550000;font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;}"
         s = s & "td.xt{font-family:" & Session("CMFPNGNFTFC") & ";color:#dddddd;font-size:" & Session("CMFPNGNFTSZ") + 1 & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"

         s = s & "td.tbh1{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTIT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;}"


         s = s & "table.tbb{border-style:none;visibility:hidden;}"

         s = s & "table.tbh2{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";}"
         s = s & "table.tbh{BORDER-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMNBOT") & ";background-color:" & Session("CMFMNBOT") & ";" & XspAlpha(0,100,1) & ";}"
         s = s & "table.tbp{border-style:SOLID;border-width:1;border-color:#EEEEEE;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"
         s = s & "table.TDCM{border-style:outset;border-width:1;border-color:#ffffff;" & XspEx(Session("APLBGD"),"#CCCCCC",0) & ";}"
         s = s & "table.TWIND{border-style:solid;border-width:1;border-color:#eeeeee;" & XspEx(Session("APLBGD"),"#AAAAAA",0) & ";}"
         s = s & "table.TWINE{border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"
         s = s & "table.TWINE2{border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f9f9f9;}"


     end if

''' Fin Estilos ................

         s = s & "</style>"

      CorStyler3 = s & Session("CMFPNIPFTFC")


End Function
%>