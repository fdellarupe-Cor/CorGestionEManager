<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0
  bModificar       = False
  ofv.ObtenerAtributos Session("Form"),""
  bModificar       = ofv.AModificar


  sSClick          = request.querystring("SSClick")

  sHTML  = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & "  if (document.all.festad) { " & chr(13)
    ''' wxs = wxs & " alert('yaaaa!!'); " & Chr(13)
     wxs = wxs & " document.festad.submit(); " & Chr(13)
     wxs = wxs & " } else { " & chr(13)
     wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
     wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
     wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "<body onload='IniWin();' >"





''''''''''
  sImagen0   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#064;</FONT>"
  sImagen1   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0202;</FONT>"
  sImagen2   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0130;</FONT>"
  sImagen3   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0194;</FONT>"
  sImagen4   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >Ñ</FONT>"
  sImagen5   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >&#0204;</FONT>"
  sImagen6   = "<FONT FACE=WebDINGS COLOR=" & session("CMFMNBOT") & " SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" >q</FONT>"

'''''''


  StyL  = " STYLE='cursor:hand;filter:dropShadow(OFFY=1,OFFX=1,color=#f9f9f9);"
  StyL  = StyL & "border-style:solid;border-width:1;border-color:#dddddd;font-weight:normal;"
  StyL  = StyL & "' "

  StyL2  = " STYLE='"
  StyL2  = StyL2 & "border-style:outset;border-width:1;border-color:#ffffff;"
  StyL2  = StyL2 & "filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr="
  StyL2  = StyL2 & "#ffffff, endColorStr=#aaaaaa, gradientType=0);"
  StyL2  = StyL2 & "' "

  sHTML1 = ""
  sHTML1 = sHTML1 & ofv.GenTabla("panelM","aaa " & sTyl2 & " ","","1%","",0,0,-2,-1,"","","","no")
  sHTML1 = sHTML1 & ofv.GenRow("","","","","","","","","no")
'''''''''''''''''''''''''''
'' TODAVIA NO HAY NADA '''''
''''''''''''''''''''''''''''''

  sHTML1 = sHTML1 & ofv.GenFRow()
  sHTML1 = sHTML1 & ofv.GenFTabla()



  xSolasp = ""  
  xSolasp = xSolasp & "BarraREQSolNV.asp?tsk=" & tsk & "&SSCLICK=" & SSCLICK

    sHTML = sHTML & "<form name=festad action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK1' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
