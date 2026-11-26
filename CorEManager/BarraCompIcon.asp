<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=-10
  bModificar = False
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  sTipo    = request.querystring("Tipo")
  sCodigo  = request.querystring("Codigo")
  sFuncion = request.querystring("Funcion")
  sSClick  = request.querystring("SSClick")
  sMenu    = request.querystring("Menu")
  sSector  = request.querystring("Sector")
  sNIV1    = request.querystring("NIV1")
  sNIV2    = request.querystring("NIV2")
  sNIV3    = request.querystring("NIV3")
  sNIV4    = request.querystring("NIV4")
  sNIV5    = request.querystring("NIV5")

  sPerfil      = request.querystring("Perfil")
  sComplemento = request.querystring("Complemento")
  sComplemento = ofv.convertircar(sComplemento,"_"," ")

 '' response.write "recibe " & request.querystring & "<br>"

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
   '  wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
   '  wxs = wxs & "  if (document.all.festad) { " & chr(13)
    ''' wxs = wxs & " alert('yaaaa!!'); " & Chr(13)
     wxs = wxs & " document.getElementById('festad').submit(); " & Chr(13)
    ' wxs = wxs & " } else { " & chr(13)
    ' wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
    ' wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
    ' wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "<body bgcolor='red' >"


  sImagen1   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0202;</FONT>"






  sHTML1 = sHTML1 & ofv.GenTabla("panel2","aaa  STYLE='background-color:#dddddd;' ","","1%","",0,0,1,2,"","","","no")

  sHTML1 = sHTML1 & ofv.GenRow("","","","","20","","","","no")

 ' StyL  = " STYLE='cursor:pointer;filter:dropShadow(OFFY=1,OFFX=1,color=#dddddd);"
 ' StyL  = StyL & "border-right-style:solid;border-right-width:1;border-right-color:#dddddd;font-weight:normal;"
 ' StyL  = StyL & "' "

      StyL  = " STYLE='cursor:pointer;font-weight:normal;padding-left:5px;padding-right:5px;"
      StyL  = StyL & "border-style:solid;border-width:1;border-color:#000000;background-color:#aa99cc;"
      StyL  = StyL & "' "


sPrint = "S"

if  sAcceso = "N" then
    sPrint = "N"
elseif  session("USACC") = "R" and sAcceso = "L" then
        sPrint = "N"
elseif  session("USRES") = "B" and sResol = "A" then
        sPrint = "N"
end if

  if  Session("Segco") = "S" and sPrint = "S" then
      t2    = "<a href=generarComp.asp?id="& sCodigo &" target=_blank  title=" & chr(34)
      t2    = t2 & " Vista&nbsp;previa&nbsp;de&nbsp;impresion&nbsp;del&nbsp;componente" & chr(34) & " >" & sImagen1 & "</a>"
      sHTML1 = sHTML1 & ofv.GenCelda("","aT  " & StyL & " ","center","6","","","",t2,"si")
  end if


  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "DESA" then
      if  bModificar then
          t2    = "<a href='edobjsdk.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2    = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal><IMG SRC=Imagenes\ICorsecure.jpg title=" & chr(34)
          t2    = t2 & " Permisos&nbsp;del&nbsp;Esquema" & chr(34) & " border=0></a>"
       ''   sHTML = sHTML & ofv.GenCelda("","vt","","2","","","",t2,"si")
      end if
  end if



  if sFuncion = "E" and Session("Segce") = "S" then

  if session("ModMain") = "WK"  then
     t2 = "<a href='coWKFREDIR.asp?sopcion=VER&"
  elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "PARA" then
     t2 = "<a href='coWKFREDIR.asp?sopcion=APR&"
  elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) <> "PARA" AND NOT Session("SegcT") = "S" then
     t2 = "<a href='coWKFREDIR.asp?sopcion=VER2&"
  else
     t2 = "<a href='edobjwfl.asp?"
  end if

  t2 = t2 & "Tipobjeto=" & sTipo & "&Codobjeto="
  t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal><IMG SRC=Imagenes\ICorSInfoUser.jpg title=" & chr(34)
  t2 = t2 & " Circuito&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;del&nbsp;Esquema" & " " & session("ModMain") & " " & Session("Form2") & chr(34) & " border=0></a>"
'''  sHTML = sHTML & ofv.GenCelda("","vt","","2","","","",t2,"si")


  end if


  if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "DESA"  then
      if  bModificar then
          t2    = "<a href='edobjdist.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2    = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPDST" & "' target=Principal><IMG SRC=Imagenes\ICorLdist.jpg title=" & chr(34)
          t2    = t2 & " Lista&nbsp;de&nbsp;distribuci&oacute;n&nbsp;" & chr(34) & " border=0></a>"
         ''' sHTML = sHTML & ofv.GenCelda("","vt","","2","","","",t2,"si")
      end if
  end if





  sHTML1 = sHTML1 & ofv.GenFRow()
  sHTML1 = sHTML1 & ofv.GenFTabla()

   if session("WEBDESK") = "S" then

        sHTML = sHTML & ofv.GenTabla("panel","aaa  ","","100%","",0,0,"-2","-2","","","","no")
        sHTML = sHTML & ofv.GenRow("","","","","","","","","no")
        sHTML = sHTML & ofv.GenCelda("","ut","right","99%","","","",sHTML1,"si")
        sHTML = sHTML & ofv.GenCelda("","ut","right","1%","","","","&nbsp;","si")
        sHTML = sHTML & ofv.GenFRow()
        sHTML = sHTML & ofv.GenFTabla()

   else

   end if



  xSolasp = ""
  xSolasp = xSolasp & "BarraCompSol.asp?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
  xSolasp = xSolasp & "&Funcion=" & sFuncion & "&SSCLICK=" & sSClick & "&Menu=" & sMenu & "&Sector=" & sSector
  xSolasp = xSolasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

    sHTML = sHTML & "<form name=festad id=festad action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK1' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " IniWin(); "
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "aaaaa</body></html>"


end if

Response.Write sHTML

%>
