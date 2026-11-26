<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  bModificar = False
  bModificaresp = False
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  bModificaresp = ofv.AModificar
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
  sMAINAPL = request.querystring("MAINAPL")

  session("modalidad") = "ESP"

 '' response.write "recibe " & request.querystring & "<br>"

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = "<html>"

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


  siconmas = "ICorAcciones.jpg"





''''''''''
  sImagen0   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#064;</FONT>"
  sImagen1   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0202;</FONT>"
  sImagen2   = "<FONT FACE=corflaticon COLOR='#0000ff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000ff';"" >&#0130;</FONT>"
  sImagen2O  = "<FONT FACE=corflaticon COLOR='#ff0000' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ff0000';"" >&#094;</FONT>"
  sImagen2H  = "<FONT FACE=corflaticon COLOR='#FFFFFF' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#FFFFFF';"" >&#095;</FONT>"
  sImagen3   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0194;</FONT>"
  sImagen4   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >Ñ</FONT>"
  sImagen5   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0204;</FONT>"
  sImagen6   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >q</FONT>"
  sImagen7   = "<FONT FACE=corflaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#043;</FONT>"

'''''''


  StyL  = " STYLE='cursor:pointer;filter:dropShadow(OFFY=1,OFFX=1,color=#dddddd);"
  StyL  = StyL & "border-right-style:solid;border-right-width:1;border-right-color:#dddddd;font-weight:normal;"
  StyL  = StyL & "' "

  StyL2  = " STYLE='"
  StyL2  = StyL2 & "border-style:outset;border-width:1;border-color:#ffffff;"
  StyL2  = StyL2 & "filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr="
  StyL2  = StyL2 & "#ffffff, endColorStr=#aaaaaa, gradientType=0);"
  StyL2  = StyL2 & "' "

  sHTML1 = ""
  sHTML1 = sHTML1 & ofv.GenTabla("panelM","aaa ","","1%","",0,0,-2,-1,"","","","no")
  sHTML1 = sHTML1 & ofv.GenRow("","","","","","","","","no")

 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

    sHTML1 = sHTML1 & ofv.GenCelda("","UT  ","","","","","","&nbsp;","si")

' HISTORIAL DE OBSERVACIONES -----------------
  if  Session("Segce") = "S" then

      t2 = "<a href='COOBSHIST.asp?"
      t2 = t2 & "Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2H & "</a>"
      SBOTTIT = "HISTORIAL&nbsp;de&nbsp;Observaciones. "
      sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

  end if

 ELSE

  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      sEsquema = rs(1)
      siconmas = rs("archivo")
      sPerfil  = rs(3)
      sAcceso  = rs(16)
      sResol   = rs(17)
      sTipoCPT = rs(21)
  end if
  call ofv.cerrarconsulta(rs)

sPrint = "S"

if sAcceso = "N" then
   sPrint = "N"
elseif session("USACC") = "R" and sAcceso = "L" then
   sPrint = "N"
elseif session("USRES") = "B" and sResol = "A" then
   sPrint = "N"
end if

  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "DESA" then
      if  bModificar then
          t2 = "<a href='EDperfilesdocProp.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPPRF"  & "' target=Principal>"
          t2 = t2 & sImagen0 & "</a>"
          SBOTTIT = "Editar&nbsp;perfil&nbsp;del&nbsp;esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")
      end if
  end if


  if  Session("Segco") = "S" and sPrint = "S" then
      t2 = "<a href='Reportes.asp?ID=" & sCodigo & "&Titulo=Esquemas_Normativos&perfil="
      t2 = t2 & sPerfil & "&tipobj=" & xtipobj & "&GLBSOLN=" & "ESPREP"  & "' Target=_blank>"
      ' t2 = t2 & sPerfil & "&tipobj=" & sTipo & "&GLBSOLN=" & "ESPREP"  & "' Target=_blank>"

      t2 = t2 & sImagen1 & "</a>"
      SBOTTIT = "Vista&nbsp;previa&nbsp;de&nbsp;impresión&nbsp;del&nbsp;Esquema."
      sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","center","6","","","",t2,"si")

      if  sFuncion = "Exx" and Session("Segct") = "S" then
          if  bModificar then
              t2 = "<a href=edcor2word.asp?ID=" & sCodigo & "&Titulo=Esquemas_Normativos&perfil="
              t2 = t2 & sPerfil & "&tipobj=" & xtipobj & "&GLBSOLN=" & "ESPWRD"  & " Target=_blank>"
              t2 = t2 & sImagen6 & "</a>"
              SBOTTIT = "Convertir&nbsp;a&nbsp;Word."
              sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","center","6","","","",t2,"si")
          end if
      end if

  end if


  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND ( (sTipoCPT = "D" and LEFT(Session("Form2"),4) = "DESA") or (sTipoCPT = "C") ) then
      if  bModificar  or bModificaresp  then
          t2 = "<a href='edobjsdk.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal>"
          t2 = t2 & sImagen4 & "</a>"
          SBOTTIT = "Permisos&nbsp;del&nbsp;Esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")
      end if
  end if

  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND Session("CORUNQ") = "S" then
          t2 = "<a href='CoobjURL.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal>"
          t2 = t2 & sImagen7 & "</a>"
          SBOTTIT = "Direccion&nbsp;del&nbsp;Esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")
  end if

  if  sFuncion = "E" and Session("Segce") = "S" then

      if  session("ModMain") = "WK"  then
          t2 = "<a href='coWKFREDIR.asp?sopcion=VER&"
      elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "PARA" then
          t2 = "<a href='coWKFREDIR.asp?sopcion=APR&"
      elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) <> "PARA" AND NOT Session("SegcT") = "S" then
          t2 = "<a href='coWKFREDIR.asp?sopcion=VER2&"
      else
          t2 = "<a href='edobjwfl.asp?"
      end if

      t2 = t2 & "PASO=AUTO&Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2 & "</a>"
      SBOTTIT = "Circuito&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;del&nbsp;Esquema."
      sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

  end if


' AGREGADO NUEVO  PASO OBSERVACION-----------------
if  sFuncion = "E" and Session("Segce") = "S" then

      if  session("ModMain") = "WK"  then
          t2 = "<a href='coWKFREDIR.asp?sopcion=VER&"
      elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "PARA" then
          t2 = "<a href='coWKFREDIR.asp?sopcion=APR&"
      elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) <> "PARA" AND NOT Session("SegcT") = "S" then
          t2 = "<a href='coWKFREDIR.asp?sopcion=VER2&"
      else
          t2 = "<a href='edobjwfl.asp?"
      end if

      t2 = t2 & "PASO=OBS&Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2 & sImagen2O & "</a>"
      SBOTTIT = "Circuito&nbsp;de&nbsp;Observaciones. "
      sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

  end if

' HISTORIAL DE OBSERVACIONES -----------------
if  (sFuncion = "E" and Session("Segce") = "S" ) or session("ModMain")="WK" then

      t2 = "<a href='COOBSHIST.asp?"
      t2 = t2 & "Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2H & "</a>"
      SBOTTIT = "HISTORIAL&nbsp;de&nbsp;Observaciones. " ' & request.querystring
      sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

  end if

'------------------------------

  if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "DESA"  then
      if  bModificar or bModificaresp then

          t2 = "<a href='edobjdist.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPDST" & "' target=Principal>"
          t2 = t2 & sImagen3 & "</a>"
          SBOTTIT = "Lista&nbsp;de&nbsp;distribuci&oacute;n."
          sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

      end if
  end if

  if  sFuncion = "Exx" and Session("Segct") = "S" and session("ModMain") = "ED"  then
      if  bModificar then

          t2 = "<a href='edGenEstad.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPGST"  & "' target=Principal>"
          t2 = t2 & sImagen5 & "</a>"
          SBOTTIT = "Estadisticas&nbsp;del&nbsp;Esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

     end if
  end if

  if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "DESA"  then
      if  bModificar or bModificaresp then

          t2 = "<a href='edobjACD.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPDST" & "' target=Principal>"
          t2 = t2 & sImagen5 & "</a>"
          SBOTTIT = "Lista de accesos directos para el documento."
          sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

      end if
  end if

  '''''    sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",bModificaresp & "-" & Session("Segct") & "-" & session("ModMain") & "-" & LEFT(Session("Form2"),4),"si")


 END IF



  sHTML1 = sHTML1 & ofv.GenFRow()
  sHTML1 = sHTML1 & ofv.GenFTabla()

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

  sHTML = sHTML & ofv.GenTabla("panel","aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ","","100%","",0,0,"-2","-2","","","","no")
  sHTML = sHTML & ofv.GenRow("","","","","","","","","no")
  sHTML = sHTML & ofv.GenCelda("","ut","right","1%","","","","&nbsp;","si")
  sHTML = sHTML & ofv.GenCelda("","ut","","1%","","","",sHTML1,"si")
  sHTML = sHTML & ofv.GenCelda("","ut","","99%","","","","&nbsp;","si")
  sHTML = sHTML & ofv.GenFRow()
  sHTML = sHTML & ofv.GenFTabla()


  xSolasp = ""


  xSolasp = xSolasp & "BarraSol.asp"
  xSolasp = xSolasp & "?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
  xSolasp = xSolasp & "&Funcion=" & sFuncion & "&SSCLICK=" & sSClick
  xSolasp = xSolasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

    sHTML = sHTML & "<form name=festad action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK1' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
