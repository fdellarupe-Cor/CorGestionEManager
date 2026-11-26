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
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  bModificaresp = ofv.AModificar
  if session("APLMDCLI")="" then
    	    aplcheck()
  end if

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

 'response.write "recibe " & Session("Form2") &"          "  &session("APLMDCLI")


 'response.write sTipo & " sTipo " & "<br>"

  set cn = ofv.conectar(ofv.strconn2)


     sql="select * from objprty where codobjeto='" & sCodigo &  "'" & " and autor='" & session("usuario") & "'"
     set rs = ofv.crearconsultaEx(sql,cn,1,parametros)
     if not rs.eof then
       bModificaresp=true
     else
      bModificaresp=false or ofv.AModificar
    end if

   call ofv.cerrarconsulta(rs)
  sHTML = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & "  if (document.getElementById('festad')) { " & chr(13)
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


  sImagen0   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57422;</FONT>"
  sImagen0p   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57354;</FONT>"
  sImagen1   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57415;</FONT>"
  sImagen2   = "<FONT FACE=CorFlaticon COLOR='#0000ff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000ff';"" >&#57387;</FONT>"
  sImagen2O  = "<FONT FACE=CorFlaticon COLOR='#ff0000' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ff0000';"" >&#57362;</FONT>"
  sImagen2H  = "<FONT FACE=CorFlaticon COLOR='#FFFFFF' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#FFFFFF';"" >&#57436;</FONT>"
  sImagen3   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57349;</FONT>"
  sImagen4   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57393;</FONT>"
  sImagen5   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57383;</FONT>"
  sImagen6   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >q</FONT>"
  sImagen7   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#57417;</FONT>"


'''''''
''''''''''
  sImagen0   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#064;</FONT>"
  sImagen0p   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0163;</FONT>"
  sImagen1   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0202;</FONT>"
  sImagen2   = "<FONT FACE=CorFlaticon COLOR='#0000ff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000ff';"" >&#0130;</FONT>"
  sImagen2O  = "<FONT FACE=CorFlaticon COLOR='#ff0000' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ff0000';"" >&#094;</FONT>"
  sImagen2H  = "<FONT FACE=CorFlaticon COLOR='#FFFFFF' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#FFFFFF';"" >&#095;</FONT>"
  sImagen3   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0194;</FONT>"
  sImagen4   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >Ñ</FONT>"
  sImagen5   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0204;</FONT>"
  sImagen6   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >q</FONT>"
  sImagen7   = "<FONT FACE=CorFlaticon COLOR='#ffffff' SIZE=3 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#043;</FONT>"

'''''''



   if session("WEBDESK") = "S" then

      StyL  = " STYLE='cursor:pointer;font-weight:normal;padding-left:5px;padding-right:5px;"
      StyL  = StyL & "border-style:solid;border-width:1;border-color:#000000;background-color:#aa99cc;"
      StyL  = StyL & "' "

      sHTML1 = ""
      sHTML1 = sHTML1 & ofv.GenTabla("panelM","aaa   STYLE='background-color:#dddddd;' ","","1%","",0,0,1,2,"","","","no")
      sHTML1 = sHTML1 & ofv.GenRow("","","","","","","","","no")

    '  sHTML1 = sHTML1 & ofv.GenCelda("","aT " & StyL & " ","","1","","","","&nbsp;","si")



   else

      StyL  = " STYLE='cursor:pointer;filter:dropShadow(OFFY=1,OFFX=1,color=#dddddd);"
      StyL  = StyL & "border-right-style:solid;border-right-width:1;border-right-color:#dddddd;font-weight:normal;"
      StyL  = StyL & "' "

      sHTML1 = ""
      sHTML1 = sHTML1 & ofv.GenTabla("panelM","aaa ","","1%","",0,0,-2,-1,"","","","no")
      sHTML1 = sHTML1 & ofv.GenRow("","","","","","","","","no")

   end if





 IF SMAINAPL = "OR" OR SMAINAPL = "PL" THEN

    sHTML1 = sHTML1 & ofv.GenCelda("","UT  ","","","","","","&nbsp;","si")

'   HISTORIAL DE OBSERVACIONES -----------------
    if  Session("Segce") = "S" then

        t2 = "<a href='COOBSHIST.asp?"
        t2 = t2 & "Tipobjeto=" & sTipo & "&Codobjeto="
        t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
        t2 = t2 & sImagen2H & "</a>"
        SBOTTIT = "HISTORIAL&nbsp;de&nbsp;Observaciones. "
        sHTML1 = sHTML1 & ofv.GenCelda("","aT title='" & SBOTTIT & "' " & StyL & " ","","6","","","",t2,"si")

    end if

 ELSE

      scktAuto = "NO"
      scktObs  = "NO"

  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      sEsquema = rs(1)
      siconmas = rs("archivo")
      sPerfil  = rs(3)
      sAcceso  = rs(16)
      sResol   = rs(17)
      scktAuto = rs(19)
      scktObs  = rs(20)
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


  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" then
      if  bModificar then
          t2 = "<a href='edportada.asp?SKIND=1&Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPPRF"  & "' target=Principal>"
          t2 = t2 & sImagen0p & "</a>"
          SBOTTIT = "Editar&nbsp;Portada&nbsp;del&nbsp;esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")
      end if
  end if

  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "DESA"  then
      if  bModificar or bModificaresp then
          t2 = "<a href='EDperfilesdocProp.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPPRF"  & "' target=Principal>"
          t2 = t2 & sImagen0 & "</a>"
          SBOTTIT = "Editar&nbsp;perfil&nbsp;del&nbsp;esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")
      end if
  end if

  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED"  then
      if  bModificar then
          t2 = "<a href='EDPRINTPNL.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPPRF"  & "' target=Principal>"
          t2 = t2 & sImagen1 & "</a>"
          SBOTTIT = "Configura&nbsp;impresion&nbsp;del&nbsp;esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")
      end if

  ELSE


     sPrint = "N"

     Strsql = "Select * From OBJPRINTPNL where  CODOBJETO = '" & sCodigo & "' AND TIPOBJETO = '" & sTipo & "' "
 '' RESPONSE.WRITE STRSQL & "<BR>"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     if rs.eof then

        strsql = "Insert into OBJPRINTPNL ("
        strsql = strsql & " CODOBJETO,TIPOBJETO,IMPADMITE,IMPTIPO,IMPARCHIVO"
        strsql = strsql & " ) "
        strsql = strsql & " values ("
        strsql = strsql & "'" & sCodigo & "','" & sTipo & "','S','D',''"
        strsql = strsql & ")"
        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


        Strsql = "Select * From OBJPRINTPNL where  CODOBJETO = '" & sCodigo & "' AND TIPOBJETO = '" & sTipo & "' "
 '' RESPONSE.WRITE STRSQL & "<BR>"
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     END IF

     if not rs.eof then

        sADMITE     = (rs(3) & "")
        simptipo    = (rs(4) & "")
        simpARCHIVO = (rs(5) & "")

        IF  sADMITE     = "S" THEN
            sPrint = "S"

            IF  simptipo     = "DA" AND simpARCHIVO <> "" THEN
                sPrint = "SA"


            END IF

        END IF

     END IF


  end if




'RESPONSE.WRITE Session("Form2")
'IF  session("APLMDCLI")="BC" OR  session("APLMDCLI")="TARSHOP" THEN

  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND ( (sTipoCPT = "D" and LEFT(Session("Form2"),4) = "ELIM") OR (sTipoCPT = "D" and LEFT(Session("Form2"),4) = "PARA") OR  (sTipoCPT = "D" and LEFT(Session("Form2"),4) = "TEMA") OR  (sTipoCPT = "D" and LEFT(Session("Form2"),4) = "DESA") or (sTipoCPT = "C") ) then
      if  bModificar    or bModificaresp  then
          t2 = "<a href='edobjsdk.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal>"
          t2 = t2 & sImagen4 & "</a>"
          SBOTTIT = "Permisos&nbsp;del&nbsp;Esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxxx & " ","","6","","","",t2,"si")
      end if
  end if

'ELSE
'  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND ( (sTipoCPT = "D" and LEFT(Session("Form2"),4) = "DESA") or (sTipoCPT = "C") ) then
'      if  bModificar    or bModificaresp  then
'          t2 = "<a href='edobjsdk.asp?Tipobjeto=" & sTipo & "&Codobjeto="
'          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal>"
'          t2 = t2 & sImagen4 & "</a>"
'          SBOTTIT = "Permisos&nbsp;del&nbsp;Esquema."
'          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxxx & " ","","6","","","",t2,"si")
'      end if
'  end if
'END IF
  if  sFuncion = "E" and Session("Segct") = "S" and session("ModMain") = "ED" AND Session("CORUNQ") = "S" then
          t2 = "<a href='CoobjURL.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPSDK"  & "' target=Principal>"
          t2 = t2 & sImagen7 & "</a>"
          SBOTTIT = "Direccion&nbsp;del&nbsp;Esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")
  end if

  if  sFuncion = "E" and Session("Segce") = "S" then

      if  session("ModMain") = "WK"  then
          t2 = "<a href='coWKFREDIR.asp?sopcion=VER&"
      elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) = "PARA" then
          t2 = "<a href='coWKFREDIR.asp?sopcion=APR&"
      elseif session("ModMain") = "ED" AND LEFT(Session("Form2"),4) <> "PARA" AND NOT Session("SegcT") = "S" then
          t2 = "<a href='coWKFREDIR.asp?sopcion=VER2&"
      else
      	  if scktAuto = "NO" then
      	     t2 = "<a href='coWKFREDIR.asp?sopcion=NOVA&"
          else
             t2 = "<a href='edobjwfl.asp?"
          end if
      end if

      t2 = t2 & "PASO=AUTO&Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2 & "</a>"
      SBOTTIT = "Circuito&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;del&nbsp;Esquema."
      sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxxx & " ","","6","","","",t2,"si")

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
      if scktObs  = "NO" then
      	     t2 = "<a href='coWKFREDIR.asp?sopcion=NOVA&"
          else
             t2 = "<a href='edobjwfl.asp?"
          end if
      end if

      t2 = t2 & "PASO=OBS&Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2 & sImagen2O & "</a>"
      SBOTTIT = "Circuito&nbsp;de&nbsp;Observaciones. "
      sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxxx & " ","","6","","","",t2,"si")

  end if

' HISTORIAL DE OBSERVACIONES -----------------
if  sFuncion = "E" and Session("Segce") = "S" then

      t2 = "<a href='COOBSHIST.asp?"
      t2 = t2 & "Tipobjeto=" & sTipo & "&Codobjeto="
      t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPWFL"  & "' target=Principal>"
      t2 = t2 & sImagen2H & "</a>"
      SBOTTIT = "HISTORIAL&nbsp;de&nbsp;Observaciones. "
      sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")

  end if

'------------------------------
  if session("APLMDCLI")="BGAL" THEN
   XARBOL= (LEFT(Session("Form2"),4) = "DESA") OR (Session("Form2") = "PARA APROBAR")
  ELSE
   XARBOL= TRUE 'LEFT(Session("Form2"),4) = "DESA"
  END IF

  'RESPONSE.WRITE session("APLMDCLI") & " " & Session("Form2")
  if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED" AND XARBOL  then

  'if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED"   then
      if  bModificar or bModificaresp  then

          t2 = "<a href='edobjdist.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPDST" & "' target=Principal>"
          t2 = t2 & sImagen3 & "</a>"
          SBOTTIT = "Lista&nbsp;de&nbsp;distribuci&oacute;n."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")

      end if
  end if

  if  sFuncion = "Exx" and Session("Segct") = "S" and session("ModMain") = "ED"  then
      if  bModificar then

          t2 = "<a href='edGenEstad.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPGST"  & "' target=Principal>"
          t2 = t2 & sImagen5 & "</a>"
          SBOTTIT = "Estadisticas&nbsp;del&nbsp;Esquema."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")

     end if
  end if
  'response.write sFuncion & "       " & Session("Segct") &  LEFT(Session("Form2"),4) & bModificar
 ' if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED"   then

 if session("APLMDCLI")="BGAL" THEN
    XARBOL= (LEFT(Session("Form2"),4) = "DESA") OR (Session("Form2") = "PARA APROBAR")
   ELSE
    XARBOL= LEFT(Session("Form2"),4) = "DESA"
  END IF

  if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED" AND XARBOL  then
      if  bModificar or bModificaresp then

          t2 = "<a href='edobjACD.asp?Tipobjeto=" & sTipo & "&Codobjeto="
          t2 = t2 & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=" & "ESPDST" & "' target=Principal>"
          t2 = t2 & sImagen5 & "</a>"
          SBOTTIT = "Lista de accesos directos para el documento."
          sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")

      end if
  end if

''''Agregado para incorporar la opcion de "test"

    'response.write sFuncion & " sFuncion " & "<br>"
    'response.write Session("Segct") & " Session("Segct") " & "<br>"
    'response.write session("ModMain") & " session("ModMain") " & "<br>"


  if  sFuncion = "E" and Session("Segct") = "S"  and session("ModMain") = "ED" AND XARBOL then

''if  sFuncion = "E" and Session("Segco") = "S"  then
     ''and session("ModMain") = "ED" AND XARBOL
       'if  bModificar or bModificaresp then
            shref= " Consultar.asp?Codigo=" & sCodigo & "&Tipo=" & STIPO

            t2 = "<a class='bietx' href='"& shref
            t2 = t2 & "&GLBSOLN=" & "ESPDST" & "' target=_blank>"
      ''''''t2 = t2 & sImagen5 & "</a>"
            t2 = t2 & "TestGral" & "</a>"
            SBOTTIT = "Test General"
            sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","","6","","","",t2,"si")

       'end if
  end if
''''Agregado para incorporar la opcion de "test"

 END IF

if  (Session("Segco") = "S" and (sPrint = "S" OR sPrint = "SA") and session("ModMain") <> "ED") or  (  sFuncion = "E"  and session("ModMain") = "ED" ) then

      IF sPrint = "S" THEN
         t2 = "<a class='bietx' href='Reportes.asp?ID=" & sCodigo & "&Titulo=Esquemas_Normativos&perfil="
         t2 = t2 & sPerfil & "&tipobj=" & xtipobj & "&GLBSOLN=" & "ESPREP"  & "' Target=_blank>"
         t2 = t2 & "Impresion" & "</a>"


         SBOTTIT = "Vista&nbsp;previa&nbsp;de&nbsp;impresión&nbsp;del&nbsp;Esquema."
         sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","center","6","","","",t2,"si")

     ELSE

         t2 = "<a class='bietx' href='" & simpARCHIVO & "' Target=_blank>"
         't2 = t2 & sImagen1 & "</a>"
         t2 = t2 & "Impresion" & "</a>"

         SBOTTIT = "Archivo&nbsp;de&nbsp;impresión&nbsp;del&nbsp;Esquema."
         sHTML1 = sHTML1 & ofv.GenCelda("","bietx title='" & SBOTTIT & "' " & StyLxxx & " ","center","6","","","",t2,"si")



     END IF



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

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;border-style:solid;border-color:#ffffff;border-width:1;"
        else
            styl = XspEx("#888888","#cccccc",0) & ";border-style:solid;border-color:#ffffff;border-width:1;"
        end if

        sHTML = sHTML & ofv.GenTabla("panel","aaa style='" & styl & "' ","","100%","",0,0,"-2","-2","","","","no")
        sHTML = sHTML & ofv.GenRow("","","","","","","","","no")

        sHTML = sHTML & ofv.GenCelda("","ut","right","1%","","","","&nbsp;","si")
        sHTML = sHTML & ofv.GenCelda("","ut","","1%","","","",sHTML1,"si")

        sHTML = sHTML & ofv.GenCelda("","ut","","99%","","","","&nbsp;","si")
        sHTML = sHTML & ofv.GenFRow()
        sHTML = sHTML & ofv.GenFTabla()

   end if





  xSolasp = ""


  xSolasp = xSolasp & "BarraSol.asp"
  xSolasp = xSolasp & "?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector & "&MAINAPL=" & SMAINAPL
  xSolasp = xSolasp & "&Funcion=" & sFuncion & "&SSCLICK=" & sSClick
  xSolasp = xSolasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

    sHTML = sHTML & "<form name=festad  id=festad action='" & xSolasp  & "' "
    sHTML = sHTML & " method=post target='WK1' style='display:none;' >"
    sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
    sHTML = sHTML & "</form>"



  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
