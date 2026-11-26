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
  if session("Nomodificar") = "S" then bModificar = False

  sTipo     = request.querystring("Tipo")
  sCodigo   = request.querystring("Codigo")
  sFuncion  = request.querystring("Funcion")
  sSClick   = request.querystring("SSClick")
  sMenu     = request.querystring("Menu")
  sSector   = request.querystring("Sector")

  sPerfil      = request.querystring("Perfil")
  sComplemento = request.querystring("Complemento")
  sComplemento = ofv.convertircar(sComplemento,"_"," ")


''  response.write "recibe " & request.querystring & "<br>"

  set cn = ofv.conectar(ofv.strconn2)



  sHTML = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
  '   wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
 '    wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
     wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
  '   wxs = wxs & " } else { " & chr(13)
  '   wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
  '   wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
  '   wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "<body  >"


''''''''''''''''''''''''

  dim stabla(10)
  dim solord(10)

  solord(0)  = 0
  solord(1)  = 1
  solord(2)  = 2
  solord(3)  = 3
  solord(4)  = 4
  solord(5)  = 5
  solord(6)  = 6
  solord(7)  = 7
  solord(8)  = 8
  solord(9)  = 9
  solord(10) = 10


  dim Nombres(10)


  dim ASPs(10)

 if sFuncion <> "C" then

  Nombres(0) = "WorkFlow"
  Nombres(1) = "General"
  Nombres(2) = "Propiedades"
  Nombres(3) = "Estructura"
  Nombres(4) = "Detalle"
  Nombres(5) = "Imagenes"
  Nombres(6) = "Compuesto"
  Nombres(7) = "Archivos"
  Nombres(8) = "N"
  Nombres(9) = "N"
  Nombres(10) = "N"

  ASPs(0) = "edGenKomp.asp"
  ASPs(1) = "edGenKomp.asp"
  ASPs(2) = "N"
  ASPs(3) = "edEstrKomp.asp"
  ASPs(4) = "N"
  ASPs(5) = "edImagen2.asp"
  ASPs(6) = "edCompuesto2.asp"
  ASPs(7) = "edimgarc2.asp"
  ASPs(8) = "N"
  ASPs(9) = "N"
  ASPs(10) = "N"

 else

  '''  sAccion = "imagencont2.asp?ID=" & ID

  Nombres(0) = "WorkFlow"
  Nombres(1) = "Imagenes"

  Nombres(2) = "General"
  Nombres(3) = "Estructura"
  Nombres(4) = "N"
  Nombres(5) = "N"
  Nombres(6) = "Compuesto"
  Nombres(7) = "Archivos"
  Nombres(8) = "N"
  Nombres(9) = "N"
  Nombres(10) = "N"

  ASPs(0) = "imagencont2.asp"
  ASPs(1) = "imagencont2.asp"
  ASPs(2) = "imagencontG.asp"
  ASPs(3) = "imagencont3.asp"
  ASPs(4) = "N"
  ASPs(5) = "N"
  ASPs(6) = "N"
  ASPs(7) = "imagencont4.asp"
  ASPs(8) = "N"
  ASPs(9) = "N"
  ASPs(10) = "N"

 end if



  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Comperfil Where tipoperfil='" & sPerfil & "'"
'''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  redim Solapas(1)
  redim Solapas2(1)
  Solapas(1) = Nombres(1)
  Solapas2(1) = ASPs(1)
  x = 1
  if  not rs.eof then
  for z = 2 to 6
    if z <> 4 and ((z <> 2 and sFuncion <> "C") or (sFuncion = "C") )  then
      if rs(z) Then
        x = x + 1
        redim preserve Solapas(x)
        redim preserve Solapas2(x)
        Solapas(x) = Nombres(z)
        Solapas2(x) = ASPs(z)
        if z = 5 then
           x = x + 1
           redim preserve Solapas(x)
           redim preserve Solapas2(x)
           Solapas(x) = Nombres(7)
           Solapas2(x) = ASPs(7)
        end if
      end if
    end if
  next
  end if

  call ofv.cerrarconsulta(rs)

''''''''''''''''''''''''''''''''''''''
  if session("ModMain") = "WK" and sFuncion = "E" then
     Solapas(0)   = "WORKFLOW"
     Solapas2(0)  = ASPs(0)
     xstrsol      = 0

  else
     xstrsol      = 1

  end if

  xfinal = ubound(Solapas)
  if  Session("Segco") <> "S" then
      xfinal = 1
  end if

  xSolasp = ""
  xSolasp = xSolasp & "BarraCompSol.asp?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
  xSolasp = xSolasp & "&Funcion=" & sFuncion & "&Menu=" & sMenu & "&Sector=" & sSector & "&SSCLICK="

 r1 = ""


     If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
        xGenConta = ofv.GetTit(clng(ofv.ModTitCls) + 5)
        xGenConta = ofv.FTitDat1

        xsclase = "xts"

     else



        xsclase = "xts"

     End If

      xsimagenf  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >&#0167;</font>&nbsp;&nbsp;"
      xsimagend  = "<font FACE=corflaticon COLOR='#ffffff' SIZE=2 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#ffffff';"" >0</font>&nbsp;&nbsp;"


  for x = xstrsol to xfinal
      if solapas(solord(x)) <> "N" then
         sSolapa =  solapas(solord(x))

         If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
            sInput = replace(Ucase(sSolapa),"*"," ")
            xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"
         else
            sInput = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
            xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"
         End If

         t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
         t3 = t3 & sInput
         t3 = t3 & "</A>"



         t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

         r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")




         if clng(sSClick) = clng(solord(x)) then

           if sFuncion <> "C" then

            xverasp = solapas2(solord(x)) & "?Tipocomp=" & sTipo & "&Codcomp=" & sCodigo
            xverasp = xverasp & "&Menu=" & sMenu & "&Sector=" & sSector
            xverasp = xverasp & "&GLBSOLN=S" & solord(x)
           else
            xverasp = solapas2(solord(x)) & "?id=" & sCodigo
           end if

            sHTML1 = ""
            sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "
            sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
            sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
            sHTML1 = sHTML1 & "</form>" & chr(13)

            sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
            sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
            sHTML1 = sHTML1 & "</script>" & Chr(13)


        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

          xlfun = " onclick=""document.getElementById('MDBEDL').style.display='block'; document.getElementById('DBEDTD1').style.display='none'; document.getElementById('DBEDTD2').style.display='block'; document.getElementById('MM2').style.display='block'; "" "
          xlfun2 = " onclick=""document.getElementById('MDBEDL').style.display='none'; document.getElementById('DBEDTD1').style.display='block'; document.getElementById('DBEDTD2').style.display='none'; document.getElementById('MM2').style.display='none'; "" "

           macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

          xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


          macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
          macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

           marow = ofv.GenRow("","","","","20","","",macel,"si")

           sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si") & chr(13)



         end if


      end if
  next



        sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >" & chr(13)

        IF  Session("FULLGUI") = "N" THEN
            styl = "background-color:#888888;"
        else
            styl = XspEx("#888888","#cccccc",0) & ";"
        end if

            styl = ""

        If ofv.ModTitCls = "1" Then

           sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si") & chr(13)



        else

           sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si") & chr(13)


        end if

           sHTML = sHTML & "</div>" & chr(13)



  sHTML = sHTML & sHTML1

  call ofv.cerrarconn(cn)

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " IniWin(); "
     wxs = wxs & "</script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
