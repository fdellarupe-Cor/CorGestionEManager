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
  sNIV1     = request.querystring("NIV1")
  sNIV2     = request.querystring("NIV2")
  sNIV3     = request.querystring("NIV3")
  sNIV4     = request.querystring("NIV4")
  sNIV5     = request.querystring("NIV5")

  sPerfil      = request.querystring("Perfil")
  sComplemento = request.querystring("Complemento")
  sComplemento = ofv.convertircar(sComplemento,"_"," ")

''  response.write "recibe " & request.querystring & "<br>"

 wsSClick = "V" & sSClick

 if  sSClick = "" then     sSClick      = 99


  set cn = ofv.conectar(ofv.strconn2)



  sHTML = "<html>"

     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
     wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
     wxs = wxs & " } else { " & chr(13)
     wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
     wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
     wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "<body onload='IniWin();' >"


  dim stabla(20)
  dim solord(20)

  solord(0) = 0
  solord(1) = 15
  solord(2) = 1
  solord(3) = 6
  solord(4) = 3
  solord(5) = 2
  solord(6) = 4
  solord(7) = 5
  solord(8) = 7
  solord(9) = 8
  solord(10) = 11
  solord(11) = 10
  solord(12) = 9
  solord(13) = 12
  solord(14) = 13
  solord(15) = 14
  solord(16) = 16
  solord(17) = 17
  solord(18) = 18
  solord(19) = 19
  solord(20) = 20






  dim Nombres(20)

  dim ASPs(20)

  ASPs(0) = "imagencont2.asp?skind=1"


  if sFuncion = "E" then
     ASPs(1) = "edGenKomp.asp?skind=1"
  else
     ASPs(1) = "imagencontG.asp?skind=1"
  end if

  if sFuncion = "E" then
     ASPs(2) = "edEstrKomp.asp?skind=1"
  else
     ASPs(2) = "imagencont3.asp?skind=1"
  end if

  ASPs(3) = "N"

  if sFuncion = "E" then
    ASPs(4) = "edImagen2.asp?skind=1"
  else
    ASPs(4) = "imagencont2.asp?skind=1"
  end if

  if sFuncion = "E" then
    ASPs(5) = "edCompuesto2.asp?skind=1"
  else
    ASPs(5) = "ImagenIndex.asp?skind=1"
  end if




  if sFuncion = "E" then
    ASPs(6) = "edimgarc2.asp?skind=1"
  else
    ASPs(6) = "imagencont4.asp?skind=1"
  end if



  ASPs(7)  = "N"
  ASPs(8)  = "N"
  ASPs(9)  = "N"
  ASPs(10) = "N"
  ASPs(11) = "N"
  ASPs(12) = "N"
  ASPs(13) = "N"
  ASPs(14) = "N"
  ASPs(15) = "N"

  perfilsp = "S"







  siconmas = "ICorAcciones.jpg"



  dim Solapas(20)
  dim Solapas2(20)

  for z=0 to 20
      Solapas(z) = "N"
  next



  Strsql = "Select * From comperfil Where tipoperfil='" & sPerfil & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if Session("Segco") <> "S" then
    for xx = 1 to 20
        Nombres(xx) = "N"
    next

     Solapas(15)  = REPLACE(RS(31)," ","*")
     Solapas2(15) = ASPs(15)
     xstrsol      = 1
     xfinal       = 1
     sSClick      = 99


  else

    for xx = 1 to 20
        solord(xx) = 17
        Nombres(xx) = "N"
    next



    for xx = 2 to 16
        stabla (xx - 1) = rs(xx + 30)
    next



    yy = 0
    for xx = 1 to 15

        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           solord (stabla (xx)) = xx
        end if
    next
    tyy = yy




  if session("ModMain") = "WK"  then
     Solapas(0)  = "WORKFLOW"
     Solapas2(0) = ASPs(0)
     xstrsol     = 0
  else
     xstrsol     = 1
  end if




  X=0
  for z = 2 to 16

    if rs(z) Then

      if (sFuncion = "E" and z <> 4  ) or (sFuncion = "C" and z <> 21 and z <> 4 and z <> 61 )  then
          x = x + 1
          Solapas(z-1)  = REPLACE(RS(z + 15)," ","*")
          Solapas2(z-1) = ASPs(z-1)




      end if
    end if

  next


  xfinal = ubound(Solapas)

 end if

 ''' response.write Solapas2(0) & xestado & sFuncion & "<br>"

  call ofv.cerrarconsulta(rs)


  xSolasp = ""
  xSolasp = xSolasp & "BarraCompSolold.asp?Tipo=" & sTipo & "&Codigo="
  xSolasp = xSolasp & sCodigo & "&Menu=" & sMenu & "&Sector=" & sSector
  xSolasp = xSolasp & "&Funcion=" & sFuncion
  xSolasp = xSolasp & "&Perfil=" & sPerfil & "&Complemento=" & sComplemento
  xSolasp = xSolasp & "&NIV1=0&NIV2=0&NIV3=0&NIV4=0&NIV5=0"
  xSolasp = xSolasp & "&SSCLICK="

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

 wsSClick2 = sSClick

  xsolini = 99


  for x = xstrsol to xfinal
      if solapas(solord(x)) <> "N" then
         sSolapa =  solapas(solord(x))

         if  xsolini = 99 then  xsolini = solord(x)

         if  sSClick = 99 then  sSClick = xsolini

         If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then


            sInput  = replace(Ucase(sSolapa),"*"," ") & "&nbsp;"
            xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"


         else

            sInput  = "&nbsp;" & replace(Ucase(sSolapa),"*","&nbsp;") & "&nbsp;"
            xsclase = "at style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;'"

         End If


        if session("WEBDESK") = "S" then


           sInput  = "<font style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;' >" & replace(Capitalize(replace(Ucase(sSolapa),"*"," "))," ","&nbsp;") & "</font>"
           sLink   = OFV.GenLink("", "cmb  style='font-size:8pt;cursor:pointer;text-decoration:none;color:#993333;font-weight:bold;'  ", xSolasp & solord(x), "WK1", "", "", sInput, "si")
           Celda   = Celda & wwwsaltador & "<span ><font color='#ff0000' size=2 >&#0124;</font>&nbsp;&nbsp;" & sLink & "&nbsp;&nbsp;</span> "

           wwwestilo = " style='background-color:#f0f0f0;border-LEFT-style:solid;border-LEFT-width:1;border-LEFT-color:#FF0000;padding-left:10px;' "

           if (solord(x)) = 0 then
              wwwestilo = " style='background-color:#ffffff;border-style:solid;border-width:2;border-color:#990000;padding-left:10px;' "
           end if


           CTABELEN = ofv.GenCelda("","uT  " & wwwestilo & " ","","","","","",sLink,"si")
           FTABELEN = FTABELEN & ofv.GenRow("","","","","20","","",CTABELEN,"si")



           if sFuncion = "E" or sFuncion = "ED"  then

           else

           end if


        ELSE


         t3 = xsimagenf & "<a  href='" & xSolasp & solord(x) & "' target='WK1' >"
         t3 = t3 & sInput
         t3 = t3 & "</A>"



         t4 = ofv.GenCelda("",xsclase,"","","","","",t3,"si")

         r1 = r1 & ofv.GenRow("","","","","20","","",t4,"si")

        end if

         if clng(sSClick) = clng(solord(x)) then

            xverasp = solapas2(solord(x)) & "&Tipocomp=" & sTipo & "&Codcomp=" & sCodigo & "&id=" & sCodigo
            xverasp = xverasp & "&Menu=" & sMenu & "&Sector=" & sSector
            xverasp = xverasp & "&GLBSOLN=S" & solord(x)
            xverasp = xverasp & "&NIV1=" & sNIV1  & "&NIV2=" & sNIV2  & "&NIV3=" & sNIV3  & "&NIV4=" & sNIV4  & "&NIV5=" & sNIV5

            sHTML1 = "" & chr(13)
            sHTML1 = sHTML1 & "<form name=SSOLACT action='" & xverasp & "' "
            sHTML1 = sHTML1 & " method=post target='Principal' style='display:none;' >" & chr(13)
            sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
            sHTML1 = sHTML1 & "</form>" & chr(13)

            sHTML1 = sHTML1 & "<script language=JavaScript>" & Chr(13)
            sHTML1 = sHTML1 & " window.top.CorSol='" & replace(Ucase(sSolapa),"*"," ") & "'; " & Chr(13)
            sHTML1 = sHTML1 & "</script>" & Chr(13)


         if session("WEBDESK") = "S" then




         ELSE

           IF  Session("FULLGUI") = "N" THEN
               styl = "background-color:#888888;"
           else
               styl = XspEx("#888888","#cccccc",0) & ";"
           end if

            xlfun = " onclick=""document.all.MDBEDL.style.display='block'; document.all.DBEDTD1.style.display='none'; document.all.DBEDTD2.style.display='block'; document.all.MM2.style.display='block'; "" "
            xlfun2 = " onclick=""document.all.MDBEDL.style.display='none'; document.all.DBEDTD1.style.display='block'; document.all.DBEDTD2.style.display='none'; document.all.MM2.style.display='none'; "" "

            macel = ofv.GenCelda("","at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;padding-left:3px;' " & xlfun,"","90%","","","",replace(Ucase(sSolapa),"*"," "),"si")

            xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#cc0000;cursor:pointer;"


            macel = macel & OFV.gencelda("DBEDTD1", xsclase2 & "display:block;' ", "center title='Ver solapas.' " & xlfun, "center", "1", "", "", "1", "si")
            macel = macel & OFV.gencelda("DBEDTD2", xsclase2 & "display:none;' ", "center title='Ocultar Solapas.' " & xlfun2, "center", "1", "", "", "0", "si")

            marow = ofv.GenRow("","","","","20","","",macel,"si")

            sHTML = sHTML & OFV.GenTabla("DBED", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", marow, "si")

         end if

        end if


      end if
  next

        if session("WEBDESK") = "S" then


        '''    sHTML = sHTML & Celda
       '''     sHTML = sHTML & "<span ><font color='#ff0000' size=2 >&#0124;</font></span>"

            sHTML = sHTML & "<div id='WKTBSP1'  >"
            sHTML = sHTML & OFV.GenTabla("WKTBSP1TT", "aaa  ", "", "100%", "", "0", "", "1", "1", "", "", FTABELEN, "si")
            sHTML = sHTML & "</div>"

        ELSE


           sHTML = sHTML & "<div id='MDBEDL' style='display:none;height:240px;overflow:auto;" & WStyl & "' >"

           IF  Session("FULLGUI") = "N" THEN
               styl = "background-color:#888888;"
           else
               styl = XspEx("#888888","#cccccc",0) & ";"
           end if

            styl = ""

           If ofv.ModTitCls = "1" Then

             sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")

       ''    sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,-2,-2,"","",r1,"si")

           else

             sHTML = sHTML & OFV.GenTabla("DBEDL", "aaa style='border-style:solid;border-color:#ffffff;border-width:1;" & styl & "' ", "", "100%", "", "0", "", "-2", "-2", "", "", r1, "si")
        ''   sHTML = sHTML & ofv.GenTabla("DBED","","","90%","",0,0,2,2,"","",r1,"si")

           end if

           sHTML = sHTML & "</div>"

       end if

  sHTML = sHTML & sHTML1

  ' sHTML = sHTML & armasolapaComp(sCodigo,sTipo,"")
    sHTML = sHTML & "<div id=contm>" & armasolapaComp(sCodigo, sTipo,sFuncionSol) & "</div>"
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"


end if

Response.Write sHTML

%>
