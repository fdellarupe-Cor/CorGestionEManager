<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()





      Response.expires = 0

      sOpcion = Request.querystring("Opcion")

        If Session("USRES") <> "B" Then XFondo = " background='CorFondos/" & Session("CMIMENU") & "' "

      sHTML = "<html>"
      sHTML = sHTML & "<style><link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"">"

      shtml = shtml & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/line-awesome/css/line-awesome.css"" />"
	  shtml = shtml & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/line-awesome/css/line-awesome-font-awesome.css"" />"

      shtml = shtml & "</STYLE>"
      sHTML = sHTML & "<body bgcolor='#dddddd' >"

      Celda = ""

      Formulario = ""



      If sOpcion = "I" Then

       '   Fila = OFV.genrow("", "", "", "", "", "", "", Celda, "si")
       '   sHTML = sHTML & OFV.GenTabla("", "", "", "100%", "100%", "", "", "-1", "-2", "", "", Fila, "si")

      Else
            XFondo = "N"



          Dim aAplicacion()
            Dim aLinks()
            Dim aLinksImg()
            Dim aLinksLet()
            Dim aLinksTLT()
            Dim aLinksURL()
            Dim aLinksTRG()
            Dim aLinksDSC()

            Set cx = OFV.conectar(OFV.strconn1)
            StrSql = "select * from CorAplic"
            Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
            If Not rx.EOF Then
               If rx.fields.Count > 8 Then
                  Call OFV.cerrarconsulta(rx)
                  StrSql = "select * from CorAplic where enable <> 0 order by ordenapli "
                  Set rx = OFV.crearconsultaEx(StrSql, cx, 1, "")
                  ind = 0
                  Do Until rx.EOF
                     ReDim Preserve aAplicacion(ind)
                     ReDim Preserve aLinks(ind)
                     ReDim Preserve aLinksImg(ind)
                     ReDim Preserve aLinksLet(ind)
                     ReDim Preserve aLinksTLT(ind)
                     ReDim Preserve aLinksURL(ind)
                     ReDim Preserve aLinksTRG(ind)
                     ReDim Preserve aLinksDSC(ind)

                     aAplicacion(ind) = rx(1)
                     aLinks(ind) = "<a class=bt title='" & rx(12)
                     aLinks(ind) = aLinks(ind) & "' href='../COREss20/" & rx(11)
                     If UCase(rx(11)) = UCase("Menus.asp") Then
                        aLinks(ind) = aLinks(ind) & "?Opcion=" & UCase(rx(1))
                     End If
                     aLinks(ind) = aLinks(ind) & "' target='" & rx(10)
                     aLinks(ind) = aLinks(ind) & "' >"
                     aLinksImg(ind) = "<img border=0 src='../COREss20/Imagenes/" & rx(6) & "' >"
                     apllet = Replace(rx(2), "*", "")
                     aLinksLet(ind) = "<b>" & Replace(apllet, " ", "&nbsp;") & "</b>"

'                     aLinks(ind) = aLinks(ind) & "</a>"

                     aLinksTLT(ind) = rx(12)
                     aLinksURL(ind) = "../COREss20/" & rx(11)
                     If UCase(rx(11)) = UCase("Menus.asp") Then
                        aLinksURL(ind) = aLinksURL(ind) & "?Opcion=" & UCase(rx(1)) & "&INICIA=S"

                     else
                        aLinksURL(ind) = rx(11)
                     End If
                     aLinksTRG(ind) = rx(10)
                     aLinksDSC(ind) = ucase(Replace(rx(2), "*", ""))

                     ind = ind + 1
                     rx.movenext
                  Loop


               Else

            ReDim Preserve aAplicacion(4)
          aAplicacion(0) = "CORESS20"
          aAplicacion(1) = "CORCHT20"
          aAplicacion(2) = "COREP20"
          aAplicacion(3) = "COREM20"
          aAplicacion(4) = "COREC20"

          ReDim Preserve aLinks(4)
          aLinks(0) = "<a title='COR&nbsp;ESS&nbsp;2.0' href=Menus.asp?Opcion=CORESS20 target=Menus><img border=0 src=Imagenes/iCorESS.jpg></a>"
          aLinks(1) = "<a title='COR&nbsp;Chat&nbsp;2.0' href=../CORChat20/chat.htm target=_blank><img border=0 src=Imagenes/iCorChat2.jpg></a>"
          aLinks(2) = "<a title='COR&nbsp;EPlanning&nbsp;2.0' href=Menus.asp?Opcion=COREP20 target=Menus><img border=0 src=Imagenes/iCorplanning.jpg></a>"
          aLinks(3) = "<a title='COR&nbsp;EManager&nbsp;2.0' href=Menus.asp?Opcion=COREM20 target=Menus><img border=0 src=Imagenes/iCormanager.jpg></a></td>"
          aLinks(4) = "<a title='COR&nbsp;EComponent&nbsp;2.0' href=Menus.asp?Opcion=COREC20 target=Menus><img border=0 src=Imagenes/iCorComponent.jpg></a>"

            End If
            End If
            Call OFV.cerrarconsulta(rx)

        if  session("SMPLMENU") <> "S" then


        If Session("USRES") <> "xB" Then
           sInput = "<a title='Ayuda' href=../COREss20/Ayuda/NormasHelp.asp target=info><img border=0 src='../COREss20/Imagenes/iCorSHelp.jpg' ></a>"
         '  Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a title='About' href=CorAbout.asp target=info><img border=0 src='../COREss20/Imagenes/iCorAbout.jpg' ></a>"
         '  Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a title='Usuario' href=InfoUser.asp target=info><img border=0 src='../COREss20/Imagenes/iCorSInfoUser.jpg' ></a>"
         '  Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
        Else
           sInput = "<a  class=bt title='Ayuda' href=../COREss20/Ayuda/NormasHelp.asp target=info><b>Ayuda</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a  class=bt title='About' href=CorAbout.asp target=info><b>Acerca&nbsp;de</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")
           sInput = "<a  class=bt title='Usuario' href=InfoUser.asp target=info><b>Usuario</b></a>"
           Celda = Celda & OFV.gencelda("", "bt", "", "", "", "", "", sInput, "si")

        End If

        end if

          xcount = 0
          scp    = ""

          Set cn = OFV.conectar(OFV.strconn0)

          Fila2 = ""

          For n = 0 To UBound(aAplicacion)

               atrok = "N"
              StrSql = "Select Alias,usrEstado From atrmXusuario Where Tipo = 'A' "
              StrSql = StrSql & " And usuario = '" & Session("segperfil") & "' And "
              StrSql = StrSql & "  Aplicacion = '" & aAplicacion(n) & "'"
              Set rs1 = OFV.crearconsultaEx(StrSql, cn, 1, "")
              If rs1.EOF Then
                 atrok = "O"
              ElseIf rs1(1) Then
                     atrok = "S"
                  Else
                     atrok = "N"
              End If
              Call OFV.cerrarconsulta(rs1)

              If atrok = "O" Then


                 StrSql = "Select Alias,grpEstado From atrmXgrupo Where Tipo = 'A' "
                 StrSql = StrSql & " And  "

                 if  Session("DELLOG") = "S" then

                     XMFTGRP = split(session("SegGrpLst"),chr(9))

                     xcount = ubound(XMFTGRP) + 1
                     XMFSEP  = ""
                     StrSql = StrSql & " grupo in ("

                     for XMFI = 0 to ubound(XMFTGRP)
                         StrSql  = StrSql  & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
                         XMFSEP  = ","
                     next

                     StrSql = StrSql & ") "

                 else
                     StrSql = StrSql & " grupo = '" & Session("seggrp") & "' "
                 end if


                 StrSql = StrSql & " And Aplicacion = '" & aAplicacion(n) & "'"
                 Set rs1 = OFV.crearconsultaEx(StrSql, cN, 1, "")
                 If  rs1.EOF Then
                        atrok = "O"
                 Else
                     if  Session("DELLOG") = "S" then
                         XMFSUM = 0
                         XMFSUMN = 0
                         do until rs1.eof
                            XMFSUM = XMFSUM + 1
                            If rs1(1) Then
                               atrok = "S"
                               exit do
                            Else
                               atrok   = "N"
                               XMFSUMN = XMFSUMN + 1
                            End If
                         rs1.movenext
                         loop

                         if  atrok = "N" then
                             if  xcount >  XMFSUMN then atrok = "S"
                         end if

                     else
                         If rs1(1) Then
                            atrok = "S"
                         Else
                            atrok = "N"
                         End If
                     end if
                 End If
                 Call OFV.cerrarconsulta(rs1)

              End If




              If atrok = "O" Or atrok = "S"  Then
                            If XFondo = "N" Then
                               XFondo = n
                               XFdsc  = aLinksDSC(n)
                           '    sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
                           '    Formulario = OFV.GenForm("initaplx", "aa style='display:none;' ", "../COREss20/Menus.asp?Opcion=" & aAplicacion(n), "Menus", "post", "", sInput, "si")
                           '    scp = "<script language='javascript'> document.initaplx.submit(); </script>"
                            End If


                            if  sopcion = aAplicacion(n) then
                                XFondo = n
                                XFdsc  = aLinksDSC(n)
                            end if



                 If Session("USRES") <> "xB" Then
                    sInput = aLinks(n) & aLinksImg(n) & "</a>"
                 Else
                    sInput = aLinks(n) & aLinksLet(n) & "</a>"
                 End If

                 xsclase = "at "
                 xsclase1 = xsclase & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;padding-left:3px;' "

                 if session("WEBDESK") = "S" then
                    xsclase = "nt style='color:#888888;' "
                    xsclase1 = xsclase & " style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;padding-left:5px;' "
                 end if

        if  session("SMPLMENU") <> "S" then



                 sinput1 = "<a class=" & xsclase & " title='" & aLinksTLT(n) & "' href='../COREss20/footer.asp?Opcion="
                 sinput1 = sinput1 & aAplicacion(n) & "' target='Footer' >-</a>"

                 sinput2 = "<a  class=" & xsclase & "  title='" & aLinksTLT(n) & "' href='../COREss20/footer.asp?Opcion="
                 sinput2 = sinput2 & aAplicacion(n) & "' target='Footer' >" & aLinksDSC(n) & "</a>"

                 xlfun = " onclick=""document.getElementById('DFTRL').style.display='none' "" "


                 if session("WEBDESK") = "S" then
                    Celda2 = ""
                    sinput2 = "&nbsp;&nbsp;" & sinput2
                 else
                    Celda2 = OFV.gencelda("", xsclase, "center " & xlfun, "2", "", "", "", sInput1, "si")
                 end if


                 Celda2 = Celda2 & OFV.gencelda("", xsclase1 & xlfun, "", "", "", "", "", sInput2, "si")

                 Fila2 = Fila2 & OFV.genrow("", "", "", "", "20", "", "", Celda2, "si")

           end if


              End If

             Next

                 xsclase = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "
                 xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:Corflaticon;font-size:10pt;color:#ffffff;cursor:pointer;"

                 if session("WEBDESK") = "S" then
                    xsclase = "ut style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;color:#0000aa;cursor:pointer;' "


                    xsclase2 = "ut style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:sblock;font-family:CorFlaticon;font-size:10pt;color:#0000aa;cursor:pointer;"

                 end if

        if  session("SMPLMENU") <> "S" then


         '   xsimagena  = "<b><font FACE=CorFlaticon COLOR='#0000aa' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa';"" >s</font></b>"
          '  xsimagenb  = "<b><font FACE=CorFlaticon COLOR='#0000aa' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa';"" >&#57387;</font></b>"
          '  xsimagenc  = "<b><font FACE=CorFlaticon COLOR='#0000aa' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa';"" >&#57369;</font></b>"

             xsimagena  = "<b><font FACE=CorFlaticon COLOR='#0000aa' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa';"" >s</font></b>"
			 xsimagenb  = "<b><font FACE=CorFlaticon COLOR='#0000aa' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa';"" >&#0128;</font></b>"
			 'xsimagenb  = "<b style='font-family:CorFlaticon;' onmouseover=this.color='#bb0000'; onmouseout=this.color='#0000aa';>&#0128;</b>"
			 xsimagenc  = "<b><font FACE=CorFlaticon COLOR='#0000aa' SIZE=1 onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa';"" >r</font></b>"

        '''    xsclase = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;' "

            xlfun = " onclick=""document.getElementById('DFTRL').style.display='block';document.getElementById('DFTRTD1').style.display='none';document.getElementById('DFTRTD2').style.display='block'; "" "
            xlfun2 = " onclick=""document.getElementById('DFTRL').style.display='none';document.getElementById('DFTRTD1').style.display='block';document.getElementById('DFTRTD2').style.display='none'; "" "

            Celda = OFV.gencelda("", "at style='padding-left:3px;filter: chroma(color=#ffffff);' ", "", "2", "", "", "", "<IMG  valign=middle border=0 src='../coress20/imagenes/logocorP4ico.gif' style='' >", "si")

            Celda = Celda & OFV.gencelda("", xsclase & xlfun, "", "90%", "", "", "", "&nbsp;" & XFdsc, "si")

'////////////////////// agregado buscador ///////////////////////////




'////////////////////// hasta aca buscador ///////////////////////////

            if session("WEBDESK") = "S" then
               xlfun3 = "style='cursor:pointer ;' onclick='document.nulo.submit();' onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='#0000aa'; "" "
               xsaccion = "../coremanager/novedmenu.asp"
               xssInput = OFV.GenerarInput("boton", "", "hidden", "", "1", "it", "")
               xsFormulario = OFV.GenForm("nulo", " aa style='display:none;' ", xsaccion, "info", "post", "", xssInput, "si")

      if Session("UsrCorp") then

			   Celda = Celda & OFV.gencelda("","","", "2%", "", "", "",selecEmpre() ,"si")
	 else
	  ' Celda = Celda & OFV.gencelda("","","", "2%", "", "", "","&nbsp;" ,"si")
     end if

               Celda = Celda & OFV.gencelda("", xsclase & xlfun3, "", "2%", "", "", "", "&nbsp;Home&nbsp;", "si")

            end if


           '' xsclase = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#ffffff;cursor:pointer;' "
         ''   xsclase2 = "at style='border-right-style:solid;border-right-width:1;border-right-color:#dddddd;display:block;font-family:corflaticon;font-size:10pt;color:#ffffff;cursor:pointer;"

            '128214

            Celda = Celda & OFV.gencelda("DFTRTD1", xsclase2 & "display:xblock;' ", "center title='Lista de aplicaciones.' " & xlfun, "", "1", "", "", "2", "si")
            Celda = Celda & OFV.gencelda("DFTRTD2", xsclase2 & "display:none;' ", "center title='Ocultar aplicaciones.' " & xlfun2, "", "1", "", "", "&#592;", "si")


           ' Celda = Celda & OFV.gencelda("DFTRTD2", xsclase2 & "display:none;' ", "center title='Ocultar aplicaciones.' " & xlfun2, "", "1", "", "", "0", "si")

       '    Celda = Celda & OFV.gencelda("", xsclase2, "center title='Ayuda.' onclick='mostrarHLP();' STYLE='display:block;' ", "", "1", "", "", xsimagena, "si")
            Celda = Celda & OFV.gencelda("", xsclase2, "center title='Usuario.' onclick='mostrarUSR();' STYLE='display:xblock;' ", "", "1", "", "", xsimagenb , "si")
            Celda = Celda & OFV.gencelda("", xsclase2, "center title='Logoff.' onclick='document.logoff.submit();return false;' STYLE='display:xblock;' ", "", "1", "", "", xsimagenc, "si")


            Fila  = OFV.genrow("", "", "", "", "25", "", "", Celda, "si")

        end if


          If XFondo <> "N" and xfondo <> "" Then

             sInput = OFV.GenerarInput("boton2", "menu", "hidden", "", "10", "", "")
             Formulario = OFV.GenForm("initaplx", "aa style='display:none;' ",aLinksURL(XFondo), "'" & aLinksTRG(XFondo) & "'", "post", "", sInput, "si")
             scp = "<script language='javascript'> document.initaplx.submit(); </script>"
          End If



          if  session("SMPLMENU") <> "S" then

            if session("WEBDESK") = "S" then

               styl  = XspEx("#cccccc","#f9f9f9",0) & ";border-bottom-style:solid;border-bottom-color:#eeeeee;border-bottom-width:1;"
               styl2 = "background-color:#f0f0f0;border-style:solid;border-color:#888888;border-width:1;"


            else


              IF  Session("FULLGUI") = "N" THEN
                  styl = "background-color:#888888;border-style:solid;border-color:#ffffff;border-width:1;"
                  styl2 = ""
              else
                  styl = XspEx("#888888","#cccccc",0) & ";border-style:solid;border-color:#ffffff;border-width:1;"
                  styl2 = ""
              end if

            end if

              sHTML = sHTML & OFV.GenTabla("DFTR", "aaa   style='" & styl & "' ", "", "100%", "", "0", "", "-2", "-1", "", "", Fila, "si")

              styl = ""

              sHTML = sHTML & OFV.GenTabla("DFTRL", "aaa style='display:none;" & styl2 & "' ", "", "100%", "", "0", "", "-2", "1", "", "", Fila2, "si")
          end if

      End If


      sHTML = sHTML & Formulario & scp & "</body></html>"



function selecEmpre()


  if not Session("UsrCorp") then
  selecEmpre="&nbsp;"
  else

  IF  Session("EMPMUN") = "S" THEN

  XXXX = NEWEMPRE("NO")

    'session("ASP")="novedmenu.ASP"

  '  sEMPRE = request.querystring("EMPRE")
'
  '  IF SEMPRE = "" THEN
   '    SEMPRE = ATREMP(0,0)
   ' ELSE
   '    XXXX = NEWEMPRE(SEMPRE)
'  END IF

	sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
	sImagen4si = "<B><FONT FACE=CorFlatIcon COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenI3 = "<B><FONT FACE=""CorFlatIcon"" COLOR=#6699CC SIZE=2>=</FONT></B>"

    SINPUT3 = "&nbsp;<A class=eee" & xsclase & " title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='window.top.shwDem2();'>Empresa</a>"

    'Filac    = ""
    'Celdac   = ""
    'sHref   = ""

    'sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
    'sLink   = ofv.GenLink("","",sHref,"ftemas","","",sImage,"si")
    Celdac   = ofv.GenCelda("","","","8%","","","",SINPUT3,"si")
    'Celdac   = Celdac & ofv.GenCelda(""," ","center","5%","","","",sLink,"si")
    'Celdac   = Celdac & ofv.GenCelda("","  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

    Filac    = ofv.GenRow("","","top","","","","",Celdac,"si")

    Tablac   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Filac,"si")

  '----------------------------------------------
    sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
    sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & " window.top.shwDem2();" & chr(34)
    sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
    TABLAc = Tablac & "<DIV id='DEM12' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:220px;left:20px;height:280px;width:380px;border-style:outset;border-color:#888888;border-width:2;border-radius: 15px 15px 15px 15px;overflow:auto;'>"


    TABLAX = TablaX & OFV.gentabla("", "aaa style='border-radius: 15px 15px 15px 15px;' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
    TABLAX = TablaX & "<tr><td  "
    TABLAX = TablaX & "class=ttt colspan=3 align=left  STYLE='border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
    TABLAX = TablaX & "<tr><td  "
    TABLAX = TablaX & " >&nbsp;</td></tr>"

    FOR IAEMP=1 TO TAEMP

     '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

       IF ATREMP(IAEMP,2) = 1 THEN

        IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
           TABLAX = TablaX & "<tr  >"
        ELSE
           TABLAX = TablaX & "<tr>"
        END IF

        TABLAX = TablaX & "<td width=10% "
        IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
           TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
           xbord = " style='border-style:solid;border-width:1;' "
           ximg = sImagen4si & "&nbsp;"
        ELSE
           TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
           xbord = "  "
           ximg = sImagenI3 & "&nbsp;"
        END IF




        TABLAX = TablaX & "<td class=vti width=60% " & xbord & " >"
        sHref = "../coress20/novedmenuEstilo.asp?EMPRE=" & ATREMP(IAEMP,0)
        'sHref = ""
        sLink = ofv.GenLink("","cc  onclick=window.top.ejec('"  & sHref & "');","#","","","","<b>" & replace(ATREMP(IAEMP,4) & ""," ","&nbsp;") & "</b>","si")
        TABLAX = TablaX & sLink
        TABLAX = TablaX & "</td>"

        sImage = "<img src='../COREss20/Empresa/" & ATREMP(IAEMP,5) & "' width='100' >"
      '''  sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
        TABLAX = TablaX & "<td width=40% "
        TABLAX = TablaX & "class=vt  >" & sImage & "</td>"
        TABLAX = TablaX & "</tr>"
       END IF

    NEXT

    TABLAX = TablaX & OFV.genFtabla()

    TABLAc = Tablac & TABLAX

    TABLAc = Tablac & "</DIV>"

    xsHTML = xsHTML & TABLAc

  end if
sf="<form id=xfx name=xfx action='' method=post style='display:none;visibility:hidden;'></form>"

selecEmpre=xsHTML & sf
end if

end function

Response.Write sHTML         '& xshtml

%>
