
<%

 FUNCTION GENFMTCONTROLES(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


'''''''''''''''''''''''''''''''''''''''''''''''''

  Celdat  = "<td class=uts colspan=2 style='border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</td>"
  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

  sfidkD  = sfidkc
  stexto1 = stextoact

  Celdat  = "<td class=uts >" & sfidkD & "</td>"
  Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"
  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")


  'Encabezado Descripcion
  FNSHTML = FNSHTML & ofv.GenTabla("","","","100%","",0,0,0,0,"","","","no")

  if nID <> "no" then

  strsql = "SELECT Objcontrol2.* FROM Objcontrol2 inner join Objcontrol on (Objcontrol2.control = Objcontrol.objid)   where "
'  strsql = strsql & "Objcontrol2.control = Objcontrol.objid "
  strsql = strsql & " Objcontrol.objid = " & nID

  if scambios = "NO" then
     strsql = strsql & " and  (isnull(Objcontrol.modi,0) = 0  or (isnull(Objcontrol.modi,0) <> 0 and Objcontrol.tipmodi <> 'E')) "
     strsql = strsql & " and (isnull(Objcontrol2.modi,0) = 0 or (isnull(Objcontrol2.modi,0) <> 0 and Objcontrol2.tipmodi <> 'E')) "
  else
     strsql = strsql & " and ( isnull(Objcontrol.modi,0) <> 0  "
     strsql = strsql & " or isnull(Objcontrol2.modi,0) <> 0 ) "
  end if

  strsql = strsql & " ORDER BY Objcontrol2.secuencia, Objcontrol2.descri, Objcontrol2.OBJID "
' response.write strsql


  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posicion = 0
  do until rs2.eof
     posicion = posicion + 1
     nAccion = rs2(0)
     RCHEK = "N"
     'Texto Descripcion
     Celda      = ofv.GenCelda("","ut colspan=4","","100%","","","","&nbsp;","si")
     FNSHTML      = FNSHTML & ofv.GenRow("","","","","20","","",Celda,"si") & chr(13)

    pppp = rs2(7)
   'response.write pppp & " 001" & bVersion & SESSION("SHWMOD")
    ppps = rs2(6)
   'response.write ppps & " 002"

      xfclas    = "dt"
      sImage    = "&nbsp;"
      if bVersion = "SI" and rs2(6) and SESSION("SHWMOD") <> "N" then



      if rs2(7) = "M" then
         xftitl = "Modificado."
         xfclas = "btbs"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")

        pppp = rs2(7)
      ''response.write pppp & " 00A"
''''''para mostrar el item modificado
''''''Nuevo 03012014

         strsql = "SELECT Objcontrol2.* FROM Objcontrol2 where "
         strsql = strsql & " Objcontrol2.objid =0" & rs2("relant")

   'response.write strsql
   if scambios <> "NO" then
      set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rs3.eof then
       	RCHEK = "S"
''    	Celdakk = Celdakk & ofv.GenCelda("","","","","","","","&nbsp; aaa ","si")
      end if
   end if
   ''     response.write rchek & " 003"

''''''Nuevo 03012014 Fin
  ''''''Fin mostrar item modificado

      elseif rs2(7) = "A" then
         xftitl = "Agregado."
         xfclas = "btba"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >6</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif rs2(7) = "E" then
         xftitl = "Eliminado."
         xfclas = "btbn"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >5</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      end if
      end if


    ''''Nuevo 03012014

   ''  Celda  = ofv.GenCelda("","uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")

       Celda  = ofv.GenCelda("","uts title='" & xftitl & "' onclick='document.getElementById(" & chr(34) & rs2(0) & "M" & chr(34) & ").style.display=""block"";' ","","1%","","","","<b>" & sImage & "</b>","si")

    ''''Nuevo 03012014 Fin

      ' response.write sImage & " 0001 "

      ' estado observacion

      Celdaobs  = ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")


      if Nestado=5  then

      	     '  sAccion = "genObs.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	     '  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID

      	     '  xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      	     '  sHref   = sAccion
      	     '  sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	     '  Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Observacion de " & BtipP & bpespec & "' ","","1%","1","","",sInput,"si")
             '  Celdaobs  = ofv.GenLink("",xclas & " title='Obs de " & BtipP & bpespec & "' ",sHref,"","","",Celdaobs,"si")


              sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nAccion  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	      ' Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")'
               'Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")'

               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
               Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")




      end if


      Celda = Celda & Celdaobs


     sPoint = "T" & nAccion

      sInput1 = ""
      sInput2 = ""


     sInput = "<b>"
     'sInput = sInput  & rs2(5) & ".&nbsp;<font color=#f7aa22 face=corflaticon size=3>&#0232;</font></b>"
	 sInput = sInput  & rs2(5) & ".&nbsp;<i class='las la-angle-right'></i></b>"
     Celda  = Celda & ofv.GenCelda(sPoint,"uts ","right","10% valign=top","","","",sInput,"si")

'''''''''' arma tabla de indices '''''''''''''''''''''''''
      stexto1 = rs2(4)
      if len(stexto1) > 40 then
         stexto1 = left(stexto1,40) & "..."
      end if

      stexto1 = rs2(5) & ".&nbsp;&nbsp;" & stexto1
      Celdat  = stexto1

      stexto1 = "<A class=ut href='javascript:void(0);' "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=utis style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & sfidkB & "</td>"
      Celdat  = Celdat & "<td class=uti onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      sImagen  = "NO"
      wasotxt  = "N"

      sInput1 = RS2(4) & ""





      'link

         wlnkptr = "0" & rs2(8)
         if not isnull(wlnkptr) then
          if clng(wlnkptr) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & wlnkptr
           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          'response.write strsql & " strsql " & "<br>"
           if not rsx.eof then


                wasotxt     = rsx(14) & ""
                wasotxtstr  = CLNG("0" & rsx(15))
                wasotxtcan  = CLNG("0" & rsx(16))
                wasotxtdat  = rsx(17) & ""

                IF wasotxt = "" THEN wasotxt = "N"


              wtipoL   = rsx(6)
              wtipoLc  = rsx(7) & ""
              wUrl     = rsx(8) & ""
              wUrlDesc = rsx(9) & ""
              wtipoLNK = rsx(10) & ""

              wNV1    = rsX(11) & ""
              wNV2    = rsX(12) & ""
              wNV3    = rsX(13) & ""

              IF  WNV1 = "" THEN WNV1 = "0"
              IF  WNV2 = "" THEN WNV2 = "0"
              IF  WNV3 = "" THEN WNV3 = "0"

           else
              wtipoL   = 0
              wtipoLc  = "NO"
              wUrl     = "NO"
              wUrlDesc = ""
              wtipoLNK = "NO"

              WNV1 = "0"
              WNV2 = "0"
              WNV3 = "0"

           end if




           call ofv.cerrarconsulta(rsx)

           if wtipoLNK = "I" then

              wNomLt    = wUrlDesc
              wNomL     = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomLt)
              else
                 wNomL = "S/D"
              end if


              if not isnull(wUrl) then
                 sHref = instr(wUrl,"://")
                 if sHref > 0 then
                    sHref = wUrl
                 else
                    sHref = "HTTP://" & wUrl
                 end if
              else
                 sHref = "javascript:void(0);"
              end if
              sFuncion = " target=_blank "


           else

              wNomLt    = wUrlDesc
              wNomL     = wUrlDesc
              if wNomL <> "" then
                 wNomL = UCASE(wNomLt)
              else
                 wNomL = "S/D"
              end if



                 Strsql = "Select * From csGenObjetos where codobjeto = " & wtipoL
                 set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                'response.write strsql & " strsql " & "<br>"
                 if not rsx.eof then
                    wNomL = rsx(1)
                 else
                    wNomL = "S/D"
                 end if
                 call ofv.cerrarconsulta(rsx)


         '     sActcon = "Consultar.asp?Codigo=" & wtipoL & "&Tipo=" & wtipoLc
         '     sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
              sHref = "&nbsp;"

              sActcon = "Consultar.asp?Codigo=" & wtipoL & "&Tipo=" & wtipoLc
              sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
              sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3


              sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoL & "'); return false;" & Chr(34)

           end if

           slnktit = wNomLt

            SArma = wNomL

            Slnkdata = wNomL

            sImagen  = "SI"

            sLink       = ofv.GenLink("","btbn title='" & slnktit & "' ",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),sgllnk,"si")
            sLink2       = ofv.GenLink("","nt title='" & slnktit & "' ",GenKeyLnks(sHref),"","",GenKeyLnks(sFuncion),"NNNNNN","si")


          IF wasotxt =  "N" then

            sLink2 = replace(sLink2,"NNNNNN",slnkdata)


            Celdak = ""
            Celdak = Celdak & ofv.GenCelda("","btbn","","2","","","",sLink,"si")
            Celdak = Celdak & ofv.GenCelda("","uti","","","","","",sLink2,"si")

            Filak = ofv.GenRow("","","","","","","",Celdak,"si")
            SArma = ofv.GenTabla("","TWINE","","300","","0","0","0","-1","","",Filak,"si")

            sinput2 = SArma

         end if


        end if
       end if


      if len(sInput1) > 0 then

          sArma = sInput1

        IF wasotxt = "S" THEN
           if  wasotxtstr > 0  and wasotxtcan > 0 then

               sfontstyl = " style='font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("APLFT") & ";font-size:" & Session("CMFPNGNFTSZ") + 2 & "pt;font-weight:800;' "

               sArma = left(sArma,wasotxtstr - 1) & "" & replace(sLink2,"NNNNNN","&nbsp;<u><font " & sfontstyl & " >" & mid(sArma,wasotxtstr,wasotxtcan) & "</font></u>&nbsp;") & "" & right(sArma,len(sArma) - (wasotxtstr + wasotxtcan - 1))


             '''  sArma = left(sArma,wasotxtstr - 1) & "" & replace(sLink2,"NNNNNN","&nbsp;<u><b><font size=2>" & mid(sArma,wasotxtstr,wasotxtcan) & "</font></b></u>&nbsp;") & "" & right(sArma,len(sArma) - (wasotxtstr + wasotxtcan - 1))
           end if

        end if

        sArma = replace(sArma,chr(13),"<br>")

        IF wasotxt = "N" THEN

           sArma = sArma & "<BR><BR>" & sinput2


        end if

        sinput = sArma

     else

         sinput = sinput2

     end if


     Celda  = Celda & ofv.GenCelda("","UTG style='text-align:justify;'  ","","99%","","","",sInput,"si")

     FNSHTML  = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si") & chr(13)

'' Nuevo 03012014
  	  if  RCHEK = "S"   then
          Celdakk = ""
          sInput  = ""
  ''      Celdakk = Celdakk & ofv.GenCelda("","utg colspan=3 style='text-align:justify;color:#00aa00;' ","","","","","","-Texto Anterior-","si")
          Celdakk = Celdakk & ofv.GenCelda("","utg colspan=6 style='text-align:justify;color:#0054a9;' ","","25%","","","","<u>Version Anterior:</u>","si")
          Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")
          sInput  = replace(rs3(4)," ","&nbsp;")
          Celdakk = ""
          Celdakk = Celdakk & ofv.GenCelda("","utg valign=top style='text-align:justify;color:#0054a9;' ","","15%","","","",sInput,"si")
          Filatt = Filatt & ofv.GenRow("","","","","","","",Celdakk,"si")
          SArmatt = ofv.GenTabla("","AAA style='BACKGROUND-COLOR:#f9f9f9;border-style:solid;border-width:1;border-color:#0054a9;' ","","80%","","0","0","2","2","","",Filatt,"si")
          Celdakk = ""
          Celdakk = Celdakk & ofv.GenCelda("","ut colspan=7 align=CENTER ","","","","","",SArmatt,"si")
          Fila = ""
          Fila = Fila & ofv.GenRow("","","","","","","",Celdakk,"si")
          Celdakk = ""
          Celdakk = Celdakk & ofv.GenCelda("","ut colspan=7 align=CENTER ","","","","","","&nbsp;&nbsp;","si")
          Fila = Fila & ofv.GenRow("","","","","","","",Celdakk,"si")
 ''       FNSHTML = FNSHTML & ofv.GenRow(rs2(0) & "M","aa style='xdisplay:none;border-style:solid;border-width:1;border-color:#0054a9;' ","","","","","",Celdakk,"si")
          Celda  = ofv.GenCelda("","UTG style='text-align:justify;' colspan=4 ","","99%","","","",ofv.GenTabla("","","","95%","","0","0","0","0","","",Fila,"si"),"si")
          FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si") & chr(13)
          Fila = ""
      end if
     ''''Nuevo 03012014 Fin

      rs2.movenext
  loop

 else
     h1 = ofv.GenCelda("","xx colspan=4","","88%","","","","&nbsp;","si")
     FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
     h1 = ofv.GenCelda("","dt","","20","","","","&nbsp;","si")
     sInput = "<p align=left><b>"
     sInput = sInput  & "&nbsp;Sin&nbsp;datos&nbsp;disponibles&nbsp;</b></p>"
     h1 = h1 & ofv.GenCelda("","","","5%","","","","&nbsp;","si")
     h1 = h1 & ofv.GenCelda("","et","","95%","","","",sInput,"si")
     FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")

  end if


    h1 = ofv.GenCelda("","dt","","1%","","","","&nbsp;","si")
    h1 = h1 & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
    h1 = h1 & ofv.GenCelda("","","","5%","","","","&nbsp;","si")
    h1 = h1 & ofv.GenCelda("","","","95%","","","","&nbsp;","si")
    FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
    FNSHTML = FNSHTML & ofv.GenFTabla()

  GENFMTCONTROLES = "<div id=zglos class='norma-card' style='border-radius: 0px 0px 10px 10px;border-top:none;max-width:100%;box-shadow:none;'>" & FNSHTML    & "</div>"
  'GENFMTCONTROLES = "<div id=zglos >" & FNSHTML    & "</div>"


 END FUNCTION

%>