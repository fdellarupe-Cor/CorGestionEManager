
<%


 FUNCTION GENFMTESQUEMAS(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""


  ''Celdat  = "<td class=uts colspan=2 style='border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</td>"
  ''Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

  sfidkD  = sfidkc
  stexto1 = stextoact

 '' Celdat  = "<td class=uts >" & sfidkD & "</td>"
 '' Celdat  = Celdat & "<td class=ut  >hhhhhhhhhh" & stexto1 & "</td>"
 '' Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")


''' response.write vuelta & "<br>" & chr(13)
''' response.write " nID " & nID & "</div>"

    if nID <> "no" then

  Strsql = "SELECT OBJESQUEMA2.*, OBJCUENTA.descri, OBJCUENTA.numero, OBJCUENTA.objid "
  Strsql = Strsql & "FROM OBJESQUEMA2, OBJESQUEMA, OBJCUENTA where "
  Strsql = Strsql & "OBJESQUEMA2.codesquema = OBJESQUEMA.objid and "
  Strsql = Strsql & "OBJESQUEMA2.codcuenta = OBJCUENTA.objid "
  Strsql = Strsql & "and Objesquema2.codesquema=0" & nID

  if scambios = "NO" then
     strsql = strsql & " and (OBJESQUEMA.modi = 0 or OBJESQUEMA.modi is null or (OBJESQUEMA.modi <> 0 and OBJESQUEMA.tipmodi <> 'E')) "
     strsql = strsql & " and (OBJESQUEMA2.modi = 0 or (OBJESQUEMA2.modi <> 0 and OBJESQUEMA2.tipmodi <> 'E')) "
  else
     strsql = strsql & " and (OBJESQUEMA.modi <> 0  "
     strsql = strsql & " or OBJESQUEMA2.modi <> 0)  "
  end if

''' response.write nid & " 001"
 '   response.write strsql 

  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  nEsquema = ID
  FNSHTML = FNSHTML & ofv.GenTabla("","aaa ","","100%","",0,0,0,0,"","","","no")
  'Encabezados Cuenta Contable y Sentido

      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","tt colspan=3","","1%","","","","Sentido","si")
      Celda  = Celda & ofv.GenCelda("","tt colspan=2","","50%","","","","Cuenta&nbsp;Contable","si")
      Celda  = Celda & ofv.GenCelda("","tt","","50%","","","","Referencia","si")
 ''   FNSHTML  = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si")

  posicion = 0
  do until rs2.eof
     posicion = posicion + 1
     nCuenta  = rs2(0)
     nref     = rs2(6)
     sPoint   = "T" & rs2(0)
     RCHEK    = "N"

      xfclas    = "dt"
      sImage    = "&nbsp;"
      if bVersion = "SI" and rs2(7) and SESSION("SHWMOD") <> "N" then
      if rs2(8) = "M" then
         xftitl = "Modificado."
         xfclas = "btbs"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")

  ''''''para mostrar el item modificado
         srelant = rs2(11)
  ''     response.write " srelantaaaa = " & srelant & "</div>"

         Strsql = "SELECT OBJESQUEMA2.*, OBJCUENTA.descri, OBJCUENTA.numero, OBJCUENTA.objid "
         Strsql = Strsql & "FROM OBJESQUEMA2, OBJESQUEMA, OBJCUENTA where "
         Strsql = Strsql & "OBJESQUEMA2.codesquema = OBJESQUEMA.objid and "
         Strsql = Strsql & "OBJESQUEMA2.codcuenta = OBJCUENTA.objid and "
         Strsql = Strsql & "OBJESQUEMA2.objid=0" & srelant

  ''     response.write strsql & "</div>"

  if scambios <> "NO" then
     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs3.eof then
            RCHEK = "S"
''          Celdakk = Celdakk & ofv.GenCelda("","","","","","","","&nbsp; aaa ","si")
     end if
  end if

  '''  response.write strsql

  ''''''Fin mostrar item modificado

 ''     response.write strsql

      elseif rs2(8) = "A" then
         xftitl = "Agregado."
         xfclas = "btba"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >6</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif rs2(8) = "E" then
         xftitl = "Eliminado."
         xfclas = "btbn"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >5</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      end if
      end if

''''' Celda  = ofv.GenCelda("","uts title='" & xftitl & "' onclick='document.getElementById(" & chr(34) & rs2(0) & "M" & chr(34) & ").style.display=""block"";' ","","1%","","","","<b>" & sImage & "</b>","si")
      Celda  = ofv.GenCelda(sPoint,"uts title='" & xftitl & "' onclick='document.getElementById(" & chr(34) & rs2(0) & "M" & chr(34) & ").style.display=""block"";' ","","1%","","","","<b>" & sImage & "</b>","si")


'' response.write strsql

''  response.write sImage & " 0001 "
''  response.write Celda  & " 0002 "



      ' estado observacion

      Celdaobs  = ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")




      if Nestado=5  then


              sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nCuenta  & "&ID2=" & nID

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
                Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")


      end if


      Celda = Celda & Celdaobs



  'Texto Sentido
      sInput = rs2(5)
      if  sInput = "Debe" THEN
          sInput = "DB.&nbsp;"
      ELSE
          sInput = "&nbsp;&nbsp;&nbsp;HB.&nbsp;"
      END IF
      Celda  = Celda & ofv.GenCelda("","Ut valign=top ","","1%","","","",sInput,"si")
''    response.write Celda  & " Texto "

  'Texto Cuentas Contables
      sInput = replace(rs2(16)," ","&nbsp;")
      Celda  = Celda & ofv.GenCelda(sPoint,"ut valign=top ","","15%","","","","<b>" & sInput & "</b>","si")


      sInput = replace(rs2(15)," ","&nbsp;")
      Celda  = Celda & ofv.GenCelda("","Ut valign=top  ","","35%","","","",sInput,"si")
''    response.write Celda  & " Cuenta "


'''''''''''' arma tabla de indices '''''''''''''''''''''''''
       stexto1 = replace(rs2(16)," ","&nbsp;") & "&nbsp;&nbsp;" & replace(rs2(15)," ","&nbsp;")
       Celdat  = stexto1

      stexto1 = "<A class=ut href='javascript:void(0);' "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=utis style='cursor:pointer;' onclick=""creawin('" & sPoint & "');"" >" & sfidkB & "</td>"
      Celdat  = Celdat & "<td class=uti onclick=""creawin('" & sPoint & "');"" >" & stexto1 & "</td>"
      Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


  'Texto Referencia

      sInput1 = ""
      sInput2 = ""



'''''''''''''''''''''''''''''''''''''''''''
      sImagen  = "NO"
      wasotxt  = "N"

      sInput1 = nref & ""





      'link

         wlnkptr = "0" & rs2(9)
         if not isnull(wlnkptr) then
          if clng(wlnkptr) > 0 then

           Strsql = "SELECT * From objlnk where objid = " & wlnkptr
           set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
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
           else
              wtipoL   = 0
              wtipoLc  = "NO"
              wUrl     = "NO"
              wUrlDesc = ""
              wtipoLNK = "NO"
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
                 if not rsx.eof then
                    wNomL = rsx(1)
                 else
                    wNomL = "S/D"
                 end if
                 call ofv.cerrarconsulta(rsx)


              sActcon = "Consultar.asp?Codigo=" & wtipoL & "&Tipo=" & wtipoLc
              sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
              sHref = "&nbsp;"

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


 ''  response.write strsql

 '' Nuevo 03012014
   if scambios <> "NO" then
''      if not rs3.eof then
         if RCHEK = "S" then
''       	  Celdakk = Celdakk & ofv.GenCelda("","","","","","","",sLink,"si")
         end if
''      end if
   end if
 '' Nuevo 03012014 fin

 '' response.write strsql



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


            '''   sArma = left(sArma,wasotxtstr - 1) & "" & replace(sLink2,"NNNNNN","&nbsp;<u><b><font size=2>" & mid(sArma,wasotxtstr,wasotxtcan) & "</font></b></u>&nbsp;") & "" & right(sArma,len(sArma) - (wasotxtstr + wasotxtcan - 1))
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


''''''''''''''''''''''''''''''''''''''''''''



      Celda  = Celda & ofv.GenCelda("","UTG style='text-align:justify;' valign=top ","","50%","","","",sInput,"si")
      FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si")

'' response.write rchek & "001"

'' Nuevo 03012014

         	  if  RCHEK = "S" then
      '' 	  	  sArmakk = ""
                Celdakk = ""
                sInput = rs2(5)
                if  sInput = "DEBE" THEN
                    sInput = "DB.&nbsp;"
                ELSE
                    sInput = "&nbsp;&nbsp;&nbsp;HB.&nbsp;"
                END IF
''              Celdakk = Celdakk & ofv.GenCelda("","aaa colspan=2 style='text-align:justify;color:#00aa00;' ","","","","","","-Version Anterior-","si")
                Celdakk = Celdakk & ofv.GenCelda("","utg colspan=6 style='text-align:justify;color:#0054a9;' ","","25%","","","","<u>Version Anterior:</u>","si")
                Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")
                Celdakk  = ""
                sInput = sInput & replace(rs3(16)," ","&nbsp;")
                Celdakk = ""
                Celdakk  = Celdakk & ofv.GenCelda("","ut valign=top style='text-align:justify;color:#0054a9;' ","","15%","","","","<b>" & sInput & "</b>","si")
                sInput = replace(rs3(15)," ","&nbsp;")
''              Celdakk = ""
                Celdakk  = Celdakk & ofv.GenCelda("","Ut valign=top  style='text-align:justify;color:#0054a9;' ","","35%","","","",sInput,"si")
                Filatt = Filatt & ofv.GenRow("","","","","","","",Celdakk,"si")
                SArmatt = ofv.GenTabla("","AAA style='BACKGROUND-COLOR:#f9f9f9;border-style:solid;border-width:1;border-color:#0054a9;' ","","80%","","0","0","2","2","","",Filatt,"si")
                Celdakk = ""
                Celdakk = Celdakk & ofv.GenCelda("","ut colspan=7 align=CENTER ","","","","","",SArmatt,"si")
                Fila = ""
                Fila = Fila & ofv.GenRow("","","","","","","",Celdakk,"si")
                Celdakk = ""
                Celdakk = Celdakk & ofv.GenCelda("","ut colspan=7 align=CENTER ","","","","","","&nbsp;&nbsp;","si")
                Fila = Fila & ofv.GenRow("","","","","","","",Celdakk,"si")
''              FNSHTML = FNSHTML & ofv.GenRow(rs2(0) & "M","aa style='display:none;border-style:solid;border-width:1;border-color:#00aa00;' ","","","","","",Celdakk,"si")
''              FNSHTML = FNSHTML & ofv.GenRow(rs2(0) & "M","aa style='xdisplay:none;border-style:solid;border-width:1;border-color:#0054a9;' ","","","","","",Celdakk,"si")
                Celda  = ofv.GenCelda("","UTG style='text-align:justify;' colspan=7 ","","99%","","","",ofv.GenTabla("","","","95%","","0","0","0","0","","",Fila,"si"),"si")
                FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si")
            end if

 '' Nuevo 03012014 FIN

     Celda = ofv.GenCelda("","dt colspan=5 ","","20","","","","&nbsp;","si")
     FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si")


      rs2.movenext

    loop


    FNSHTML = FNSHTML & ofv.GenFTabla()

  call ofv.cerrarconsulta(rs2)
    end if

  GENFMTESQUEMAS = "<div id=zglos>" & FNSHTML  & "</div>"


 END FUNCTION


%>