
<%

 FUNCTION GENFMTTABLAS(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""




FNSHTML = FNSHTML &  ofv.GenTabla("TBLMAIN","aa style=""visibility:'visible';"" ","","100%","","0","","0","0","","","","no") & chr(13)




    strsql = "Select * From Tabperfil Where Tabperfil = '" & nPerfil & "'"
 '   response.write strsql
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
      ncolcant = rs2(7)
      if cint(ncolcant) > 10 then ncolcant = 10


nBorde = rs2(2)
nblkfc = rs2(3)
nblkbc = rs2(4)
nblkff = replace(rs2(5)," ","-")
'nblkfs = rs2(6)
nblkfs = cnumer(rs2(6)) & "pt"

nctitfc = rs2(8)
nctitbc = rs2(9)
nctitff = replace(rs2(10)," ","-")
'nctitfs = rs2(11)
nctitfs = cnumer(rs2(11)) & "pt"


ncconfc = rs2(12)
ncconbc = rs2(13)
ncconff = replace(rs2(14)," ","-")
'ncconfs = rs2(15)
ncconfs = cnumer(rs2(15)) & "pt"




  redim preserve aColumnas(cint(ncolcant) - 1)
  redim preserve aColumnasa(cint(ncolcant) - 1)

  redim preserve aColumnascol(cint(ncolcant) - 1)
  redim preserve aColumnasrow(cint(ncolcant) - 1)

      for ix = 0 to UBound(aColumnas)
        aColumnas(ix) = rs2(ix + 16)
        aColumnasa(ix) = rs2(ix + 26)
          aColumnascol (ix) = 0
          aColumnasrow (ix) = 0
      next



    end if

    call ofv.cerrarconsulta(rs2)


 ib=1
 redim preserve aBloques(0)
 redim preserve aBloquess(0)
 redim preserve aBloquest(0)
 xhtml = ""

  Strsql = "SELECT * From Objtabblk Where Tabdef = " & nTabla & " order by tabblksec"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    nID      = rs(0)
    nSec     = rs(4)
    nBTitulo = rs(5)

  wcolspan  = "N"
  wrowspan  = "N"
  wcolspanc = 0
  wrowspanc = 0
  wcolspann = 0
  wcolspanw = "N"


    if len(nBTitulo) > 0 then
       ltitu = len(nBTitulo)
       if  nBTitulo <> space(ltitu) and Ucase(nBTitulo) <> "NO" then

           redim preserve aBloques(ib)
           redim preserve aBloquess(ib)
           redim preserve aBloquest(ib)

           aBloques(ib)  = rs(0)
           aBloquess(ib) = rs(4)
           aBloquest(ib) = rs(5)

           sPoint  = "'BLK" & aBloques(ib) & "'"
           sPointt  = "'TBL" & aBloques(ib) & "'"

           ib = ib + 1

           nStilo  = " Style=" & chr(34) &"font-family:'" & nblkff & "';font-size:" & nblkfs & ";color:" & nblkfc & ";background-color:" & nblkbc & ";text-decoration:none;text-indent:6px;border-style:solid;border-width:1;border-color:#000000; " & chr(34)

           nfonty  = " <font  face=" & chr(34) & replace(nblkff,"-"," ") & chr(34) & " color=" & nblkfc & " >"

           npoint  = nfonty & nSec & "&nbsp;&nbsp;" & replace(nBTitulo," ","&nbsp;") & "&nbsp;&nbsp;" & nfontx
           Celda   = ofv.GenCelda("","ut " & nStilo & " ","","","","","",npoint,"si")
           Fila    = ofv.GenRow("","","","","","","",Celda,"si")

           xHTML   = xHTML & "<tr><td name=" & sPoint & " id=" & sPoint & " >"

           xHTML   = xHTML & ofv.GenTabla(sPointt,"aaa style='margin-top:14px;margin-bottom:6px;' ","","","","0","0","-1","1","","",fila,"si") & chr(13) & chr(13)

           xHTML   = xHTML & "</td></tr>"

       end if
    end if

    cstil = ""

    if nBorde then

       cstil = "border-right-style:solid;border-right-width:1;border-right-color:#cccccc;border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#cccccc;"

    end if




    fila = ""
    strsql = "SELECT * From Objtabcon Where tabblk = " & nID & " order by tabcontipo,tabconsec"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    do until rs2.eof

       nStilot = " Style= " &chr(34)&"font-family:'"  & nctitff  & "';font-size:" & nctitfs & ";color:" & nctitfc &  ";background-color:" & nctitbc & ";"
       nfontt  = " <font face=" & chr(34) & replace(nctitff,"-"," ") & chr(34) & " color=" & nctitfc & " >"


       nStiloc = " Style= "&chr(34) &"font-family:'"  & ncconff  & "';font-size:" & ncconfs & ";color:" & ncconfc &  ";background-color:" & ncconbc & ";"
       nfontc  = " <font face=" & chr(34) & replace(ncconff,"-"," ") & chr(34) & " color=" & ncconfc & " >"


      for ix = 0 to UBound(aColumnas)

          aColumnascol (ix) = 0
          aColumnasrow (ix) = 0

      next


''''''''''''''''''''''''''
    strsql = "SELECT * From Objtabcmb Where TABobj = '" & rs2(26) & "'  "
    strsql = strsql & " order by tabcol "
'   response.write strsql & "<br>" & nid & "<br>"
    set rs21 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    do until rs21.eof

       if  rs21(8) <> "N" then aColumnasrow ((clng(rs21(6)) - 1)) = rs21(8)

       if  rs21(8) <> "N" then
           if  wcolspanw  = "N" then
               if  wrowspan  = "N" then
                   wrowspanc = rs21(8)
                   wcolspanc = rs21(7)
                   wcolspann = clng(rs21(6))
                   wcolspanw = "R"

                   if wcolspanc = "N" then wcolspanc = 1
                   if wrowspanc = "N" then wrowspanc = 1


               end if
           end if
       end if

       if  rs21(7) <> "N" then aColumnascol ((clng(rs21(6)) - 1)) = rs21(7)
       if  rs21(7) <> "N" then
           if  wcolspanw  = "N" then
               if  wcolspan  = "N" then
                   wcolspanc = rs21(7)
                   wcolspann = clng(rs21(6))
                   wcolspanw = "S"

                   if wcolspanc = "N" then wcolspanc = 1

               end if
           end if

       end if





       rs21.movenext
    loop
''''''''''''''''''''''''''




      'Tipo
      nStilo = nStiloc
      nFontr = ncconbc
      nfonta = nfontc
      if rs2(5) = "1" then
         nStilo = nStilot
         nFontr = nctitbc
         nfonta = nfontt
      end if



      nStilo = nStilo & cstil &  chr(34)

      wcsc = 0



'response.write nStilo

      'Columnas
      celda = ""

      ' estado observacion

      Celdaobs  = ""


      if Nestado=5  then




               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs2(0)  & "&ID2=" & nID  & "&ID3=" & nTabla

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       'Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
               'Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
               Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")




      end if


      Celda = Celda & Celdaobs


      for ix = 0 to UBound(aColumnas)


''''''''''''''''''

        wxrs = ""
        wxcs = ""

        if  wcolspanw = "R" then
            if  wrowspan  = "N" and ix = (wcolspann-1) then
                wrowspan  = "S"
                wcolspan  = "S"

                wxrs = wrowspanc
                wxcs = wcolspanc

                if wxrs = 0 then wxrs = 1
                if wxcs = 0 then wxcs = 1

                wxrs = " ROWSPAN=" & wxrs & " "
                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101



            elseif  wrowspan  = "S" and ix = (wcolspann-1) then

                wxrs = ""

                if wcolspanc > 1 then
                   wcsc = wcolspanc
                else
                   wcsc = 1
                end if

            end if




        end if

        if  wcolspanw = "S"  then
            if  wcolspan  = "N" and ix = (wcolspann-1) then
                wcolspan  = "S"


                wxrs = ""
                wxcs = wcolspanc

                if wxcs = 0 then wxcs = 1

                wxcs = " COLSPAN=" & wxcs & " "

                wcsc = 101



            end if
        end if

        if  wcsc = 0 or wcsc > 100 then


''''''+++++++++++++++++++++++

       '     sInput = rs2(ix + 6)
       '     Celda = Celda & ofv.GenCelda("","vt "  & cstil & wxrs & " " & wxcs ,"","","","","",sinput ,"si")


           sSize = cstr(aColumnas(ix)) & "% "
           sConten = rs2(ix + 6)
           if isnull(sconten) then sconten = " "
           lconten = len(sconten)
           if lconten > 0 then
              if sConten = "" or sConten = space(lconten) then sConten = "&nbsp;"
           end if
           if lconten = 0 then sConten = "<font color=" & nFontr & ">-</font>"

           sConten = nfonta &  replace(sConten,chr(13),"<br>") & nfontx
           select case cstr(aColumnasa(ix))
                  case "0"
                         xalig = " align=center "
                  case "1"
                         xalig = " align=left "
                  case "2"
                         xalig = " align=right "
           end select

           Celda = Celda & ofv.GenCelda("","aaa" & nStilo & xalig & " valign=top " & wxrs & " " & wxcs,"",sSize,"","","",sConten,"si")




'''''''++++++++++++++++++++++++++


            if  wcsc > 100 then
                if wcolspanc > 1 then
                   wcsc = wcolspanc - 1
                else
                   wcsc = 0
                end if
            end if

        else

            wcsc = wcsc - 1

            if  wcsc = 0 then
                wcolspan  = "N"
            end if

            if  wcolspanw = "S" and wcolspan  = "N" then

                wcolspanw = "N"

                for ix2 = ix + 1 to UBound(aColumnas)

                    if aColumnascol (ix2) <> 0 then
                       wcolspanc = aColumnascol (ix2)
                       wcolspann = ix2 + 1
                       wcolspanw = "S"
                       exit for
                    end if

                next

            end if

        end if





'''''''''''''''''''

      next

      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

      if  wrowspan  <> "N" then

          wrowspanc = wrowspanc - 1

          if  wrowspanc = 0 then
              wrowspan  = "N"
              wcolspann = 0
              wcolspanw = "N"
              wcolspanc = 0
          end if


      end if


      rs2.movenext
    loop

    xHTML   = xHTML & "<tr><td>"

    if  nBorde then
        xHTML = xHTML & ofv.GenTabla("","","","100%","","0","0","-2","2","#ffffff","",fila,"si") & chr(13)
    else
        xHTML = xHTML & ofv.GenTabla("","","","100%","","0","0","-2","1","#ffffff","",fila,"si") & chr(13)
    end if

    xHTML   = xHTML & "</td></tr>"

    call ofv.cerrarconsulta(rs2)

    rs.movenext
    loop

  call ofv.cerrarconsulta(rs)


  if  ubound(aBloques) > 1 then

      nStilo = " Style=" & chr (34) &"font-family:'"   & nblkff  & "';font-size:" & nblkfs & ";color:" & nblkfc & ";background-color:" & nblkbc & ";text-decoration:none; " & chr(34)
      nfonto = " <font face=" & chr(34) & replace(nblkff,"-"," ") & chr(34) & " color=" & nblkfc & " >"

      Fila   = ""

    for ib=1 to ubound(aBloques)

'''''''''' arma tabla de indices '''''''''''''''''''''''''
        sPoint  = "BLK" & aBloques(ib)

        stexto1 = aBloquess(ib) & "&nbsp;&nbsp;" & aBloquest(ib)
        Celdat  = stexto1

        stexto1 = "<A class=ut href='javascript:void(0);' "
        stexto1 = stexto1 & "   >" & Celdat & "</a>"

        Celdat  = "<td class=utis style='cursor:pointer;' onclick=""document.all['TBLMAIN'].style.visibility='hidden';creawin('" & sPoint & "');document.all['TBLMAIN'].style.visibility='visible';"" >" & sfidkB & "</td>"
        Celdat  = Celdat & "<td class=uti onclick=""document.all['TBLMAIN'].style.visibility='hidden';creawin('" & sPoint & "');document.all['TBLMAIN'].style.visibility='visible';"" >" & stexto1 & "</td>"
        Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    next


  end if


  FNSHTML = FNSHTML & xHTML

  FNSHTML = FNSHTML & "</table>"

  GENFMTTABLAS = "<div id=zglos class='norma-card' style='border-radius: 0px 0px 10px 10px;' >" & FNSHTML & "</div>"


 END FUNCTION

function cnumer(cdato)
   ldato = len(cdato)

   for kk=1 to ldato
       if isnumeric(mid(cdato,kk,1)) then
           cnumer = cnumer & mid(cdato,kk,1)
       end if
   next


end function

%>