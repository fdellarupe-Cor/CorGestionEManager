
<%

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 FUNCTION GENFMTESTRUCTURA(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""
''response.write id

  Fila  = ""

  FNSHTML = FNSHTML &  ofv.GenTabla("","","","100%  ","","0","","0","0","","","","no")

  wcantlin = 0

Dim Inx
dim nrs(100)


  strsql = " SELECT * FROM CSESTDET "
  strsql = strsql & " where detcon = " & ID & "  "
   if marca <> "CA" then
     strsql = strsql & " and (modi = 0 or (modi <> 0 and tipmodi <> 'E')) "
  end if
  strsql = strsql & " and not ( "
  strsql = strsql & " NOVA = 'N' OR (BDATO is null AND DETDEPEN > 0 ) "
  strsql = strsql & " ) and stmodi = '' "
  strsql = strsql & " order by TSEC  "

''response.write strsql
''response.write marca & " 01"

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorType = 3
Rs.CacheSize = 100

Rs.Open strsql, cn

' ars = rs.GetRows()
if not rs.eof then
  ars = rs.GetRows()
''response.write id & "ok"

else
 campos = rs.fields.count - 1
 redim  ars(cint(campos),0)
 ars(39,0) = "N"
 wentro=1
end if


rs.Close
set rs = nothing

  sant = "X"
  bseci = 0
  bsecix = 0
  redim preserve bsec(8)
  redim preserve bnum(8)
  redim preserve bnumG(8)
  for i=0 to 8
      bsec(i) = 0
      bnum(i) = 0
      bnumG(i) = 0
  next

 inx = 0



 do while wentro = 0



for Inx = inx to UBound(ars,2)



      SCLAVE   = ars(0,inx)
      Stip     = ars(5,inx)
      sForm    = "edestrdet" & SCLAVE
      smodi    = ars(29,inx)
      stipmodi = ars(30,inx)

      if marca = "CA" then
         Mostrar = "SI"
         if Stip = "G" or  Stip = "H" or  Stip = "I"  or  Stip = "J" then
            if smodiant and stipmodiant = "E" then
               smodi = smodiant
               stipmodi = stipmodiant
            elseif not smodi then
                 Mostrar = "NO"

            end if
         else
            smodiant = ars(29,inx)
            stipmodiant = ars(30,inx)
         end if

      end if

      if Mostrar <> "" then
         if Marca = "SI" then
            if Mostrar = "NO" then
               if trim(SID) = trim(sclave) then
                  MCtipo = Stip
                  Mostrar = "SI"
               end if
            else
               if Stip = MCtipo or Stip < MCtipo then
                  Mostrar = "NO"
               end if
            end if
         end if
      end if


      pmk = "<form name="  & chr(34) & "T" & ars(0,inx)  & chr(34) &  " method=post >"
      pmk = pmk & "<input type=hidden value=a ></form>"

      sImage      = "&nbsp;"
      xfclas      = "dt"

      RCHEK = "N"
      Celdakk = ""

      if bVersion = "SI" and smodi and SESSION("SHWMOD") <> "N" then




      if stipmodi = "M" then
         xfclas = "btbs"
         xftitl = "Modificado."

'''''''' sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>" & ars(0,inx)

         sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
''''''para mostrar el item modificado

         Strsql = "Select * from CSESTDET where objid = 0" & ars(34,inx)
         set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rs3.eof then
           RCHEK = "S"
         end if
  ''response.write strsql

      elseif stipmodi = "A" then
         xfclas = "btba"
         xftitl = "Agregado."
         sglmod = "<div class='" & xfclas & "' ><font class=btic >6</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      elseif stipmodi = "E" then
         xfclas = "btbn"
         xftitl = "Eliminado."
         sglmod = "<div class='" & xfclas & "' ><font class=btic >5</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
      end if
      end if




      supersec = "" & ars(38,inx)

      supersecarr = split(supersec,".")

      spsecn = ubound(supersecarr)


      sPoint = "T" & ars(0,inx)

      ' estado observacion

  '''response.write nCodigo & " nCodigo " & chr(13)
  '''response.write sTipo   & " sTipo "  & chr(13)

      if Nestado=5  and origen <> "ED"  then

         set cnx = ofv.conectar(ofv.strconn5)
         strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
       ''response.write strsql
	 set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
         if not rss.eof then
            idcircuito = rss("segckt")
            rsscondinx = rss(0)
         end if
         call ofv.cerrarconsulta(rss)
         call ofv.cerrarconn(cnx)

     '''response.write idcircuito & " idcircuito "
     '''response.write rsscondinx & " rsscondinx "

         strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codcircuito = '"& idcircuito &"'  and codinx1 = '" & rsscondinx & "' and "
         strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and "
         strsql = strsql & " codusr = '" & session("usuario") & "' "
         strsql = strsql & " order by objid"
        'response.write strsql
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

         sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
         sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & sclave  & "&ID2=" & ID

         xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
         sHref   = sAccion
         if rs.eof then
            sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"

         else
            sInput  = "<div style='border-style:solid;border-color:#DC143C;border-width:1;' >" & "OBS" & "</div>"

         end if
                    ''Celdaobs  = ofv.Gencelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
         ''Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")
              Celdaobs= ofv.GenLink("","ut" & "  ",sHref,"WK4","","",sInput,"si")
         Celdaobs = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",Celdaobs,"si")

      end if


      Celda  = Celdaobs & ofv.GenCelda("","uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")

        if Marca  = "CA" then
        	 if   RCHEK = "S" then
''        	 	Celdakk = Celdakk & ofv.GenCelda("","","","","","","","&nbsp;","si")
           end if
        end if

      'Secuencia
      select case stip
             case "A"
                  t = 7
                  bseci = 0
                  bsecix = 0
                  bsec(1) = ars(4,inx)
                  bnum(0) = bnum(0) + 1
                  bnumG(0) = bnumG(0) + 1
                  for i=bseci + 1 to 8
                      bnum(i) = 0
                      bnumG(i) = 0
                  next
             case "B"
                  t = 8
                  bseci = 1
                  bsecix = 1
                  bsec(2) = ars(4,inx)
                  bnum(1) = bnum(1) + 1
                  bnumG(1) = bnumG(1) + 1
                  for i=bseci + 1 to 8
                      bnum(i) = 0
                      bnumG(i) = 0
                  next
             case "C"
                  t = 9
                  bseci = 2
                  bsecix = 2
                  bsec(3) = ars(4,inx)
                  bnum(2) = bnum(2) + 1
                  bnumG(2) = bnumG(2) + 1
                  for i=bseci + 1 to 8
                      bnum(i) = 0
                      bnumG(i) = 0
                  next
             case "D"
                  t = 10
                  bseci = 3
                  bsecix = 3
                  bsec(4) = ars(4,inx)
                  bnum(3) = bnum(3) + 1
                  bnumG(3) = bnumG(3) + 1
                  for i=bseci + 1 to 8
                      bnum(i) = 0
                      bnumG(i) = 0
                  next
             case "E"
                  t = 11
                  bseci = 4
                  bsecix = 4
                  bsec(5) = ars(4,inx)
                  bnum(4) = bnum(4) + 1
                  bnumG(4) = bnumG(4) + 1
                  for i=bseci + 1 to 8
                      bnum(i) = 0
                      bnumG(i) = 0
                  next
             case "F"
                  t = 12
                  bseci = 5
                  bsecix = 5
                  bsec(6) = ars(4,inx)
                  bnum(5) = bnum(5) + 1
                  bnumG(5) = bnumG(5) + 1
                  for i=bseci + 1 to 8
                      bnum(i) = 0
                      bnumG(i) = 0
                  next
             case "G"
                  t = 13




                  select case cstr(spsecn & "")
                         case "1"
                              ta = 7
                              bseci = 0
                         case "2"
                               'ta = 8
                              ta = 7
                              bseci = 1
                         case "3"
                              ta = 9
                              bseci = 2
                         case "4"
                              ta = 10
                              bseci = 3
                         case "5"
                              ta = 11
                              bseci = 4
                         case "6"
                              ta = 12
                              bseci = 5
                  end select


                  bnum(6) = bnum(6) + 1
                  bnumG(6) = bnumG(6) + 1

                  bsecix = bseci
                  bsec(7) = ars(4,inx)

             case "H"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bnum(7) = bnum(7) + 1
                  bnumG(7) = bnumG(7) + 1
                  bsecix = bseci + 1
                  bsec(7) = ars(4,inx)
             case "I"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bnum(8) = bnum(8) + 1
                  bnumG(8) = bnumG(8) + 1
                  bsecix = bseci + 1
                  bsec(7) = ars(4,inx)
             case "J"
                  select case sant
                         case "A"
                              bseci = 0
                         case "B"
                              bseci = 1
                         case "C"
                              bseci = 2
                         case "D"
                              bseci = 3
                         case "E"
                              bseci = 4
                         case "F"
                              bseci = 5
                  end select
                  bnum(8) = bnum(8) + 1
                  bnumG(8) = bnumG(8) + 1
                  bsecix = bseci + 1
                  bsec(8) = ars(4,inx)

      end select

  if stip <> "H" and stip <> "I"  and stip <> "J" then

     for ti = 0 to 17
         Pdef(t,ti) = Pdef(t-7,ti)
     next
     Pdef(t,18) = "SI"


     if ars(31,inx) = "NO" then
        for ti = 11 to 28
            if (ars(ti,inx)&"") = "" then
               sti = ""
            else
               sti             = ars(ti,inx)
               Pdef(t,ti - 11) = sti
            end if



        next
        Pdef(t,18) = "NO"

     end if
  end if

      BsecD = ""
      for i=1 to bsecix
          BsecD = BsecD & bsec(i) & "."
      next





      'response.write strsql
      if Pdef(t,18) = "NO" then
         tx = t
      else
         tx = t - 7
      end if


         PNumCls = Pdef(tx,14)
         if PNumCls = "" then PNumCls = "Simple"



         if stip <> "H" and stip <> "I"  and stip <> "J" then
            PNumTip(tx) = Pdef(tx,13)
            if  PNumTip(tx) = "" then PNumTip(tx) = "1"
         end if

         if stip = "G" then
            if Pdef(tx,9) = "NO" then
               MPvin = "NO"
            else
                MPvin = "SI"
            end if
         end if

' Anula la numeracion para este tipo o elemento
        if  PNumTip(tx) = "NO" then PNumCls = PNumTip(tx)

         PFont  = Pdef(tx,0)
         PFsize = Pdef(tx,1)
         PFcol  = Pdef(tx,2)
         Pbgcol = Pdef(tx,3)
         PBorde = Pdef(tx,4)
         PFtipo = Pdef(tx,5)

         if marca = "CA" then
            PFont  = "VERDANA"
            if smodi then
               PFcol  = "#000080"
            else
''               PFcol  = "#aaaaaa"
                 PFcol  = "#666666"
            end if
            PFsize = "8"
            Pbgcol = ""
            PBorde = ""
            PFtipo = ""
         end if


         PSang  = Pdef(tx,6)
         txi = tx
         if tx > 6 then txi = tx - 7
         PEspa  = Pdef(txi,8)



         Palign = Pdef(tx,16)
         if Palign = "" then Palign = "left"
         Ptalig = Pdef(tx,17)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' arma span inicial ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      if Marca <> "SI" and Marca <> "CA" then
        Celda = Celda & Ofv.Gencelda("","cc","","1%","","","",BsecD & ars(4,inx),"si")
      else
        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
        if Marca = "CA" then
           if   RCHEK = "S" then
''             CeldaKK = CeldaKK & ofv.GenCelda("","","","","","","","&nbsp;","si")
           end if
        end if
      end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      txx = ta

      xmpvin = "NO"

      if stip = "G" and Pdef(txx,9) = "SI" then
         xmpvin = "SI"
      end if

      if stip = "G" and MPvin  = "NO" then
         xmpvin = "NO"
      end if

     xPNumS = "NO"

     if stip = "G" then
        if Pnumer = "SI" Then
           select case Ucase(Pnumcls)
                case "SIMPLE"
                     xPnumS = "SI"
                case "COMPUESTA"
                     xPnumS = "SI"
           end select

        end if
      end if




      if PInden = "SI" then

            if  bsecix = 7 then
                bspan = "aaa colspan=2"
            else

                if stip <> "G" and stip <> "H" and stip <> "I"  and stip <> "J" then
                   wwcolsp =  10
                else
                   wwcolsp =  8
                   if (xPNumS = "NO"  and stip = "G") Then
                       wwcolsp = wwcolsp + 1
                      ''  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a1","si")
                   end if
                   if (xmpvin  = "NO"  and stip = "G") Then
                       wwcolsp = wwcolsp + 1
                     ''   Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a2","si")
                   end if

                end if
                Bspan = "aaa colspan=" & wwcolsp - bsecix

            end if
           CeldaKKp = ""
           for i=1 to bsecix
               Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
               if Marca = "CA" then
                  if   RCHEK = "S" then CeldaKKp = CeldaKKp & ofv.GenCelda("","","","","","","","&nbsp;","si")
               end if
           next
      else
            if stip <> "G" and stip <> "H" and stip <> "I"  and stip <> "J" then
                Bspan = "aaa colspan=10"
            else
                wwcolsp =  8
                if (xPNumS = "NO"  and stip = "G") Then
                    wwcolsp = wwcolsp + 1
               ''     Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a1","si")
                end if
                if (xmpvin  = "NO"  and stip = "G") Then
                    wwcolsp = wwcolsp + 1
              ''      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;a2","si")
                end if

                Bspan = "aaa colspan=" & wwcolsp
                if stip <> "G" then
                	 Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
                   if Marca = "CA" then
                      if   RCHEK = "S" then CeldaKKp = ofv.GenCelda("","","","","","","","&nbsp;","si")
                   end if
                end if
            end if
      end if


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

         if  stip <> "H" and stip <> "I"  and stip <> "J" then
             if PBorde <> "" and PBorde <> "NO" then
                PBCol = ""
                if Ucase(PBorde) = "SOLID" then PBCol = ";border-color:#000000"
                bspan = bspan & " STYLE={border-width:thin;border-style:" & PBorde & PBCol & ";} "
             end if
         end if

         PFtipoS = ""
         PFtipoF = ""
         if PFtipo <> "" then
            for VT = 1 to len(PFtipo)
                select case mid(PFtipo,vt,1)
                       case "1"
                            PFtipoS = PFtipoS & "<B>"
                            PFtipoF = PFtipoF & "</B>"
                       case "2"
                            PFtipoS = PFtipoS & "<I>"
                            PFtipoF = PFtipoF & "</I>"
                       case "4"
                            PFtipoS = PFtipoS & "<U>"
                            PFtipoF = PFtipoF & "</U>"
                       case "8"
                            PFtipoS = PFtipoS & "<S>"
                            PFtipoF = PFtipoF & "</S>"
                end select
            next
         end if

         pStyle = PFtipoS & "<Font face='" & pfont & "' color='" & pfcol
         pStyle = pStyle & "' STYLE='font-size:" & pfsize & "pt;'>"
         pStyleF = "</Font>" & PFtipoF

 ' TITULOS Y SUBTITULOS --------------------------------------------------

     if stip <> "G" and stip <> "H" and stip <> "I"  and stip <> "J" then
      'Descripcion
      PNumS = ""
      if Pnumer = "SI" Then
         select case Ucase(Pnumcls)
                case "SIMPLE"
                     select case PnumTip(tx)
                            case "1"
                                 BnumS = BnumG(bseci)
                            case else
                                 BnumS = ConvNumer(BnumG(bseci),PnumTip(tx))
                     end select

                     PnumS = BnumS & ". "
                case "COMPUESTA"
                     for i=0 to bseci
                     select case PnumTip(i + 7)
                            case "NO"
                                 BnumS = ""
                            case "1"
                                 BnumS = BnumG(i) & "."
                            case else
                                 BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
                     end select

                     PnumS = PnumS & BnumS
                     next
                     PnumS = PnumS & " "
         end select

      end if

      Celda = Celda & ofv.GenCelda(sPoint,bspan,Palign,"","",pbgcol,"",pStyle & PnumS  & ars(7,inx) & pStyleF,"si")



       if Marca = "CA" then

         if   RCHEK = "S" then
         	    Filatt = ""
      	      Celdakk = ""
      	      Celdakk = Celdakk & ofv.GenCelda("","utg colspan=12 style='text-align:justify;color:#0054a9;' ","","30%","","","","<u>Version Anterior:</u>","si")
       	      Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

         	    Srs37VKK = rs3(7)
              Celdakk = ""
              Celdakk = Celdakkp
''            CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & "dt valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"",Pnums  & Srs37VKK & "","si")
              CeldaKK = CeldaKK & ofv.GenCelda("",replace(bspan,"aaa","utg"),Ptalig & "utg valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"",Pnums & Srs37VKK & "","si")
      	      Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

         end if

       end if

   '   Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;x1","si")

'''''''''' arma tabla de indices '''''''''''''''''''''''''
      stexto1 = sfidkB & "&nbsp;&nbsp;" & replace((PnumS & ars(7,inx) & "")," ","&nbsp;")

      for nn=0 to (bseci+1)
          stexto1 = "&nbsp;&nbsp;&nbsp;&nbsp;" & stexto1
      next

      Celdat = stexto1

      stexto1 = "<A class=nt href='javascript:void(0);' "
    ''  stexto1 = stexto1 & " onclick=""creawin('T" & ars(0,inx) & "'); return false;"" "
      stexto1 = stexto1 & "   ><big>" & Celdat & "</big></a>"

      Celdat = "<td class=ut width=100% style='cursor:pointer;' onclick=""creawin('T" & ars(0,inx) & "');"" >" & stexto1 & "</td>"
      Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




     end if

 ' PARRAFOS DE TEXTO    --------------------------------------------------

     if stip = "G" then

        PNumS = ""
        if Pnumer = "SI" Then
           select case Ucase(Pnumcls)
                case "NO"
                     BnumG(6) = BnumG(6) - 1

                case "SIMPLE"
                     select case PnumTip(tx)
                            case "1"
                                 BnumS = BnumG(6)
                            case else
                                 BnumS = ConvNumer(BnumG(6),PnumTip(tx))
                     end select

                     PnumS = BnumS & ". "
                case "COMPUESTA"
                     for i=0 to 6
                         if BnumG(i) > 0 then
                            select case PnumTip(i + 7)
                              case "NO"
                                   BnumS = ""
                              case "1"
                                   BnumS = BnumG(i) & "."
                              case else
                                   BnumS = ConvNumer(BnumG(i),PnumTip(i + 7)) & "."
                            end select
                            PnumS = PnumS & BnumS
                         end if
                     next
                     PnumS = PnumS & " "
           end select

        end if

      tx = ta

      if  Ptalig = "" then
          Ptalig    = Pdef(tx,17)
          if Ptalig = "" then Ptalig = "left"
      end if

      if PSang = "" or PSang = " " then PSang = 0
      PSangS = ""

      for Vs = 1 to clng(PSang)
          PSangS = PSangS & "&nbsp;"
      next

      if Pdef(tx,9) <> "SI" then
         pVin = ""
      else
         PVFont  = Pdef(tx,15)
         PVtipo  = Pdef(tx,10)
         PVind   = Pdef(tx,11)

         if PVind = "" then PVind = 0

         PVindS = ""
         for Vi = 1 to pVind
             PVindS = PVindS & "&nbsp;"
         next

         pVin = PFtipoS & "<Font face='" & PVFont & "' color='" & pfcol
         pVin = pVin & "' STYLE={font-size:" & pfsize & "pt;}>"
         pVin = pVin & PVtipo & "</font>" & PFtipoS & PVindS  
      end if

      'Texto
      if MPVin <> "SI" then
         pVin = ""
      end if

      if PnumS <> "" then
         Celda = Celda & ofv.GenCelda("","aaa valign=TOP","","1%","","","",pStyle & PnumS  & pStyleF,"si")
      end if

      if PVin <> "" then
         Celda = Celda & ofv.GenCelda("","aaa valign=TOP","","1%","","","",pVin ,"si")
      end if

      sTextoV =  ars(8,inx)


      ' sTextoV =reemplazaspan(sTextoV)
  '    response.write "ssss" & len(sTextoV)






''''''''''''''''''''


'''''''''''''''''''

      sImagen  = "NO"
      wasotxt  = "N"


      'link relacionados al contenido
       descri=""
       wstextoLNK = ars(32,inx) & ""

       if not isnull(wstextoLNK) AND LEN(wstextoLNK) > 0 and wstextoLNK <> " " then

          if wstextoLNK <> "0" then

             wlink  = clng(wstextoLNK)
             strsql = "SELECT *  From objlnk where objid = " & wlink
            'response.write strsql
             set rst = ofv.crearconsultaEx(StrSql,cn,1,volver)
             if not rst.eof then

                wasotxt     = rst(14) & ""
                wasotxtstr  = CLNG("0" & rst(15))
                wasotxtcan  = CLNG("0" & rst(16))
                wasotxtdat  = rst(17) & ""

                IF wasotxt = "" THEN wasotxt = "N"



                wNV1    = rst(11) & ""
                wNV2    = rst(12) & ""
                wNV3    = rst(13) & ""

                IF  WNV1 = "" THEN WNV1 = "0"
                IF  WNV2 = "" THEN WNV2 = "0"
                IF  WNV3 = "" THEN WNV3 = "0"

                codobjetoL = clng(rst("codobjetoL"))
                wtipoLc    = rst(7)
                if trim(rst("Ltipo")) = "L" then

                   wUrlDesc = rst(9) & ""


                   wNomL    = wUrlDesc
                   if wNomL = "" then
                      wNomL = "S/D"


                      Strsql = "Select * From csGenObjetos where estado <> 4 and codobjeto = " & codobjetoL

                      set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rst2.eof then
                         wNomL = rst2(1)
                      else
                         wNomL =""' "S/D"
                      end if
                      call ofv.cerrarconsulta(rst2)

                   end if

                   sActcon = "Consultar.asp?Codigo=" & codobjetoL & "&Tipo=" & wtipoLc
                   sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                   sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3



	           sHref    = "&nbsp;"
	     ' response.write sActcon
	           sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)


	           slnktit  = wNomL
                   wNomL    = pStyle & "&nbsp;<u><i>" & wNomL & "</i></u>&nbsp;" & pStyleF
	           sImagen  = wNomL



	           sLink    = OFV.GenLink("", "nt title='" & slnktit & "' ",GenKeyLnks(sHref), "", "", sFuncion, "NNNNNN", "si")
                   descri   = sLink
                   call ofv.cerrarconsulta(rst2)

               else

                   if  not rst.eof then

                       wUrl  = rst("Lurl") & ""

                       if not isnull(wUrl) then
                          sHrefu = instr(wUrl,"://")
                          if sHrefu > 0 then
                             sHrefu = wUrl
                          else
                             sHrefu = "HTTP://" & wUrl
                          end if
                       else
                          sHrefu = "javascript:void(0);"
                       end if

	               slnktit  = rst("LurlDesc")
                       wNomL    = pStyle & "&nbsp;<u><i>" & rst("LurlDesc") & "</i></u>&nbsp;" & pStyleF
	               sImagen  = wNomL

                       descri = ofv.GenLink("","nt title='" & slnktit & "' ",GenKeyLnks(sHrefu),"_blank","","","NNNNNN","si")

                   end if

               end if


           end if
           call ofv.cerrarconsulta(rst)

         end if

      end if



     IF wasotxt = "S" THEN

      if len(sTextoV) > 0 then

         if  wasotxtstr > 0  and wasotxtcan > 0 then

            sTextoV = left(sTextoV,wasotxtstr - 1) & "" & replace(descri,"NNNNNN",pStyle & "&nbsp;<u><b style='color:" & Session("APLFT") & "' >" & mid(sTextoV,wasotxtstr,wasotxtcan) & "</b></u>&nbsp;" & pStyleF) & "" & right(sTextoV,len(sTextoV) - (wasotxtstr + wasotxtcan - 1))

         end if

         sTextoV =  replace(sTextoV,chr(13)," <br>")

      end if


     end if


      IF wasotxt = "N" THEN

         if len(sTextoV) > 0 then

             'sTextoV = supersec & "---" & ta& replace(sTextoV,chr(13)," <br>")
             sTextoV =  replace(sTextoV,chr(13)," <br>")

         end if

         if  sImagen <> "NO" then sTextoV = sTextoV & replace(descri,"NNNNNN",sImagen)

      end if


      if session("APLMDCLI")="BC" THEN
        Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP  ","","",pbgcol,"",pStyle & PsangS  & sTextoV & pStyleF,"si")
      ELSE
        Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;' ","","",pbgcol,"",pStyle & PsangS  & sTextoV & pStyleF,"si")
      END IF

       sTextoVKK = ""

       if Marca = "CA" then

           if  RCHEK = "S" then
           	   Filatt  = ""
             	 Celdakk = ""
      	       Celdakk = Celdakk & ofv.GenCelda("","utg colspan=12 style='text-align:justify;color:#0054a9;' ","","30%","","","","<u>Version Anterior:</u>","si")
        	     Filatt  = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

            	 sTextoVKK = rs3(8)
          	   Celdakk = ""
               Celdakk = Celdakkp
''             CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"",PsangS  & sTextoVKK,"si")
               CeldaKK = CeldaKK & ofv.GenCelda("",replace(bspan,"aaa","utg"),Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"",PsangS  & sTextoVKK,"si")
        	     Filatt  = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

           end if

       end if

      if PnumS = "" then
 '        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;x1","si")
      end if

      if PVin = "" then
   '      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;x2","si")
      end if

     end if

 ' IMAGENES            --------------------------------------------------

     if stip = "H" then


        sreduc = ""
        sbord = ""
        salng = ""

        if  ars(12,inx) <> "" and ars(12,inx) <> "100" then sreduc = " width=" & ars(12,inx) & "% "
        if  ars(15,inx) <> "" then
            if ars(15,inx) = "SI" then sbord = " border=1 "
        end if

        if  ars(16,inx) <> "" then
            if ars(16,inx) = "CENTRADO" then salng = " center"
            if ars(16,inx) = "DERECHA" then salng = " right"
            if ars(16,inx) = "IZQUIERDA" then salng = " left"
        end if

        if ars(9,inx) <> "" then

           slink = "<img src=" & chr(34) & GenKeyLnks("estructura/" & ars(9,inx))  & chr(34) & " "
           slink = slink & sreduc & sbord & " >"
           slink2 = slink

           slink = "<a href='" & GenKeyLnks("estructura/" & ars(9,inx))  &  "' target='_blank' >" & slink2 & "</a>"

        else
           slink = ""
        end if

        Celda = Celda & ofv.GenCelda("",bspan,salng,"","","","",slink,"si")
        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

        slinkkk = ""


        if Marca = "CA" then

           if   RCHEK = "S" then
           	  if rs3(9) & "" <> "" then
                 Filatt  = ""
                 CeldaKK   = ""
                 CeldaKK   = CeldaKK & ofv.GenCelda("","utg colspan=12 "," valign=TOP style='text-align:justify;color:#0054a9;' ","","","","", "<u>Version Anterior:</u>","si")
                 Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

                 CeldaKK   = ""
                 slinkkk = "<img src=" & chr(34) & GenKeyLnks("estructura/" & rs3(9))  & chr(34) & " "
                 slinkkk = slinkkk & sreduc & sbord & " >"
                 slink2kkk = slinkkk

                 CeldaKK   = CeldaKK & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"", slinkkk & "","si")
''               CeldaKK   = CeldaKK & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"", slinkkk & "","si")
                 Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

            end if

           end if

        end if

     end if

 ' Tablas           --------------------------------------------------

     if stip = "I" then

        slink = PasteTabla(ars(10,inx),nCodigo)
        Celda = Celda & ofv.GenCelda("",bspan,"","","","","",slink,"si")
        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

     end if

 ' COMPLEMENTOS            --------------------------------------------------

     if stip = "J" then

      'Descripcion
      sreduc = ""
      sbord = ""
      salng = ""
      if ars(12,inx) <> "" then sreduc = " width=" & ars(12,inx) & "% "
      if ars(15,inx) <> "" then
         if ars(15,inx) = "SI" then sbord = " border=1 "
      end if
      if ars(16,inx) <> "" then
         if ars(16,inx) = "CENTRADO" then salng = " center"
         if ars(16,inx) = "DERECHA" then salng = " right"
         if ars(16,inx) = "IZQUIERDA" then salng = " left"
      end if

      if ars(9,inx) <> "" then




          Strsql = "Select codobjetol,tipobjetol,* From objlnk Where ltipo = 'KL' and objid = 0" & ars(9,inx)

          set rsCompoM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsCompoM.eof then

           wtipoL   = rsCompoM(8)
	         wtipoLc  = rsCompoM(9) & ""

	         wNV1    = rsCompoM(13) & ""
	         wNV2    = rsCompoM(14) & ""
	         wNV3    = rsCompoM(15) & ""

	         IF  WNV1 = "" THEN WNV1 = "0"
	         IF  WNV2 = "" THEN WNV2 = "0"
	         IF  WNV3 = "" THEN WNV3 = "0"



              Strsql = "Select nombre,TIPOCOMP From csacomple Where codcomp = " & rsCompoM(0)
             ' response.write strsql & chr(13)
              set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              if  not rsComponentes.eof then
                  sComponente = rsComponentes(0)
                  sTIPOCMP    = rsComponentes(1)

              end if
              call ofv.cerrarconsulta(rsComponentes)


             sInput = sComponente

             if len(sInput) > 50 then
                 sInput = left(sInput,47) & "..."
             end if

             sARCHIVO = ""
             sARCHmas = "N"
'            Strsql = "SELECT   * FROM COMPIMAG where codcomp = 0" & rsCompoM(0) & " ORDER BY SECUENCIA, OBJID "

             Nv=WNV3
             if nv=0 then
               Nv=wnv2
                if nv=0 then
                 nv=wnv1
                end if
             end if

             Strsql = "SELECT   * FROM COMPIMAG where objid = 0" & NV & " ORDER BY SECUENCIA, OBJID "

            ' response.write strsql
             set rsComponentes = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             do until rsComponentes.eof
                sARCHIVO = rsComponentes(3)
                rsComponentes.movenext
                if not rsComponentes.eof then
                   sARCHmas = "S"
                   exit do
                end if
             loop
             call ofv.cerrarconsulta(rsComponentes)


             sInput = "&nbsp;&nbsp;<B>" & replace(sInput," ","&nbsp;") & "&nbsp;" & "</b>"

             if sARCHmas = "S" then

                sInput = sInput & "&nbsp;&nbsp;(mas imagenes)"
             end if

             sImage = "<font face='corflaticon' size=5 color=#f7aa22 >N</font>" & sInput
              sHref       = "EditarImagen2.asp?ID=" & wtipoL
              sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

             sHref = GenKeyLnks(sHref)
             ' sHref = GenKeyLnks("imagencont4.asp?ID=" & rsCompoM(0))
             sLink2 = ofv.GenLink("","NT title='Visualizar el componente' ",sHref,"_blank","","",sImage,"si")

             slink = "<BR>" & sLink2

             IF SARCHIVO <> "" THEN
                slink = slink & "<BR><BR><img src=" & chr(34) & GenKeyLnks("COMPONENTES/" & SARCHIVO)  & chr(34) & " "
                slink = slink & sreduc & sbord & " >"
             END IF

             slink = slink & "<BR>"



          else
             slink = ""

          end if
          call ofv.cerrarconsulta(rsCompoM)







      else
         slink = ""
      end if

      Celda = Celda & ofv.GenCelda("",bspan,salng,"","","","",slink,"si")
      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")

      slinkxx = ""

      if Marca = "CA" then
         sImage = "<font face='corflaticon' size=5 color=#f7aa22 >N</font>" & sInput
         sHref       = "EditarImagen2.asp?ID=" & wtipoL
         sHref       = sHref & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3

         sHref = GenKeyLnks(sHref)
         sLink2xx = ofv.GenLink("","NT title='Visualizar el componente' ",sHref,"_blank","","",sImage,"si")

             slinkxx = "<BR>" & sLink2xx

             IF SARCHIVO <> "" THEN
                slinkxx = slinkxx & "<BR><BR><img src=" & chr(34) & GenKeyLnks("COMPONENTES/" & SARCHIVO)  & chr(34) & " "
                slinkxx = slinkxx & sreduc & sbord & " >"
             END IF

             slinkxx = slinkxx & "<BR>"
             if   RCHEK = "S" then
             	Filatt  = ""
                 	Celdakk = ""
           	      Celdakk = Celdakk & ofv.GenCelda("","utg colspan=3 style='text-align:justify;color:#0054a9;' ","","30%","","","","<u>Version Anterior:</u>","si")
                  Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")
                  Celdakk = ""
''                CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & "dt valign=TOP style='text-align:justify;color:#0054a9;border-style:solid;border-width:1;border-color:#0054a9;' ","","",Pbgcol,"", slinkxx & "","si")
                  CeldaKK = CeldaKK & ofv.GenCelda("",bspan,Ptalig & "dt valign=TOP style='text-align:justify;color:#0054a9;' ","","",Pbgcol,"", slinkxx & "","si")
                  Filatt = Filatt & ofv.GenRow("","ut","","","","","",Celdakk,"si")

             end if

      end if

    end if

'------------------------------------------------


         select case stip
                case "A"
                     sant     = "A"
                     sVinoAnt = sant

                case "B"
                     sant     = "B"
                     sVinoAnt = sant
                case "C"
                     sant     = "C"
                     sVinoAnt = sant
                case "D"
                     sant     = "D"
                     sVinoAnt = sant
                case "E"
                     sant     = "E"
                     sVinoAnt = sant
                case "F"
                     sant     = "F"
                     sVinoAnt = sant
                case else

                     sVinoAnt = stip

          end select


      Fila = ofv.GenRow("","","","","","","",celda,"si")


      wcantlin = wcantlin + 1

      if   wcantlin > 20 then
           wcantlin = 0
      end if



    if Mostrar <> "NO"  Then

        FNSHTML = FNSHTML &  Fila



 ''response.write marca & " 03"

      if PEspa = "" then
         PEspa      = 0
         PEspHeight = ""
      elseif PEspa = 0 then
         PEspa      = 0
         PEspHeight = ""
      else
         PEspa      = 0 + PEspa
         PEspHeight = 9 * PEspa
      end if



      celda = ""
      for i=1 to 12
          Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      next
      Fila = ofv.GenRow("","","","",PEspHeight,"","",celda,"si")
      FNSHTML = FNSHTML &  Fila


      if sVinoAnt > "F" and stip < "G" then
         celda = ""
         for i=1 to 10
             Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
         next
         Fila  = ofv.GenRow("","","","","12","","",celda,"si")
         FNSHTML = FNSHTML &  Fila

      end if

       if Marca = "CA" then
          if  RCHEK = "S" then
''            FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Celdakk,"si")
      celda = ""
      for i=1 to 12
          Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      next
      Filatt = Filatt & ofv.GenRow("","","","","1","","",celda,"si")

      SArmatt = ofv.GenTabla("","AAA style='BACKGROUND-COLOR:#f9f9f9;border-style:solid;border-width:1;border-color:#0054a9;'  ","","80%","","0","0","2","2","","",Filatt,"si")

      Celdakk = ""
      Celdakk = Celdakk & ofv.GenCelda("","ut colspan=12 align=CENTER ","","","","","",SArmatt,"si")
    ''  Fila = Fila & ofv.GenRow("","","","","","","",Celdakk,"si")


       FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Celdakk,"si")
''  FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Fila,"si")

        Filatt = ""

          end if
        end if


    end if

''    FNSHTML = FNSHTML &  ofv.GenRow("","utg","","30","","","",Fila,"si")


 next

         if inx > UBound(ars,2) then
            exit do

         end if

      ssale = "no"
      do until ssale = "si"

         if     sant = "A" then

            sant = "X"
         elseif sant = "B" then

            sant = "A"
         elseif sant = "C" then

            sant = "B"
         elseif sant = "D" then

            sant = "C"
         elseif sant = "E" then

            sant = "D"
         elseif sant = "F" then

            sant = "E"
         end if

      if sant = "X" then
         ssale = "si"
      end if

      loop

 loop

      celda = ""
      for i=1 to 12
          wwidth = "1%"
          if i = 9 then wwidth = "90%"
          Celda = Celda & ofv.GenCelda("","ut","",wwidth,"","","","&nbsp;","si")
      next

      Fila = ofv.GenRow("","","","","","","",celda,"si")
  FNSHTML = FNSHTML &  Fila




' call ofv.cerrarconsulta(rs)


  celda = ""




  FNSHTML = FNSHTML & ofv.GenFTabla()


''
  GENFMTESTRUCTURA = "<div id=zglos>" & FNSHTML & "</div>"


 END FUNCTION



'''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''
function ConvNumer(CNumero,CTipo)

CLetra = "abcdefghijklmnopqrstuvwxyz"

dim CRuni(9)
CRuni(1) = "I"
CRuni(2) = "II"
CRuni(3) = "III"
CRuni(4) = "IV"
CRuni(5) = "V"
CRuni(6) = "VI"
CRuni(7) = "VII"
CRuni(8) = "VIII"
CRuni(9) = "IX"

dim CRdec(9)
CRdec(1) = "X"
CRdec(2) = "XX"
CRdec(3) = "XXX"
CRdec(4) = "XL"
CRdec(5) = "L"
CRdec(6) = "LX"
CRdec(7) = "LXX"
CRdec(8) = "LXXX"
CRdec(9) = "XC"

dim CRcen(9)
CRcen(1) = "C"
CRcen(2) = "CC"
CRcen(3) = "CCC"
CRcen(4) = "CD"
CRcen(5) = "D"
CRcen(6) = "DC"
CRcen(7) = "DCC"
CRcen(8) = "DCCC"
CRcen(9) = "CM"


select case Ctipo
       case "a","A"
          Ci = 0
          CR = CNumero

             for ci=0 to 28
                 if CR < 28 then exit for
                 CR = CR - 27
             next

          if ci > 27 then
             ConvNumer = "XX"
          else
             if ci > 0 then ConvNumer = mid(CLetra,ci,1)
             ConvNumer = ConvNumer & mid(CLetra,CR,1)
          end if
          if Ctipo = "A" then ConvNumer = Ucase(ConvNumer)

       case "I","i"
          CR = CNumero
          Ci = len(CR)
          if Ci < 1 or ci > 3 then
             ConvNumer = "AAAA"
          else
             cx = 1
             for CL=ci to 1 step -1
             select case cx
                    case 1
                         ConvNumer = CRuni(mid(cr,cl,1))
                    case 2
                         ConvNumer = CRdec(mid(cr,cl,1)) & ConvNumer
                    case 3
                         ConvNumer = CRcen(mid(cr,cl,1)) & ConvNumer
             end select

             cx = cx + 1
             next
             if Ctipo = "i" then ConvNumer = Lcase(ConvNumer)
          end if
end select


end function


function reemplazaspan9(textoz)

 jsspan=""

    Dim Position
    Dim WordLength
    Word="<SPAN id="
        Position = InStr(1, Textoz, Word,0)

        WordLength = Len(Word)
        Do While Position


             idd= InStr(Position + WordLength, Textoz," ",0)
           'idd= InStr(Position + WordLength, Textoz," ",0)
             total=idd- ( position+ + WordLength )
             idPerfil =mid(textoz,Position + WordLength,total)
              sql="select a.descri from Tipoperfil a inner JOIN ecoperfil b on( b.codperfil=a.tipoPERfil and b.objid=0" & idPerfil & ")"

       '       RESPONSE.WRITE SQL
            set xpf=ofv.crearconsultaEx(sql,cn,1,parametros)
             if not xpf.eof then
                sfin=InStr(idd, Textoz,"</SPAN>",0)
                 xini=InStr(idd, Textoz,"xspanx>",0)
       '         TEXTOLD=MID(TEXTOZ,xini+ len("xspanx>"),SFIN-xini-1)
        '        response.write textold & "222"
            '   Textoz=REPLACE(Textoz,TEXTOLD,xpf(0))
                jsspan = jsspan & " document.getElementById('"&idPerfil &"').innerHTML='"&xpf(0)&"';" & chr(13)
          		 end if
             Position = InStr(Position + WordLength, Textoz,Word,0)
        Loop
        jsspan =jsspan & ""
  reemplazaspan = textoz
end function

%>