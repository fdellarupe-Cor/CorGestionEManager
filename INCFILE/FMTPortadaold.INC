
<%

 FUNCTION GENFMTPortada(FNCODOBJ,FNTIPOBJ)


 dim STab (4,5,10)


 for yyy=1 to 4

      for zzz=1 to 5

           STab (yyy,zzz,1) = "N"
 
      next

 next


 pStyle  = ""
 pStylef = ""
 
  FNSHTML = ""


  FNSHTML = FNSHTML & "<br><br><br>"


  xsfont = "arial"
  xsfontSZ = "8pt"
  xsfontCL = ""
  xsfondo = ""
  xsborde = ""
  xsalign = "left"
  xsresal = ""

  xsstyle = " ut "
  xsalinea = "left"
  xsfondea = ""

xsfontw1 = ""
xsfontw2 = ""
xsfontw3 = ""
xsfontw4 = ""

  strsql = "SELECT * from Objportada where  OBJportada.codobjeto=" & FNCODOBJ & " and  OBJportada.detcon = 0 and OBJportada.detsecuen = 99"
 '' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    if rs("detdefault") = "NO" then
       PFtipo = rs("detftipo") & ""

         if PFtipo <> "" then
            for VT = 1 to len(PFtipo)
                select case mid(PFtipo,vt,1)
                       case "1"
                            PFtipoS = PFtipoS & "<B>"
                            PFtipoF = PFtipoF & "</B>"
                            xsfontw1 = " font-weight:'bold'; "
                       case "2"
                            PFtipoS = PFtipoS & "<I>"
                            PFtipoF = PFtipoF & "</I>"
                            xsfontw2 = " font-style:'italic'; "
                       case "4"
                            PFtipoS = PFtipoS & "<U>"
                            PFtipoF = PFtipoF & "</U>"
                            xsfontw3 = " text-decoration:'underline'; "
                       case "8"
                            PFtipoS = PFtipoS & "<S>"
                            PFtipoF = PFtipoF & "</S>"
                            xsfontw4 = " text-decoration:'line-through'; "
                end select
            next
         end if


  xsfont = " font-family:'" & rs("detfont") & "'; "
  xsfontSZ = " font-size:'" & rs("detfsize") & "pt';"
  xsfontCL = " color:'" & rs("detfcolor") & "'; "
  xsfondo = rs("detbgcolor") & ""



  xsbordex = rs("detborde") & ""



  if xsbordex <> "" then
     xsborde = " style="  & chr(34) & "border-style:" & xsbordex & ";"

     if xsbordex <> "Solid" then
        xsborde = xsborde & "border-width:3;"
     else
        xsborde = xsborde & "border-width:1;border-color:'888888';"
     end if 

     xsborde = xsborde & chr(34) & " "
  end if

 '' xsalign = rs("detalign") & ""

  xsalign = rs("dettextalign") & ""

''  xsresal = rs("dettextalign") & ""

      xsstyle = " ut style="  & chr(34) & xsfont & xsfontSZ & xsfontCL & xsfontw1 & xsfontw2 & xsfontw3 & xsfontw4 & chr(34) 
      xsalinea = xsalign
      xsfondea = " bgcolor='" & xsfondo & "' "


    end if

  ''  response.write xsstyle
 ''  response.write xsalign & "Z"

  end if



  Strsql = "SELECT OBJportada.detcon, OBJportada.detsecuen, OBJportada.dettexto, OBJportada.detimagen, OBJportada.rellink,"
  Strsql = Strsql & " OBJportada.detfsize, OBJportada.detborde, OBJportada.detftipo, OBJportada.dettipo from OBJportada "
  Strsql = Strsql & " where OBJportada.codobjeto=" & FNCODOBJ & " and OBJportada.detcon > 0 "
  Strsql = Strsql & " order by OBJportada.detcon, OBJportada.detsecuen "
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
           yyy = rs(0)
           zzz = rs(1)
           STab (yyy,zzz,1) = rs(8) & ""
           STab (yyy,zzz,2) = rs(2)
           STab (yyy,zzz,3) = rs(3)
           STab (yyy,zzz,4) = rs(4)
           STab (yyy,zzz,5) = rs(5)
           STab (yyy,zzz,6) = rs(6)
           STab (yyy,zzz,7) = rs(7)
           STab (yyy,zzz,8) = "javascript:void(0);"          


   rs.movenext
  loop


FNSHTMLx = ""

  for yyy=1 to 4

      Celdax = ""

      for zzz=1 to 5

       sclase = "ut"
       sInput="&nbsp;"

       if    STab (yyy,zzz,1) = "S" then
           '  Ptalig    = Pdef(tx,17)
             if Ptalig = "" then Ptalig = "left"

             sTextoV =  STab (yyy,zzz,2)


             sImagen  = "NO"
             wasotxt  = "N"
             
        

      'link relacionados al contenido
             descri=""
             wstextoLNK = STab (yyy,zzz,4) & ""
             

             if not isnull(wstextoLNK) AND LEN(wstextoLNK) > 0 and wstextoLNK <> " " then

                if wstextoLNK <> "0" then

             wlink  = clng(wstextoLNK)
             strsql = "SELECT *  From objlnk where objid = " & wlink
         ' response.write strsql
             set rst = ofv.crearconsultaEx(StrSql,cn,1,volver)
              xdisplay=""
             if not rst.eof then
                xdisplay=" style={display:none;visibility:hidden;}"
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

               '''  sdestino = FNCODOBJ & ""


                 if cstr(codobjetoL & "")  = cstr(FNCODOBJ & "") then 


 '''' BarraSol.asp?Tipo=PCD&Codigo=2131&Menu=&Sector=&MAINAPL=EM&Funcion=CD&SSCLICK=7&NIV1=7&NIV2=409&NIV3=0&NIV4=0&NIV5=0

	           sHref = "BarraSol.asp?Tipo=" & FNTIPOBJ & "&Codigo=" & FNCODOBJ & "&Menu=&Sector=&MAINAPL=EM&Funcion=CD&SSCLICK=" & WNV1 & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3 & "&NIV4=0&NIV5=PORTADA"
                    sHref = sHref & "&ORIGEN=PORTADA"
	           sFuncion = " "
	           sTarget  = "WK1"
                     
                   
                   if genlnk="NO" THEN 
		                      sFuncion=""
		                      sHref=""
                   END IF
                   STab (yyy,zzz,8)=sHref

                 else

                   sActcon = "Consultar.asp?Codigo=" & codobjetoL & "&Tipo=" & wtipoLc
                   sActcon = sActcon & "&Datos=*smenu=PORESTADO*lista=3"
                   sActcon = sActcon & "&NIV1=" & WNV1 & "&NIV2=" & WNV2 & "&NIV3=" & WNV3
                    sActcon = sActcon & "&NIV5=PORTADA"
	           sTarget   = ""

	           sHref    = "&nbsp;"
	     ' response.write sActcon
	           sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;" & Chr(34)

                 end if

	           slnktit  = wNomL
                   wNomL    = pStyle & "&nbsp;<u><i>" & wNomL & "</i></u>&nbsp;" & pStyleF
	           sImagen  = wNomL
                  
                   
                   if genlnk="NO" THEN 
                      sFuncion=""
                       sHref=""
                   END IF

	           sLink    = OFV.GenLink("", xsstyle & " title='" & slnktit & "' ",sHref, sTarget, "", sFuncion, "NNNNNN", "si")
                   descri   = sLink
                   call ofv.cerrarconsulta(rst2)
                   STab (yyy,zzz,8)="AbrirCons('" & sActcon & "','" & wtipoLc & "'); return false;"

               else

                   if  not rst.eof then

                       wUrl  = rst("Lurl") & ""

                       if not isnull(wUrl) then
                          sHrefu = instr(wUrl,":")
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
                       
                       if genlnk="NO" THEN  sHrefu = ""
                       descri = ofv.GenLink("",xsstyle & " title='" & slnktit & "' ",GenKeyLnks(sHrefu),"_blank","","","NNNNNN","si")
                       
                       
                       STab (yyy,zzz,8)=GenKeyLnks(sHrefu)
                   end if

               end if


           end if
           call ofv.cerrarconsulta(rst)

         end if

      end if

wasotxtstr = 0
wasotxt = "N"

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

       '''  if  sImagen <> "NO" then sTextoV = sTextoV & replace(descri,"NNNNNN",sImagen)


            sTextoV = replace(descri,"NNNNNN",pStyle & sTextoV & pStyleF) 



      end if


      if session("APLMDCLI")="BC" THEN
   '     Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP  ","","",pbgcol,"",pStyle & sTextoV & pStyleF,"si")
      ELSE
   '     Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;' ","","",pbgcol,"",pStyle & PsangS  & sTextoV & pStyleF,"si")
      END IF



        sreduc = ""
        sbord = ""
        salng = ""

        if  STab (yyy,zzz,5) <> "" and STab (yyy,zzz,5) <> "100" then sreduc = " width=" & STab (yyy,zzz,5) & "% "
        if  STab (yyy,zzz,6) <> "" then
            if STab (yyy,zzz,6) = "SI" then sbord = " border=1 "
        end if

        if  STab (yyy,zzz,7) <> "" then
            if STab (yyy,zzz,7) = "CENTRADO" then salng = " center"
            if STab (yyy,zzz,7) = "DERECHA" then salng = " right"
            if STab (yyy,zzz,7) = "IZQUIERDA" then salng = " left"
        end if

        if STab (yyy,zzz,3) <> "" then

           slink = "<img  src=" & chr(34) & GenKeyLnks("estructura/" & STab (yyy,zzz,3))  & chr(34) & " "
           slink = slink & sreduc & sbord & " >"
           slink2 = slink

         ''  slink = "<a href='" & GenKeyLnks("estructura/" & STab (yyy,zzz,3))  &  "' target='_blank' >" & slink2 & "</a>"


           slink = replace(descri,"NNNNNN",slink2)
           

        else
           slink = ""
        end if

          slink = ""

           slinkx = replace(descri,"NNNNNN"," ")
           slinkx = replace(slinkx,"<A ","<A id=" & zzz & "_" & yyy )
           

           srow = ""
           'Celdax1 = ofv.GenCelda("",xsstyle  ,"center id=" & zzz & "_" & yyy ,"","","","",descri,"si")
           Celdax1 = ofv.GenCelda("",xsstyle  ,salng ,"","","","",slinkx,"si")
           srow = srow & ofv.GenRow("","","","","","","",Celdax1,"si")

           Celdax1 = ofv.GenCelda("",xsstyle,xsalinea,"","","","",pStyle & sTextoV & pStyleF,"si")
           srow = srow & ofv.GenRow("","","","","","","",Celdax1,"si")

            sInput = ofv.GenTabla("","aaa " ,"","","","0","","","","","",srow,"si")
          ' response.write sTextoV & "<br>"


           ' sInput = slink & "<br>" & pStyle & sTextoV & pStyleF

            sclase = "ut  " & xsborde

       end if 

       Celdax = Celdax & ofv.GenCelda("",sclase,"center","20%","","","",sInput,"si")



     next

        FNSHTMLx = FNSHTMLx & ofv.GenRow("","","","","","","",Celdax,"si")





  next





  FNSHTML = FNSHTML & "<br><center>" & ofv.GenTabla("","aaa  " & xsborde & xsfondea,"","80%","","0","","10","2","","",FNSHTMLx,"si") & "</center>"






'''''''''''''''''



  call ofv.cerrarconsulta(rs)

SHTMLIMG="<SCRIPT>" & CHR(13)
SHTMLIMG=SHTMLIMG & "var newImage;" & CHR(13)
for yyy=1 to 4

      Celdax = ""

      for zzz=1 to 5

      

       if    STab (yyy,zzz,1) = "S" then
           SHTMLIMG=SHTMLIMG & "  newImage=document.createElement('img');" & CHR(13)
         
           SHTMLIMG=SHTMLIMG & "newImage.src='estructura/" & STab (yyy,zzz,3) &"'" & CHR(13)
           XSIZE=TRIM(CSTR(STab (yyy,zzz,5) &""))
          ' RESPONSE.WRITE XSIZE ="" & "<BR>"
           IF XSIZE <> ""  AND XSIZE <> "100" THEN
           SHTMLIMG=SHTMLIMG & " newImage.width='"& XSIZE & "';"& CHR(13)
           SHTMLIMG=SHTMLIMG & " newImage.height='"& XSIZE & "';"& CHR(13)
           END IF
           if STab (yyy,zzz,6) = "SI" then 
           SHTMLIMG=SHTMLIMG & " newImage.setAttribute('border','1');"
           end if
           
           'SHTMLIMG=SHTMLIMG & " newImage.setAttribute('onclick'," &chr(34)  &  STab (yyy,zzz,8)  & chr(34) &");"
           'SHTMLIMG=SHTMLIMG & " newImage.setAttribute('style',cursor: hand;")
           SHTMLIMG=SHTMLIMG & "var elem2 = document.getElementById('" & ZZZ & "_" & YYY &"');"& CHR(13)
            'SHTMLIMG=SHTMLIMG & "elem2.setAttribute('href', '#');" & CHR(13)
   	   SHTMLIMG=SHTMLIMG   & " if(elem2 != null) { elem2.appendChild(newImage);}"& CHR(13)
     END IF
       
      NEXT
NEXT      
       
     SHTMLIMG=SHTMLIMG & "</SCRIPT>"& CHR(13)
  GENFMTPortada = FNSHTML & SHTMLIMG


 END FUNCTION



%>