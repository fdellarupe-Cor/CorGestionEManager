<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv 
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok 
else   
  response.expires=0   

  origen   = request.querystring("origen")
  nCodigo = request.querystring("Codobjeto")   
  sTipo   = request.querystring("Tipobjeto") 
  sdocum  = request.querystring("sdocum") 
  sdocc   = request.querystring("sdocc")      
  sdep    = request.querystring("sdep")       
  sNIV1   = request.querystring("NIV1")      
  sNIV2   = request.querystring("NIV2")     
  sNIV3   = request.querystring("NIV3")    
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")
     
   
  vuelta = ofv.vueltaasp("../",CantRegMostrar,CantRegMover) & "_ID=" & ID
''response.write request.querystring & "<br>"

  dim Adep(7)
  dim Adepi(7)
  
  Adepi(1) = "A"    
  Adepi(2) = "B"
  Adepi(3) = "C"   
  Adepi(4) = "D"   
  Adepi(5) = "E"       
  Adepi(6) = "F" 
  Adepi(7) = "I"   

  for ix = 0 to 7
      Adep(ix) = 0
  next    
 

  sfidc  = "<font face=corflaticon size=2 color=" & Session("CMFPNGNFT") & ">0</font>" 
  sfida  = "<font face=corflaticon size=2 color=" & Session("CMFPNGRFT") & ">1</font>" 
  sfidk  = "<font face=corflaticon size=2 color=#f7aa22 >t</font>" 

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

 shtml = shtml & CorPnlFnc("N","")
   
       
'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK 
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "11" AND SNIV2 <> "0" THEN 

    SDOCC =  SNIV2

    IF SNIV3 <> "0" THEN

       SPLFUNC = "<script language='javascript' > creawin('T" & SNIV3 & "');</script>"

    ELSE

       SPLFUNC = ""

    END IF             
    
   
 ELSE

    SPLFUNC = ""    

 END IF     

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"


'''''''''''''''''''''''''''''''''''''''''''''''''
    

  fila = ""

  set cn = ofv.conectar(ofv.strconn2) 

' Version

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
    
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO" 
  if not rs.eof then
    if rs(0) = "V" then 
       bVersion = "SI"
    end if    
   Nestado=rs(1) 
  end if 
  call ofv.cerrarconsulta(rs)

 
  sVall = "S"

       if sdocc = "" then
          sdocc = cstr(sdocum)
          if sdocc = "" then
             sVall = "S"
             Strsql = "SELECT * FROM OBJESTR WHERE codobjeto = " & nCodigo
             Strsql = Strsql & " order by CONsecuen "
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if not rs.eof then 
               sdocc   = cstr(rs(0))
               sdocum  = cstr(rs(0))
            end if
          end if
       end if

   '   response.write  sdocc & " " & sdocum & "<br>"

' Ubicar la Dependencia

  if (not sdep = "") and (not sdep = "C") then
     Strsql = "SELECT * FROM OBJESTDET WHERE objid = " & sdep
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        sdeptipo = rs(5)
        sdepdepen = rs(6)
        For ix = 1 to 6
            if Adepi(ix) = sdeptipo then 
               Adep(ix) = rs(0)
               Adep(ix - 1) = sdepdepen
               exit for
            end if
        next
    '    response.write  sdeptipo & " " & sdepdepen & " ix " & ix & " adep " & Adep(ix) & "<br>"
        do until sdeptipo = "A" or ix > 6 or rs.eof
           Strsql = "SELECT * FROM OBJESTDET WHERE objid = " & sdepdepen
           set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs.eof then
              sdeptipo = rs(5)
              sdepdepen = rs(6)
              For ix = 1 to 6
                  if Adepi(ix) = sdeptipo then 
                     Adep(ix) = rs(0)
                     Adep(ix - 1) = sdepdepen
                     exit for
                  end if
              next
    '    response.write  sdeptipo & " " & sdepdepen & " ix " & ix & " adep " & Adep(ix) & "<br>"
           end if
        loop 
     end if
     call ofv.cerrarconsulta(rs)
  end if

  Filat  = ""

  dblinx = "N"



' Uno o Todos los Documentos
   

    Strsql = "SELECT * FROM OBJESTR WHERE codobjeto = " & nCodigo
    if not sdocum = "" and sVall <> "S" then
       Strsql = Strsql & " and objid = " & sdocum
    end if
    Strsql = Strsql & " order by CONsecuen "

    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs.eof then 

       dblinx = "S"
 
       if sdocc = "" then
          sdocc = cstr(sdocum)
          if sdocc = "" then
             sdocc = rs(0)
             sdocum  = rs(0)
          end if

       else



       do until rs.eof

       '   if cstr(rs(0)) <> sdocc then 

          if cstr(rs(0)) = sdocc then 
  
             sfidkD = sfidkC

          else

             sfidkD = sfidkA

          end if  


             stexto1 = "<A ID='ATIT" & rs(0) & "' NAME='ATIT" & rs(0) & "' class=nt href='coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
             stexto1 = stexto1 & "&sdocum=" & sdocum & "&sdocc=" & rs(0)
             stexto1 = stexto1 & "'   >"  & rs(3) & "</a>"

            ' Celdat = "<td class=ut style='cursor:pointer;' onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
            ' Filatt  = filatt & ofv.GenRow("","","","","","","",Celdat,"si")

             Celdat  = "<td class=utis >" & sfidkD & "</td>"
             Celdat  = Celdat & "<td width=95% class=uti style='cursor:pointer;' onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
             Filatt  = filatt & ofv.GenRow("","","","","","","",Celdat,"si")


       '   end if 
       rs.movenext
       loop



       end if




      Celda = ""

      if sdocc = "" then


         ID = sdocc
         SID = SID
         Spmk = ""
         if sID <> "" then mostrar = "NO"
         Marca = "CA"


      else

         if sdep = "C" then
            Spmk    = ""
            Marca   = "CA"
         else
            Marca   = "SI"
         end if

         ID = sdocc
         SID = SID
         if sID <> "" then mostrar = "NO"


      end if



    else

       sestr  = "<font color=#336699><h5>Sin&nbsp;datos&nbsp;disponibles"
       sestr  = sestr & "</h5></font>" 
       Celdat = ofv.GenCelda("","","","1%","","","",sestr,"si")
       Filat  = ofv.GenRow("","","","","","","",Celdat,"si")

    end if        
  
    call ofv.cerrarconsulta(rs) 






  if ID <> "" then

  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

''''''''''''''''
   

  if Marca = "" then Marca = "NO"


'----------------------------------------------
     sinput2 = "" & chr(34)


  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

dim Pdef(13,18)

  strsql = "Select * From SysEstprty"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  t = 0
  do until rs.eof
  for ti = 1 to 18
      if (rs(ti) & "") = "" then
         sti = ""
      else
         sti = rs(ti) 
      end if

      Pdef(t,ti - 1) = sti

  next
  rs.movenext
  t = t + 1
  loop

  call ofv.cerrarconsulta(rs)
   
  strsql = "Select * From Objestr Where objid = " & ID
 ''' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bsecuen = ""
  bDescri = " "
  bobj = ""
  if not rs.eof then
     bsecuen = rs(4)
     bDescri = rs(3)
     bobj = rs(0)

     stexto1 = "<A ID='ATIT" & rs(0) & "' NAME='ATIT" & rs(0) & "' "  & chr(13)
     stexto1 = stexto1 & " class=nt href='coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
     stexto1 = stexto1 & "&sdocum=" & sdocum & "&sdocc=" & rs(0)
     stexto1 = stexto1 & "'   >" & sfidkc & "&nbsp;<big>" & rs(3) & "</big></a>" & chr(13)

     Celdat = "<td class=ut style='cursor:pointer;'  onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
     Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si") & chr(13)


  end if
  call ofv.cerrarconsulta(rs)



  strsql = "Select * From ObjEstPrty where detcon = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  t = 0
  do until rs.eof
  select case rs(4)
             case "A"
                  t = 0
             case "B"
                  t = 1
             case "C"
                  t = 2
             case "D"
                  t = 3
             case "E"
                  t = 4
             case "F"
                  t = 5
             case "G"
                  t = 6
   end select

  for ti = 5 to 22
  if (rs(ti) & "") = "" then
     sti = ""
  else
     sti            = rs(ti)
     Pdef(t,ti - 5) = sti
  end if



  next
  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)

  PInden = Pdef(0,7)
  if PInden <> "NO" then PInden = "SI"

  PNumer = Pdef(0,12)
  if PNumer <> "NO" then PNumer = "SI"

  PNumCls = Pdef(0,14)
  redim preserve PNumTip(13)
  for i=0 to 13
      PNumTip(i) = Pdef(i,13)
      if PNumTip(i) = "" then PNumTip(i) = "1"
  next

      pmk = "<form name="  & chr(34) & "D" & bobj  & chr(34) &  " method=post >"
      pmk = pmk & "<input type=hidden value=a ></form>"

   'Estructura


   if marca = "CA" then
      Celda =  ofv.GenCelda("","ttt","","90%","","","",sfidc & "&nbsp;Cambios en " & bDescri,"si")
   else
      Celda =  ofv.GenCelda("","ttt","","90%","","","",pmk & "&nbsp;","si")
   end if

   Celt2 =  ofv.GenCelda("","ut","","1%","","","",t2,"si")
   Celt1 =  ofv.GenCelda("","ut","","1%","","","",t1,"si")


  Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100% border=0 ","","0","","0","0","","",Fila,"si")

  styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"
  sHTML = sHTML &  ofv.GenTabla("","aaa  ","","100%  ","","0","","0","0","","","","no")





  Fila  = ""



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

'response.write strsql

Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.CursorType = 3
Rs.CacheSize = 100
 
Rs.Open strsql, cn

' ars = rs.GetRows()
if not rs.eof then 
  ars = rs.GetRows()
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
      if bVersion = "SI" and smodi and SESSION("SHWMOD") <> "N" then





      if stipmodi = "M" then
         xfclas = "btbs"
         xftitl = "Modificado."
         sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
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
      

        
      if Nestado=5  and origen <> "ED"  then 

      	  '     sAccion = "genObs.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	  '     sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & sclave  & "&ID2=" & ID
      
      	  '     xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      	  '     sHref   = sAccion
      	  '     sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	  '     Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Observacion de " & BtipP & bpespec & "' ","","","1","","",sInput,"si")
          '     Celdaobs  = ofv.GenLink("",xclas & " title='Obs de " & BtipP & bpespec & "' ",sHref,"","","",Celdaobs,"si")

               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & sclave  & "&ID2=" & ID
      
      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")




      end if

      
      Celda  = Celdaobs & ofv.GenCelda("","uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")


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
                              ta = 8
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
                  bsec(7) = ars(4,inx)

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
               PFcol  = "#aaaaaa"
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

      if Marca <> "SI" and Marca <> "CA"then
        Celda = Celda & ofv.GenCelda("","cc","","1%","","","",BsecD & ars(4,inx),"si")
      else
        Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
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

           for i=1 to bsecix
               Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
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
                if stip <> "G" then Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
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
         pStyle = pStyle & "' STYLE={font-size:" & pfsize & "pt;}>"
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
         Celda = Celda & ofv.GenCelda("","aaa valign=TOP","","1%","","","",pVin,"si")
      end if

      sTextoV =  ars(8,inx) 

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
         ' response.write strsql
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


                      Strsql = "Select * From csGenObjetos where codobjeto = " & codobjetoL
                      
                      set rst2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rst2.eof then
                         wNomL = rst2(1)
                      else
                         wNomL = "S/D"
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

         sTextoV = replace(sTextoV,chr(13)," <br>")

      end if

   
     end if

 
      IF wasotxt = "N" THEN 

         if len(sTextoV) > 0 then

             sTextoV = replace(sTextoV,chr(13)," <br>")

         end if

         if  sImagen <> "NO" then sTextoV = sTextoV & replace(descri,"NNNNNN",sImagen)

      end if



      Celda   = Celda & ofv.GenCelda("",bspan,Ptalig & " valign=TOP style='text-align:justify;' ","","",pbgcol,"",pStyle & PsangS & sTextoV & pStyleF,"si")

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

        if  ars(12,inx) <> "" then sreduc = " width=" & ars(12,inx) & "% "
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

      


          Strsql = "Select codobjetol,tipobjetol From objlnk Where ltipo = 'KL' and objid = 0" & ars(9,inx)
          set rsCompoM = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsCompoM.eof then

 
              Strsql = "Select nombre,TIPOCOMP From csacomple Where codcomp = " & rsCompoM(0)
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
             Strsql = "SELECT   * FROM COMPIMAG where codcomp = 0" & rsCompoM(0) & " ORDER BY SECUENCIA, OBJID "
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
             sHref = GenKeyLnks("EditarImagen2.asp?ID=" & rsCompoM(0))
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

        sHTML = sHTML &  Fila

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
      sHTML = sHTML &  Fila


      if sVinoAnt > "F" and stip < "G" then 
         celda = ""
         for i=1 to 10
             Celda = Celda & ofv.GenCelda("","ut","","","","","","&nbsp;","si")
         next
         Fila  = ofv.GenRow("","","","","12","","",celda,"si")
         sHTML = sHTML &  Fila
      end if

    end if


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
  sHTML = sHTML &  Fila




' call ofv.cerrarconsulta(rs)


  celda = ""




  sHTML = sHTML & ofv.GenFTabla()
  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

  call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''

 else

  shtml = shtml & "<script language=javascript >"
  shtml = shtml & "    document.all.DHEX1.style.display='block';"
  shtml = shtml & "</script>"


 end if






  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("S",bDescri)

  if  dblinx = "S" then

       shtml = shtml & "<script language=javascript >"
       shtml = shtml & "document.all.DHEX2B.style.display='block';"
       shtml = shtml & "document.all.DHEX2C.style.height='170px';"
       shtml = shtml & "</script>"

  end if

  sHTML   = sHTML & SPLFUNC

  'Botonera

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();'  ", "S", "A", "N", "glosFrm", "")


  'cambios
   Listar = ""

  if origen <> "ED" then

   if bVersion = "SI" and SESSION("SHWMOD") <> "N" then

      if  marca <> "CA" then
          xcambios = "&sdep=C"
      else
          xcambios = ""
      end if

      xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
      sAccion = "coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
      sAccion = sAccion & "&sdocum=" & sdocum & "&sdocc=" & sdocc & xcambios
      Listar  = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

   end if
  end if




  Agregar = Celt2
  Buscar  = Celt1
  Celda   = Listar

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

'''''''''''''''''

  sHTML   = sHTML & SPLFUNC

  sHTML = sHTML & GTCOBS()


  sHTML = sHTML & "</body></html>"

end if

  response.write sHTML

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




%>