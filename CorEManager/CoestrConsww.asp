<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTESTRUCTURA.inc" -->
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
  dim jsHashPefiles
   dim jsHashsectores
  Adepi(1) = "A"
  Adepi(2) = "B"
  Adepi(3) = "C"
  Adepi(4) = "D"
  Adepi(5) = "E"
  Adepi(6) = "F"
  Adepi(7) = "I"
   set cn = ofv.conectar(ofv.strconn2)
  for ix = 0 to 7
      Adep(ix) = 0
  next


  sfidc  = "<font face=wingdings size=2 color=" & Session("CMFPNGNFT") & ">0</font>"
  sfida  = "<font face=wingdings size=2 color=" & Session("CMFPNGRFT") & ">1</font>"
  sfidk  = "<font face=wingdings size=2 color=#f7aa22 >t</font>"

  sHTML = ""

  jsHashPefiles="hashperfiles"

  jsHashsectores="hashsectores"
 shtml = shtml & ofv.FormHeader(session("FormN"))  & chr(13)
' shtml=replace(shtml,".opener","")

  shtml= shtml &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & chr(13)
  if  session("APLMDCLI") = "BGAL" then
  shtml=shtml & "<script type='text/javascript' src='hashtable.js'></script>"
'  sHTML = sHTML & "<script> var hashperfiles;" & jsHashPefiles & "</script>"
   shtml = shtml & "<script language='javascript' >"
  shtml = shtml  & armaperfiles(jsHashPefiles,cn,nCodigo)   &  armasectores(jsHashsectores,cn,nCodigo)
  else
   shtml = shtml & "<script language='javascript' >"
  end if
 shtml = shtml & " function Tcreawin(spmk) {   "
 'shtml = shtml & " alert('1');        "
 shtml = shtml & " document.all[spmk].scrollIntoView(true); "
 shtml = shtml & " document.all['DBpanel'].scrollIntoView(true); "


'''' shtml = shtml & " clIndice();         "
 shtml = shtml & " }    "
 shtml = shtml & "</script>"




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

            ' Celdat = "<td class=ut style='cursor:hand;' onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
            ' Filatt  = filatt & ofv.GenRow("","","","","","","",Celdat,"si")

             Celdat  = "<td class=utis >" & sfidkD & "</td>"
             Celdat  = Celdat & "<td width=95% class=uti style='cursor:hand;' onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
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


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' >"




  if ID <> "" then



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
' response.write strsql & "<br>"
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

     Celdat = "<td class=ut style='cursor:hand;'  onclick='document.all.ATIT" & rs(0) & ".click();' >" & stexto1 & "</td>"
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
      Celda =  ofv.GenCelda("","ttt","","90%","","","",pmk & "&nbsp;" & bDescri,"si")
   end if

   Celt2 =  ofv.GenCelda("","ut","","1%","","","",t2,"si")
   Celt1 =  ofv.GenCelda("","ut","","1%","","","",t1,"si")



  styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"





  Fila  = ofv.GenRow("","","","","25","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("","aaa' ","","100%  ","","0","","0","0","","",Fila,"si") & "<br>"



'''''''''''''



  sHTML = sHTML &   GENFMTESTRUCTURA(nCodigo,STIPO)






''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

  call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''

 else

  shtml = shtml & "&nbsp;"

  shtml = shtml & "<script language=javascript >"
  shtml = shtml & "    document.all.DHEX1.style.display='block';"
  shtml = shtml & "</script>"


 end if


  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"


  Celda  =  ofv.GenCelda("","ut","","","","","","&nbsp;","si")
  Filann = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML  = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%  ","","0","","0","0","","",Filann,"si")



  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("S",bDescri)

  if  dblinx = "S" then

       shtml = shtml & "<script language=javascript >"
       shtml = shtml & "  var et=document.getElementById('DHEX2B'); if (et  != null) { document.getElementById('DHEX2B').style.display='block'}; "   '& "document.get.DHEX2B.style.display='block';"
       shtml = shtml &  " var et=document.getElementById('DHEX2C'); if (et  != null) { document.getElementById('DHEX2C').style.height='170px'}; "  '  "document.all.DHEX2C.style.height='170px';"
       shtml = shtml & "</script>"

  end if

  sHTML   = sHTML & SPLFUNC

  'Botonera
  Celt1 = ""
 ''''''' Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();'  ", "S", "A", "N", "glosFrm", "")


  'cambios
   Listar = ""

  if origen <> "ED" then

   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

        if  marca <> "CA" then
          xcambios = "&sdep=C"
        else
          xcambios = ""
        end if

       ' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
        sAccion = "coestrCons.asp?Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
        sAccion = sAccion & "&sdocum=" & sdocum & "&sdocc=" & sdocc & xcambios
        Listar  =  genbotoncambios(sAccion) 'ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

      elseif SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)


      end if


   end if
  end if




  Agregar = Celt2
  Buscar  = Celt1
  Celda   = Listar

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)
   'sHTML   = sHTML & ofv.botoneraESSX("0",rs,"","","","","","","","no",Agregar,Buscar,celda)
'''''''''''''''''

  sHTML   = sHTML & SPLFUNC

  sHTML = sHTML & GTCOBS()


  sHTML = sHTML & "</body>"


if  session("APLMDCLI") = "BGAL" then
 shtml = shtml & "<script>change_span('',"& jsHashPefiles &",'spanpf');change_span('',"& jsHashsectores &",'spansec');</script></html>"
 else
  shtml = shtml &"</html>"
 end if

end if

  response.write sHTML






%>