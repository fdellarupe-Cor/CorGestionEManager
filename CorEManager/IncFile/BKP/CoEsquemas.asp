<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sSectorx = request.querystring("Sector")

  sCambios = request.querystring("sCambios")


  if  sCambios = "" then sCambios = "NO"

  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
  sNIV3   = request.querystring("NIV3")
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")

  nCantRegAMover = request.querystring("CantRegAMover")

  if nCantRegAMover = "x" then
     xsproced = request.form("selcom")
  elseif nCantRegAMover = "A" then
     xsproced = request.querystring("selcom")
  end if

'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK 
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "8" AND SNIV2 <> "0" THEN 

    xsproced =  SNIV2

    nCantRegAMover = "A"
    

 ELSE

    SPLFUNC = ""

 END IF 

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"


'''''''''''''''''''''''''''''''''''''''''''''''''



  sMenux = request.querystring("Menu")
  if sMenux = "PORSECTOR" then
     sSelecx = " and (descri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' "
     sSelecx = sSelecx & " or descri1 = '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "') "
  else
     sSelecx = ""
  end if
  if nCantRegAMover = "x" then
   Strsql = "SELECT * from csesquemasel "
   Strsql = Strsql & "Where codobjeto=" & nCodigo
   Strsql= Strsql & " and objid = " & xsproced & sSelecx
  else
  if nCantRegAMover = "A" then
   Strsql = "SELECT * from csesquemasel "
   Strsql = Strsql & "Where codobjeto=" & nCodigo
   Strsql= Strsql & " and objid = " & xsproced & sSelecx
  else
   Strsql = "SELECT * from csesquemasel "
   Strsql = Strsql & "Where codobjeto=" & nCodigo & sSelecx
   CantRegAMover="0" & request.querystring("CantRegAMover")
   CantRegAMover=cdbl(CantRegAMover)
   nCantRegAMover = CantRegAMover
  end if
  end if
 


  CantRegAMover2 = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  if ofv.explorador = "MSIE" then
    CantRegAMostrar2 = 7
  else
    CantRegAMostrar2 = 5
  end if

  sMenu = Session("menu")
  bModificar = session("Modificar")
  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_selcom=" & xsproced
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)



  sHTML = ofv.FormHeader(session("FormN"))

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  Fila = ""

''''''''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)

  strsqlx = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1) 
  end if
  call ofv.cerrarconsulta(rs)


 '' sHTML = sHTML & ofv.GenTabla("panel","","","100% border=0 ","",0,0,0,0,"","","","no")

  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     nID       = rs(0)
     nref      = rs(2)
     bDescri   = rs(1)
     if len(nref) = 0 then nref = "NO"
''''''''''''''''''
      stexto1 = "Seleccionar&nbsp;Esquema"
      Celdat  = stexto1

      sAccion = "javascript:void(0);"
      sOnClik = " onclick=""document.all.TIND2.style.display='block';document.all.TIND1.style.display='none';"" "


      stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   ''   Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

      stexto1 = "Esquemas"

      Celdat  = "<td class=uts >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   ''   Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")



''''''''''''''''''''''
     Strsql = "SELECT * from csesquemasel "
     Strsql = Strsql & "Where codobjeto=" & nCodigo & sSelecx

     set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     do while not rsx.eof

       if clng(rsx(0))   = clng(rs(0).value) then
          sfidkD  = sfidkC
          stexto1 = rsx(1)
          stextoact = stexto1
       else
          sfidkD  = sfidkA
          stexto1 = rsx(1)
          sAccion  = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
          sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & rsx(0)
          sOnClik1 = ""
          sOnClik2 = " onclick=""document.all.ASC" & rsx(0) & ".click();"" "
 
       Celdat  = stexto1



      stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"

      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
      Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"
      Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")
      end if

      rsx.movenext
   loop  

''''''''''''''''''''''''''''''''

  else
      nID = "no"

  end if

  Fila = ""

  stitu = "<big>" & bDescri & "</big>"

  if scambios = "SI" then
     stitu = stitu & "&nbsp;(Cambios)"
  end if

  Celda = ofv.GenCelda("","ttt","","60%","","","",stitu,"si")
  Fila  = fila & ofv.GenRow("","","","","","","",Celda,"si")

' Referencia 
  if nref <> "NO" then
     nref = replace(nref,chr(13),"<br>")
  ELSE
     nref = "&nbsp;"
  end if
  Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100%","","","",nref,"si")
  Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")
'''''''''''''''''''''''''''''''''''''''''''''''''

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  sHTML = sHTML &  "<div id='mt2' style='display:none;height:350px;overflow:auto;" & styl & "' >"

  Celdat  = "<td class=uts colspan=2 style='border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</td>"
  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

  sfidkD  = sfidkc
  stexto1 = stextoact

  Celdat  = "<td class=uts >" & sfidkD & "</td>"
  Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"
  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")


''' response.write vuelta & "<br>" & chr(13)
     
    if nID <> "no" then

  Strsql = "SELECT OBJESQUEMA2.*, OBJCUENTA.descri, OBJCUENTA.numero, OBJCUENTA.objid "
  Strsql = Strsql & "FROM OBJESQUEMA2, OBJESQUEMA, OBJCUENTA where "
  Strsql = Strsql & "OBJESQUEMA2.codesquema = OBJESQUEMA.objid and "
  Strsql = Strsql & "OBJESQUEMA2.codcuenta = OBJCUENTA.objid "
  Strsql = Strsql & "and Objesquema2.codesquema=" & nID
  if scambios = "NO" then
     strsql = strsql & " and (OBJESQUEMA.modi = 0 or (OBJESQUEMA.modi <> 0 and OBJESQUEMA.tipmodi <> 'E')) "
     strsql = strsql & " and (OBJESQUEMA2.modi = 0 or (OBJESQUEMA2.modi <> 0 and OBJESQUEMA2.tipmodi <> 'E')) "
  else
     strsql = strsql & " and OBJESQUEMA.modi <> 0  "
     strsql = strsql & " and OBJESQUEMA2.modi <> 0  "
  end if


  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  nEsquema = ID
  sHTML = sHTML & ofv.GenTabla("","aaa ","","100%","",0,0,0,0,"","","","no")
  'Encabezados Cuenta Contable y Sentido

      Celda  = ""
      Celda  = Celda & ofv.GenCelda("","tt colspan=3","","1%","","","","Sentido","si")
      Celda  = Celda & ofv.GenCelda("","tt colspan=2","","50%","","","","Cuenta&nbsp;Contable","si")
      Celda  = Celda & ofv.GenCelda("","tt","","50%","","","","Referencia","si")
   '   sHTML  = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

  posicion = 0
  do until rs2.eof
     posicion = posicion + 1
     nCuenta  = rs2(0)
     nref     = rs2(6)
     sPoint   = "T" & rs2(0)


      xfclas    = "dt"
      sImage    = "&nbsp;"
      if bVersion = "SI" and rs2(7) and SESSION("SHWMOD") <> "N" then
      if rs2(8) = "M" then
         xftitl = "Modificado."
         xfclas = "btbs"
         sglmod = "<div class='" & xfclas & "' ><font class=btic >4</font></div>"
         sImage = ofv.GenLink("",xfclas,"","","","",sglmod,"si")
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

      Celda  = ofv.GenCelda("","uts title='" & xftitl & "' ","","1%","","","","<b>" & sImage & "</b>","si")


      ' estado observacion
      
      Celdaobs  = ofv.GenCelda("","ut","","1%","","","","&nbsp;","si")

  

        
      if Nestado=5  then 

      	   '    sAccion = "genObs.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	   '    sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nCuenta  & "&ID2=" & nID
      
      	   '    xsffun  = " onmouseover=""this.className='" & xclas & "r';"" onmouseout=""this.className='" & xclas & "';"" "
      	   '    sHref   = sAccion
      	   '    sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	   '    Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Observacion de " & BtipP & bpespec & "' ","","1%","1","","",sInput,"si")
           '    Celdaobs  = ofv.GenLink("",xclas & " title='Obs de " & BtipP & bpespec & "' ",sHref,"","","",Celdaobs,"si")


              sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & nCuenta  & "&ID2=" & nID
      
      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")


      end if


      Celda = Celda & Celdaobs



  'Texto Sentido
      sInput = rs2(5)
      if  sInput = "DEBE" THEN
          sInput = "DB.&nbsp;"
      ELSE
          sInput = "&nbsp;&nbsp;&nbsp;HB.&nbsp;"
      END IF
      Celda  = Celda & ofv.GenCelda("","Ut valign=top ","","1%","","","",sInput,"si")

 

  'Texto Cuentas Contables
      sInput = replace(rs2(16)," ","&nbsp;")
      Celda  = Celda & ofv.GenCelda(sPoint,"uts valign=top ","","15%","","","","<b>" & sInput & "</b>","si")


      sInput = replace(rs2(15)," ","&nbsp;")
      Celda  = Celda & ofv.GenCelda("","Ut valign=top  ","","35%","","","",sInput,"si")




'''''''''' arma tabla de indices '''''''''''''''''''''''''
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
      sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

      Celda = ofv.GenCelda("","dt colspan=5 ","","20","","","","&nbsp;","si")
      sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

      rs2.movenext

    loop


    sHTML = sHTML & ofv.GenFTabla()
 
  call ofv.cerrarconsulta(rs2)
    end if


  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"

''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

 '' call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botonera

  Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")




  'cambios
   Listar = ""

  if origen <> "ED" then

   if bVersion = "SI" and SESSION("SHWMOD") <> "N" then

      if  sCambios = "NO" then
          xcambios = "SI"
      else
          xcambios = "NO"
      end if

      xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
      sAccion = "CoEsquemas.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
      sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&selcom=" & nid & "&scambios=" & xcambios
      Listar  = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")

   end if
  end if




  Agregar = Celt2
  Buscar  = Celt1
  Celda   = Listar

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>