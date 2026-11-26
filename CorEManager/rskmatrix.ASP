<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

  session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

  sMenu = request.querystring("Menu")
  sNombre = request.querystring("Nombre")
  sTema = request.form("codtema")
  if sTema = "" or sTema = "nsnc" then sTema = 0
  sSubtema = request.form("codprodserv")
  if sSubtema = "" or sSubtema = "nsnc" then sSubtema = 0
  sEsquema = request.form("codobjeto")
  if sEsquema = "" or sEsquema = "nsnc" then sEsquema = 0
  xComp = ""
  xComp = request.form("XComplem")



  XXXX = NEWEMPRE("NO")


  sEMPRE = request.querystring("EMPRE")
  IF SEMPRE = "" THEN
     SEMPRE = ATREMP(0,0)
  ELSE
     XXXX = NEWEMPRE(SEMPRE)
  END IF

  sUN = SEMPRE

  if sUN = "" then sUN = 1
  session("UN") = sUN


  'response.write  smenu
  set cn = ofv.conectar(ofv.strconn2)



  sNombren = sNombre

  sNombre = "Consulta por " & sNombre
  snombre = ucase(snombre)

  sHTML = ofv.FormHeader(session("FormN")) & chr(13)

  sImagenM = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"


  Fila    = ""
  Celda   = ""
  sHref   = ""

  IF  Session("EMPMUN") = "S" THEN

  Fila    = ""

  SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

  Fila    = ""
  Celda   = ""
  sHref   = ""

  sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
  Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
  Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
  Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

  Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

  Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





  shtml=shtml & "<script language=javascript>" & chr(13)
  shtml=shtml & "function shwDem(){" & chr(13)
  shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
  shtml=shtml & "}" & chr(13)
  shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
  sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
  sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
  sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
  sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
  TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:150px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


  TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
  TABLAX = TablaX & "<tr><td  "
  TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
  TABLAX = TablaX & "<tr><td colspan=3 "
  TABLAX = TablaX & " >&nbsp;</td></tr>"

  FOR IAEMP=1 TO TAEMP

   '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

     IF ATREMP(IAEMP,2) = 1 THEN

      IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
         TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
      ELSE
         TABLAX = TablaX & "<tr>"
      END IF

      TABLAX = TablaX & "<td width=5 "
      IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
         TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
      ELSE
         TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
      END IF

      sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
      TABLAX = TablaX & "<td width=5 "
      TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

      TABLAX = TablaX & "<td class=vt  >"

      sHref = "rskmatrix.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)

    '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
      sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
      TABLAX = TablaX & sLink
      TABLAX = TablaX & "</td>"
      TABLAX = TablaX & "</tr>"
     END IF

  NEXT

  TABLAX = TablaX & OFV.genFtabla()

  TABLA = Tabla & TABLAX

  TABLA = Tabla & "</DIV>"

  sHTML = sHTML & TABLA

  end if

  Fila = ""




  tt="&nbsp;"



  Fila       = ""
  formx      = ""
  Formulario = ""

 'Descripcion
  sInput = "<big>SELECCIONE LOS PARAMETROS PARA LA CONSULTA</big>"
  Celda  = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")
  Fila   = ofv.genrow("","","","","30","","",Celda,"si")

  Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
  Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
  Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

  if sNombre = "CONSULTA POR CATEGORIAS" then

      sInput  = ""
      sInput2  = ""
      sInput3  = ""
      sInput4  = ""
      sInput5  = ""
      sInput6  = ""
      sInput7  = ""

      xstdoc   = "0" 
      xstsel   = "0" 
      xstsec   = "0" 

''response.write "<br><br><br>---------------------------"
''for each x in request.form
''    response.write "-- " & x & " : " & request.form(x) & " --<br>"
''next
''response.write "---------------------------"


      xx1 = request.form("Ntipodoc") & ""
      xx2 = request.form("Ntcod") & ""
      xx3 = request.form("Nlista") & ""
      xx4 = request.form("textot") & ""
      xx5 = request.form("tsel") & ""

      ttcc = request.form("ttcc") & ""




      stipodoc    = request.form("Ntipodoc") & ""
      if stipodoc = "" then
      '   stipodoc = "99999" 
         xstdoc   = "0" 
      else 
         xstdoc   = "1" 
      end if 



      tsel     = request.form("tsel") & ""

      if tsel  = "" then
         tsel  = "C"
         tgrp  = "T"
         tcod  = "99999"
         ttip  = "TODOS"
         tnam  = "TEMA"
         xstsel   = "0" 
      else 
         xstsel   = "1" 

         if tsel  = "" then
            tgrp  = "T"
         end if

      end if

      frmsen = "N" 

      if tsel  = "T" then 
         tgrp  = "M"
         tcod  = "99999"
         ttip  = "TODOS"
         tnam  = "TODOS"

      end if

      sSectort = request.form("Nlista") & ""

      sSectort    = ofv.convertircar(SsectorT,"_"," ")
      if ssectort = "" then
      '   ssectort = "99999" 
         xstsec   = "0" 
      else 
         xstsec   = "1" 
      end if 

      textot = request.form("textot") & ""




    formul = request.querystring("frm")

    if len(formul) = 0 then

         tgrp  = request.form("tgrp") & ""
         tcod  = request.form("Ntcod") & ""
         ttip  = request.form("ttip") & ""
         tnam  = request.form("tnam") & ""


    elseif formul = 4 then

      if tsel  <> "T" then 
         tgrp   = request.form("ptgrp") & ""
         tcod   = request.form("ptcod") & ""
         ttip   = request.form("pttip") & ""
         tnam   = request.form("ptnam") & ""
         textot = request.form("ptextot") & ""
      end if  

    else

      if tsel  <> "T" then
         tgrp  = request.form("tgrp") & ""
         tcod  = request.form("Ntcod") & ""
         ttip  = request.form("ttip") & ""
         tnam  = request.form("tnam") & ""

      end if  


    end if

    if  tgrp <> "M" and formul=4 then frmsen = "S"

''''''''' introduccion '''''''''''''''''''''''''''''''''''''


      sInputxx1 = ofv.GenerarInput("Ntipodoc","99999","hidden","","1","dt","")
      sInputxx2 = ofv.GenerarInput("Ntcod","99999","hidden","","1","dt","")
      sInputxx3 = ofv.GenerarInput("Nlista","99999","hidden","","1","dt","")
      sInputxx5 = ofv.GenerarInput("tsel","C","hidden","","1","dt","")
      sInputxx6  = ofv.GenerarInput("tgrp","T","hidden","","1","dt","")
      sInputxx7  = ofv.GenerarInput("ttip","TODOS","hidden","","1","dt","")
      sInputxx8  = ofv.GenerarInput("tnam","TODOS","hidden","","1","dt","")

      sInputxx9b  = ofv.GenerarInput("ttcc","BASIC","hidden","","1","dt","")

      sInputxx9a  = ofv.GenerarInput("ttcc","AVANZADA","hidden","","1","dt","")


      sInputxb1 = ofv.GenerarInput("bx1","AVANZADA","BUTTON","","1","dt","document.all.Subconsultax1.submit();")
      sInputxb2 = ofv.GenerarInput("bx2","BASICA","BUTTON","","1","dt","document.all.Subconsultax2.submit();")

       sAccion = "rskmatrix.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formuln2 = ofv.GenForm("Subconsultax2","",sAccion,"","post","",sInputxx1 & sInputxx2 & sInputxx3 & sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9B,"si")


       sAccion = "rskmatrix.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formuln1 = ofv.GenForm("Subconsultax1","",sAccion,"","post","",sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9A,"si")


       sHTML = sHTML & Formuln2 & Formuln1


      celdan = ""
      celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb1,"si")
      celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb2,"si")
      Filan  = ofv.GenRow("","","","","","","",celdan,"si")
      Filan  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;display:none;' ","","20% ","","0","","0","0","","",Filan,"si")


      celda = ofv.GenCelda("","ut COLSPAN=2  ","RIGHT","","","","",Filan,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



      sglmar = "<font id='fnt01' size=4 color='#f7aa22' face='corflaticon' onclick='alert(""Seleccione los parametros para la busqueda...."");' >&#0125;</font>&nbsp;"

      sglpnt = "<font size=2 color='#00aa00' face='corflaticon' >4</font>&nbsp;"

      sinput = ""


         sintab = "<script language='javascript' > "
         sintab = sintab & " function opsel(dtfld) {document.all[dtfld + 'tab1'].style.display='none';document.all[dtfld + 'sl'].style.display='block';} "
         sintab = sintab & " function setsel(dtfld) {"
         sintab = sintab & "   document.all.Subconsultax[dtfld].value=document.all.Subconsultax[dtfld + 'sl'].value;"
         sintab = sintab & "   document.all[dtfld +'td1'].innerHTML=document.all.Subconsultax[dtfld + 'sl'].options[document.all.Subconsultax[dtfld + 'sl'].selectedIndex].text;"
         sintab = sintab & "   document.all[dtfld + 'tab1'].style.display='block';document.all[dtfld + 'sl'].style.display='none';"
         sintab = sintab & "   document.all.Subconsultax.submit();"
         sintab = sintab & "                    } "
         sintab = sintab & "</script>"

      sinput = sinput & sintab

      celda = ofv.GenCelda("","utg COLSPAN=2 style='text-align:justify;' ","","","","","",sinput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

 if ttcc <> "BASIC" THEN



''''''''' tipo de RIESGO (SITUACION) '''''''''''''''''''''''''''''''''''''

      sinputtt = "SITUACION"

      celda = ofv.GenCelda("","ttT ","","50%","","","",sinputtt,"si")


      Strsql = "SELEct * FROM RSKSITUACION  order by descri "

      scomK = ""
      scomD = "SELECCIONE SITUACION"

      sinputdtK = ""
      sinputdtd = "SELECCIONE SITUACION"
      sintab2   = ""

      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & rsr(0)
         scomD = scomD & chr(9) & rsr(1)

         sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


         if stipodoc  = cstr(rsr(0)) then
            sinputdtd = rsr(1)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(1) & "</option>"

      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)

   '      scomK = scomK & chr(9) & "99999"
   '      scomD = scomD & chr(9) & "* CUALQUIER TIPO DE OBJETO"

   '      sintab2 = sintab2 & "<option value='99999' "

   '      if stipodoc  = "99999" then
   '         sinputdtd = "* CUALQUIER TIPO DE OBJETO"
   '         sintab2    = sintab2 & " selected "
   '      end if          

   '      sintab2 = sintab2 & " >* CUALQUIER TIPO DE OBJETO</option>"


         sintab = "<table id='Ntipodoctab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Ntipodoctd1' width='100%' class='tdcm1' onclick=""opsel('Ntipodoc');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Ntipodoctd2' width='1'    class='tdcm2' onclick=""opsel('Ntipodoc');"" >"
         sintab = sintab & "<font class='btic'  >6<font></td>" 
         sintab = sintab & "</table>"


         sintab3 = "<select name='Ntipodocsl' size='5' onchange=""setsel('Ntipodoc');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Ntipodoc",stipodoc,"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK

      Celda = Celda & ofv.GenCelda("","ut  ","center","50%","","","",sInput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

 



'''''''''''' CATEGORIAS ''''''''''''''''''''''''''''''''''''''''''''''''

  if  xstdoc   = "1" then

      sinputtt = "CATEGORIA"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")

      Strsql = "SELEct * FROM RSKCATEGO WHERE SITUACION = 0" & stipodoc & " order by descri "

      scomK = ""
      scomD = "SELECCIONE CATEGORIA"

      sinputdtK = ""
      sinputdtd = "SELECCIONE CATEGORIA"
      sintab2   = ""

      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & rsr(0)
         scomD = scomD & chr(9) & rsr(2)

         sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


         if tcod  = cstr(rsr(0)) then
            tnam       = replace(rsr(2)," ","_")
            sinputdtd  = rsr(2)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(2) & "</option>"


      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)

         scomK = scomK & chr(9) & "99999"
         scomD = scomD & chr(9) & "* TODAS LAS CATEGORIAS"

         sintab2 = sintab2 & "<option value='99999' "

         if tcod  = "99999" then
            tnam       =  "TODAS_LAS_CATEGORIAS"
            sinputdtd  = "* TODAS LAS CATEGORIAS"
            sintab2    = sintab2 & " selected "
         end if          

         sintab2 = sintab2 & " >* TODAS LAS CATEGORIAS</option>"


         sintab = "<table id='Ntcodtab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Ntcodtd1' width='100%' class='tdcm1' onclick=""opsel('Ntcod');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Ntcodtd2' width='1'    class='tdcm2' onclick=""opsel('Ntcod');"" >"
         sintab = sintab & "<font class='btic' >6<font></td>" 
         sintab = sintab & "</table>"


         sintab3 = "<select name='Ntcodsl' size='5' onchange=""setsel('Ntcod');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Ntcod",tcod,"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK


      sInput1  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
      sInput3  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")

      Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

  end if



'''''''''' sector '''''''''''''''''''''''''''''''''''''''''''''''''


  if  xstdoc   = "1" and   xstsel   = "1" then 

' if  stipodoc = "99999" or stipodoc = "PCD" then

      Strsql = "SELEct * FROM RSKitem order by descri "

      scomK = ""
      scomD = "SELECCIONE ITEM"

      sinputdtK = ""
      sinputdtd = "SELECCIONE ITEM"
      sintab2   = ""


      scomK = scomK & chr(9) & "99999"
      scomD = scomD & chr(9) & "* PONDERACION GLOBAL"

         sintab2 = sintab2 & "<option value='99999' "

         if ofv.convertircar(SsectorT," ","_")  = "99999" then
            sinputdtd = "* PONDERACION GLOBAL"
            sintab2    = sintab2 & " selected "
         end if          

         sintab2 = sintab2 & " >* PONDERACION GLOBAL</option>"


      set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
      do until rsr.eof

         scomK = scomK & chr(9) & ofv.convertircar(rsr(0)," ","_")
         scomD = scomD & chr(9) & rsr(1)


         sintab2 = sintab2 & "<option value='" & ofv.convertircar(rsr(0)," ","_") & "' "


         if ofv.convertircar(SsectorT," ","_")  = ofv.convertircar(rsr(0)," ","_") then
            sinputdtd = rsr(1)
            sintab2    = sintab2 & " selected "
         end if

         sintab2 = sintab2 & " >" & rsr(1) & "</option>"


      rsr.movenext
      loop
      call ofv.cerrarconsulta(rsr)


         sintab = "<table id='Nlistatab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
         sintab = sintab & "<td id='Nlistatd1' width='100%' class='tdcm1' onclick=""opsel('Nlista');"" >" & sinputdtd & "</td>"
         sintab = sintab & "<td id='Nlistatd2' width='1'    class='tdcm2' onclick=""opsel('Nlista');"" >"
         sintab = sintab & "<font class='btic' >6<font></td>" 
         sintab = sintab & "</table>"


         sintab3 = "<select name='Nlistasl' size='5' onchange=""setsel('Nlista');"" "
         sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
         sintab3 = sintab3 & "</select>"


     sinputdtK = ofv.GenerarInput("Nlista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")

      sInput   = sintab & sintab3 & sinputdtK

      sinputtt = "ITEM DE PONDERACION"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")
      Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput,"si")
      Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

' else

'      SsectorT = "99999"

'      xstsec   = "1"

' end if

 else

      SsectorT = "99999"

 end if 

 END IF

'''''''''''' texto a buscar '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  if  xstdoc   = "1" and   xstsel   = "1" and   xstsec   = "1"  then 

      if textot = "" then
         textot = " " 
      end if 

      sinputtt = "CONFIRMA SELECCION DE BUSQUEDA?"

      celda = ofv.GenCelda("","ttT ","","","","","",sinputtt,"si")


 end if


       sAccion = "rskmatrix.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
       Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"","post","",Fila,"si")

       Fila    = ""

         if tgrp  = "" then
            tgrp  = "T"
         end if

      sInput0  = ofv.GenerarInput("tdoc",stipodoc,"hidden","","1","dt","")
      sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
      sInput5  = ofv.GenerarInput("tcod",tcod,"hidden","","1","dt","")
      sInput6  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
      sInput7  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")




    sInput3     = ofv.GenerarInput("buscar",textot,"HIDDEN","","50","dt","")



 

    if  textot <> "" then

        sInput     = "&nbsp;"
        Celda      = Celda & ofv.GenCelda("","UT","center","","","","",sInput0 & sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7,"si")
        Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

    end if



      formx    = Fila




 
'Fin preparacion
  end if


'''' Armado del Form de despacho '''''''''''''''''''''''

    if formx <> "" then


      Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
      sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
      formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")

    sAccion    = "rskmatrixdet1.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
    Formulario = Formulario & ofv.GenForm("Subconsulta","",sAccion,"","post","",formx,"si")

    end if

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''





'''''''''''''''''''''''''''''''''''''''''

  sHTML = sHTML & "<center><BR>"


      sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;' ","","80% ","","0","","5","0","","",Formulario,"si")



  sHTML = sHTML & "</center>"

  sHTML = sHTML & HTMgenlista

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  if session("SMPLMENU") = "S" then
     sAccion = "novedmenu.asp"
     Celda   = ofv.BotonVolver(sAccion)
  end if

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "<script language=JavaScript>"
  sHTML = sHTML & " function Init() "
  sHTML = sHTML & "   { ff(); setTimeout (" & chr(34) & "Init()" & chr(34) & ", 350); }"
  sHTML = sHTML & " function ff() "
  sHTML = sHTML & "  {   "
  sHTML = sHTML & "  if (document.all.fnt01.color=='#f7aa22') { document.all.fnt01.color='#aa0000'; } "
  sHTML = sHTML & "  else { if (document.all.fnt01.color=='#aa0000') { document.all.fnt01.color='#f75500'; } "
  sHTML = sHTML & "  else {if (document.all.fnt01.color=='#f75500') { document.all.fnt01.color='#f7aa22'; } } } "
  sHTML = sHTML & "   } "


  sHTML = sHTML & "</script>"


  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML


%>