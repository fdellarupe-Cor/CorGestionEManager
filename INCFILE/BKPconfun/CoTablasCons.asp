<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTTABLAS.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sdocum  = request.querystring("sdocum")
  sNIV1   = request.querystring("NIV1")
  sNIV2   = request.querystring("NIV2")
  sNIV3   = request.querystring("NIV3")
  sNIV4   = request.querystring("NIV4")
  sNIV5   = request.querystring("NIV5")

  INXMOD  = request.querystring("INXMOD")

  if  INXMOD = "" then INXMOD = "S"


  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)

 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------

'----------------------------------------------
'''''''''''''''''''''''''''''''''''''''''''''''
'' ACTIVACION PUNTERO DEL LINK 
''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SNIV1 = "9" AND SNIV2 <> "0" THEN 

    sdocum =  SNIV2

    IF SNIV3 <> "0" THEN

       SPLFUNC = "<script language='javascript' > creawin('BLK" & SNIV3 & "');</script>"

    ELSE

       SPLFUNC = ""

    END IF 
    

 ELSE

    SPLFUNC = ""

 END IF 

       SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"


'''''''''''''''''''''''''''''''''''''''''''''''''


  set cn = ofv.conectar(ofv.strconn2)

  strsqlx = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
      Nestado=rs(1) 
  end if
  call ofv.cerrarconsulta(rs)



''''''''''''''''''''''''''''''''''''''''''''''

  sft   = "<font face=tahoma,arial size=1 color=#336699>&nbsp;<b>"
  sfta  = "<font face=tahoma,arial size=1 color=#993333>&nbsp;<b>"
  sfidc = "<font face=corflaticon size=2 color=#336699>0</font>" 
  sfida = "<font face=corflaticon size=2 color=#993333>1</font>" 
  sfitc = "<font face=corflaticon size=2 color=#336699>0</font>" 
  sfita = "<font face=corflaticon size=2 color=#993333>1</font>" 
  sftc  = "</b></font>" 
  

  Formulario = ""
  sarch      = 0

    Strsql = "SELECT * "
    Strsql = Strsql & " from ObjTabDef Where codobjeto = " & nCodigo
    Strsql = Strsql & " order by tabsecuen "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     if sdocum = "" then sdocum = cstr(rs(0))

    ' tomo todas las tablas

      stexto1 = "Seleccionar&nbsp;tabla"
      Celdat  = stexto1

      sAccion = "javascript:void(0);"
      sOnClik = " onclick=""document.all.TIND2.style.display='block';document.all.TIND1.style.display='none';"" "


      stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
      stexto1 = stexto1 & "   >" & Celdat & "</a>"


      Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
   ''   Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

      stexto1 = "Tablas"

      Celdat  = "<td class=uts >" & sfidk0 & "</td>"
      Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
      Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")


    do until rs.eof

       if  sdocum = cstr(rs(0)) then
           sarch  = rs(0)
       end if 

       if  rs(3) <> "" then
           Celda  = ofv.GenCelda("","cc","","1%","","","",stexto1,"si")
           Celda  = Celda & ofv.GenCelda("","","","2","","","","&nbsp;","si")
           Formulario = Formulario & Celda
       end if


       if  sdocum = cstr(rs(0)) then
           sfidkD  = sfidkC
           stexto1 = replace(rs(4)," ","&nbsp;")
           stextoact = stexto1 

       else
           sfidkD  = sfidkA
           stexto1 = replace(rs(4)," ","&nbsp;")
           sAccion = "CoTablasCons.asp?INXMOD=N&Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo & "&sdocum=" & rs(0)
           sOnClik1 = ""
           sOnClik2 = " onclick=""document.all.ASC" & rs(0) & ".click();"" "


            Celdat  = stexto1

            stexto1 = "<A ID='ASC" & rs(0) & "' NAME='ASC" & rs(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
            stexto1 = stexto1 & "   >" & Celdat & "</a>"

            Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
            Celdat  = Celdat & "<td class=uti  >" & stexto1 & "</td>"
   '''   Filatt  = filatt & ofv.GenRow("","","","","","","",Celdat,"si")
            Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

       end if


    rs.movenext
    loop

  else

    sestr = "<font color=#336699><h5>Sin&nbsp;datos&nbsp;disponibles"
    sestr = sestr & "</h5></font>" 
    Formulario = ofv.GenCelda("","dt","","80%","","","",sestr,"si")

  end if

  Fila = ofv.GenRow("","","","","","","",Formulario,"si")
 ' sHTML = sHTML & ofv.GenTabla("","","","5%","","0","0","0","0","","",Fila,"si")


  vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_sdocum=" & sdocum

  nTabla = sarch


  Fila = ""

  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Celdat  = "<td class=uts colspan=2 style='border-bottom-style:inset;border-bottom-width:1;border-bottom-color:#dddddd;' >&nbsp;</td>"
  Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

  sfidkD  = sfidkc
  stexto1 = stextoact

  Celdat  = "<td class=uts >" & sfidkD & "</td>"
  Celdat  = Celdat & "<td class=uti  >" & stexto1 & "</td>"
  Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")



''''''''''''''''''''''''''''''''''''''''''''''

  Strsql = "SELECT * From Objtabdef Where objid = " & nTabla
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if  not rs.eof then
      nTitulo = rs(4)
      if nTitulo = "" then nTitulo = "&nbsp;"
      nPerfil = rs(3)
  end if


'Bloque de Defaults
ncolcant = 1
nBorde = false
nblkfc = "#993333"
nblkbc = "#ffffff"
nblkff = "ARIAL"
nblkfs = "10pt"

nctitfc = "#000077"
nctitbc = "#ffffff"
nctitff = "ARIAL"
nctitfs = "8pt"

ncconfc = "#336699"
ncconbc = "#ffffff"
ncconff = "ARIAL"
ncconfs = "8pt"

redim aColumnas(0)


nStilo = " Style={font-family:" & nblkff & ";font-size:" & nblkfs & ";color:#993333;background-color:#ffffff;font-weight:bold;} "

nfontx = "</font>"

npoint = "<a name=" & chr(34) & "inicio" & chr(34) & " ></a>"  
Celda = ofv.GenCelda("","ttt","","","","","","TABLA:&nbsp;" & nTitulo,"si")
Fila  = ofv.GenRow("","","","","","","",Celda,"si")
sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si") & chr(13)

sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;filter:dropShadow(color=#bbbbbb,OFFY=3,OFFX=3,positive=yes);" & styl & "' >" & chr(13)

sHTML = sHTML & "" & GENFMTTABLAS(nCodigo,STIPO)


  sHTML = sHTML & "</div>"

  sHTML = sHTML & "</div>"


''''''' selector del glosario

  GLhtml = VerGlosario(shtml)

  sHTML = GLhtml

  call ofv.cerrarconn(cn)

''''''''''''''''''''''''''''''


  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")


  if CantRegAMover = "A" then

     if INXMOD = "S"  then

        sHTML = sHTML & "<script language='javascript' > clIndice(); </script> "

     END IF



  else 

     if INXMOD = "N"  then

        sHTML = sHTML & "<script language='javascript' > clIndice(); </script> "

     END IF


  'Botonera

      Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")

  end if


  sHTML   = sHTML & SPLFUNC


  'Botonera

 ' Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
  Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")

  Agregar = Celt2
  Buscar  = Celt1
  Celda   = ""

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML



%>