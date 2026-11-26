<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar   = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir  = ofv.AEliminar
  CantRegMostrarx = 10
  CantRegMoverx =  request.querystring("CantRegMover")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  stexto=trim(request.form("texto") &"")
  acc=request.querystring("acc")
  XXXX = NEWEMPRE("NO")

  sEMPRE = request.querystring("EMPRE")
  IF SEMPRE = "" THEN
     SEMPRE = ATREMP(0,0)
  ELSE
     recarga = NEWEMPRE(SEMPRE)
  END IF


  sUN = SEMPRE

  if sUN = "" then sUN = 1
  session("UN") = sUN


  sTipo = request.querystring("Tipo")
  sNombre = request.querystring("Nombre")

  ' response.write stipo & " 22222"
  if sTipo <> "" then
     session("stipo") = stipo
     session("nombre") = snombre

  else
     stipo = session("stipo")
     snombre = session("nombre")
  end if

  if CantRegMoverx =""  then
   CantRegMoverx=session("CantRegMoverx")
  else

   session("CantRegMoverx") =CantRegMoverx
  end if


  if acc ="R" then session("sbuscar")=""


  if stexto <> "" then
    session("sbuscar")=stexto
    else
    stexto=session("sbuscar")
  end if

  stexto=replace(stexto & "" ,"'","''")
  sImagenI   = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI2  = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
  sImagenI3  = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM  = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx   = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

  FUNCION = " onload="  & Chr(34)


  FUNCION = FUNCION & "rszdx4();"  & Chr(34)
 ''  FUNCION = FUNCION & " onmousedown = " & Chr(34) & "Abre();" & Chr(34) & " "
  FUNCION = FUNCION & " onresize = " & Chr(34) & "rszdx4();" & Chr(34) & " "

''  sHTML = ofv.MENUheader("",funcion)

  sHTML = ofv.FormHeader(session("FormN"))

 shtml=shtml & "<script language=javascript>" & chr(13)

 shtml = shtml & "function AbrirVentana(datos) { " & Chr(13)
 shtml = shtml & " var h=570;w=790 ;"
 shtml = shtml &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;}"
 shtml = shtml & " window.open(datos"
 shtml = shtml & "," & Chr(34) & "edi" & Chr(34) & "," '& Chr(34)
 'shtml = shtml & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
 shtml = shtml & "'top=0,left=0,menubar=no,toolbar=no,height=' + h +',width=' + w + ',resizable=yes,status=yes'" '& Chr(34)

 shtml = shtml & " ); " & Chr(13)
 shtml = shtml & " } " & Chr(13) & Chr(13)

 shtml = shtml & "function clIndice() { " & chr(13)
 shtml = shtml & "  if  (document.getElementById('DHEX1').style.display=='none')   " & chr(13)
 shtml = shtml & "      { document.getElementById('DHEX1').style.display='';} " & chr(13)
 shtml = shtml & "  else                                             " & chr(13)
 shtml = shtml & "      { document.getElementById('DHEX1').style.display='none';}  " & chr(13)
 shtml = shtml & " } " & chr(13)

 shtml = shtml & "function DCEXRCLC() {"
 shtml = shtml & " var xrclc0 = document.body.clientWidth;"
 shtml = shtml & " var xrclc1 = document.body.clientHeight;"
 shtml = shtml & " var xrclc2 = document.getElementById('TCEX1').clientHeight;"
 shtml = shtml & " var xrclc3 = document.getElementById('TCEX1').offsetTop;"
 shtml = shtml & " document.getElementById('mt2').style.height = xrclc1 - (xrclc3 + xrclc2) - 25;"

'' shtml = shtml & " document.all.mt2.style.height = xrclc1 - 5;"
'' shtml = shtml & " document.all.mt2.style.width  = xrclc0 - 10;"
 shtml = shtml & " document.getElementById('mt2').style.width  = '70%';"
 shtml = shtml & "}"

           shtml=shtml & "</script>" & chr(13)

 ''' sHTML = sHTML & ofv.CorFormHeader

  Fila = ""
  set cn = ofv.conectar(ofv.strconn2)



  Fila    = ""
  Celda   = ""
  sHref   = ""

  IF  Session("EMPMUN") = "S" THEN

  Fila    = ""

  SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='mostrarmenu();' >EMPRESA:&nbsp;&nbsp;</a>"

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
  shtml=shtml & "function mostrarmenu(){" & chr(13)
  shtml=shtml & "document.getElementById('zx1').style.display='';" & chr(13)
  shtml=shtml & "}" & chr(13)
  shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
  sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
  sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
  sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.getElementById('zx1').style.display='none';" & chr(34)
  sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
  TABLA = Tabla & "<DIV id='zx1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:150px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


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

      sHref = "CompTree.asp?EMPRE=" & ATREMP(IAEMP,0)
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

 shtml = shtml & "<center>"

  sglos2  = "<b><font face=corflaticon size=3 color=#ffffff>&</font></b>"
  sglos1  = "<b><font face=corflaticon size=3 color=#ffffff>q</font></b>"
  sglos3  = "<b><font face=corflaticon size=3 color=#ffffff>n</font></b>"

  sfidkA  = "<font face=corflaticon size=2 color=#aa0000 >4</font>"
  sfidkC  = "<font face=corflaticon size=2 color=#00aa00 >5</font>"

  sfidk0  = "<font face=corflaticon size=2 color=#aa0000 >5</font>"
  sfidk1  = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
  sfidk2  = "<font face=corflaticon size=2 color=#f7aa22 >4</font>"
  sfidk3  = "<font face=corflaticon size=2 color=#00aa00 >4</font>"

  sglcmp = "<div class=btba style='width:1px;clear:right;' ><font class=btic >&#0157;</font></div>"
  sglcnt = "<div class=btbs style='width:1px;clear:right;' ><font class=btic >G</font></div>"
  sglprf = "<div class=btbn style='width:1px;clear:right;' ><font class=btic >&#0128;</font></div>"
  sgllnk = "<div class=btbn style='width:1px;clear:right;' ><font class=btic >&#0126;</font></div>"






'----------------------------------------------
 Celda =  ofv.GenCelda("","vt","center","","","","","<br><big><big>Cargando Lista...</big><br><br>un momento por favor</big><br>","si")
 Fila  = ofv.GenRow("","","","","","","",celda,"si")


 shtml = shtml & "<div id='mt1' style='display:block;' ><br><br><br><br>"
 shtml = shtml & ofv.GenTabla("","","","60% style='border-style:ridge;border-width:2;border-color:#f7aa22;' ","","0","","0","0","#ffffff","",Fila,"si")
 shtml = shtml & "</div></center>"

  styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa; SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"

  Fila = ""
  Tabla = ""

  Filat = ""

  Strsql = "Select * From Tipocomp ORDER BY DESCRI"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      if  sTipo = "" then  sTipo = rs(0)
  end if
  do until rs.eof
    sHref = "CompTree.asp?acc=R&CantRegMover=0&Tipo=" & rs(0) & "&Nombre=" & ofv.convertircar(rs(1)," ","_")
    sImage = ofv.GenImage("icono","","","","Iconos\" & rs(2),"0","")
    sLink = ofv.GenLink("","nt",sHref,"","","",sImage,"si")
    Celda = ofv.GenCelda("","","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","ut",sHref,"","","",rs(1),"si")
    Celda = Celda & ofv.GenCelda("","ut","","95%","","","",sLink,"si")
    Filat = Filat & ofv.GenRow("","","","","25","","",Celda,"si")
    rs.movenext
  loop
  call ofv.cerrarconsulta(rs)

  shtml=shtml & "<table width=98% align=center><tr><td valign=top width=20% >"
  'sHTML = sHTML & "<DIV ID='DHEX1'  NAME='DHEX1'  style='HEIGHT=350px;background-color:#ffffff;border-style:outset;border-width:1;' >"

    'sHTML = sHTML & "<DIV ID='DHEX1C'  NAME='DHEX1C'  style='HEIGHT=350px;overflow:auto;border-style:solid;border-width:1;border-color:#dddddd;filter:progid:DXImageTransform.Microsoft.Gradient(startColorStr=" & Session("APLBGD") & ", endColorStr=#cccccc, gradientType=0);' >"
   ' sHTML = sHTML & "<CENTER>"
    sHTML = sHTML & ofv.GenTabla("TIND1","aaa style='display:block;filter:shadow(color=#dddddd,direction=125);' ","","90%","","0","0","2","0","","",Filat,"si")
    sHTML = sHTML & ofv.GenTabla("TIND2","aaa style='display:none;filter:shadow(color=#dddddd,direction=125);' ","","90%","","0","0","2","0","","",Filatt,"si")
   ' sHTML = sHTML & "</CENTER>"
    sHTML = sHTML & "</div>"
  sHTML = sHTML & "</div></td><td>"


  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

  Strsql = "Select * From Tipocomp Where tipo='" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      simage   = rs(2)
      sPerfil  = rs(4)
      sNombreN = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

    sImagen = ofv.GenImage("","","","","Iconos\" & simage,"0","")
    sLink = sImagen
    Celda = ofv.GenCelda("","ttt","","2%","","","",sLink,"si")
    sForm="ABSQ"
     sInput = "<input type='text' class='dt' name='texto' id='texto' value="& chr(34) & stexto & chr(34) & " size='80' style='border-style:solid;border-width:1;border-color:#666666;' >"
     sInput = sInput & ofv.GenerarInput("", "Buscar", "submit", "", "5", "bietx", "")
    sHref = "CompTree.asp?acc=R&CantRegMover=0"
    NFORM=ofv.GenForm(sForm,"",sHref,"","post","","Busqueda por Texto : " & sInput,"si")
    Celda = Celda & ofv.GenCelda("","ttt","","90%","","","","<u>LISTA DE " & sNombreN & "</u><br><br>" & NFORM,"si")
    Fila = ofv.GenRow("","","","","40","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa style=max-width:100%!important onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

    Fila = ""
 ' sHTML = sHTML &  "<center>"
  sHTML = sHTML &  "<div id='mt2'  style='display:none;border-style:qsolid;border-width:1;border-color:#dddddd;height:350px;overflow:auto;" & styl & "' >"


  Strsql = "Select * From Complementos Where descri like '%" & stexto &"%' and tipocomp='" & sTipo & "' and compun = " & session("un") & " order by descri "
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  xTrColor = "#FFFFFF"
  if Not rs.EOF then rs.move(CantRegMoverx)
   posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

 ' do until rs.eof
    'Eliminar
    if bSuprimir then
      sHref = "asp/ElComp.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tipo=" & sTipo & "&Eliminar=NSNC&Volver=CompTree.asp?xdt=dt"
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","vti","","2%","","","",posi & ".&nbsp;","si")
    end if

    sHref = "EditarComp.asp?Codcomp=" & rs(0) & "&Tipocomp=" & rs(2) & "&Nombre="
    sHref = sHref & sNombre & "&Complemento=" & ofv.convertircar(rs(3)," ","_") & "&Perfil="
    sHref = sHref & sPerfil

    sHref2 = "javascript:void(0);"
    sLink = ofv.GenLink("","ut onclick=""AbrirVentana('" & sHref & "');return false;"" ",sHref2,"","","",rs(3),"si")


    Celda = Celda & ofv.GenCelda("","ut","","95%","","","",sLink,"si")

    Fila = Fila & ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","","","",Celda,"si")

    if xTrColor = "#FFFFFF" then
       xTrColor = "#F0F0F0"
    ELSE
       xTrColor = "#FFFFFF"
    END IF

    posi=posi+1
    rs.movenext
  loop

  sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","0","0","0","","",Fila,"si")

  sHTML = sHTML & "</div>"

 ' sHTML = sHTML & "</center>"

  sHTML = sHTML & "</div>"


  shtml=shtml & "</td></tr></table>"



  'Botonera

  Celt1 = ""'ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")



  Agregar = ""
  if bAgregar then
    sAccion = "AgregarCompl.asp?Tipo=" & sTipo & "&Nombre=" & sNombre & "&Volver=" & vuelta
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar  = Celt1
  Celda   = ""



  sAccion = "ListadoPorComponente.asp?tipo=" & stipo & "&nombre=" & snombreN
  Celda = ofv.BotonConsultar(sAccion)

  'ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


  shtml = shtml & "<script language=javascript >"

  shtml = shtml & "document.getElementById('mt1').style.display='none';"
  shtml = shtml & "document.getElementById('mt2').style.display='';"
  shtml = shtml & " DCEXRCLC();"
  if  SESSION("FULLGUI") = "S" THEN
'      shtml = shtml & "   document.all.TBH1VML.textpath.string=" & bDescri & ";" & chr(13)
  ELSE
'      shtml = shtml & "   document.all.TBH1DV.innerHTML= '" & bDescri & "';" & chr(13)
  END IF
  shtml = shtml & "</script>"


'''''''''''''''''''''''''''

sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
  sHTML = sHTML &  "</body></html>"
end if

response.write sHTML
 call ofv.cerrarconsulta(rs)

call ofv.cerrarconn(cn)

%>