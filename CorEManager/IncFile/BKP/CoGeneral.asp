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
  sCodigo = nCodigo

  set cn = ofv.conectar(ofv.strconn2)

redim tobj(0)
redim tobjd(0)
oi = 0
  strsql = "SELECT * from tipobjeto"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  do until rs.eof   
     redim preserve tobj(oi)
     redim preserve tobjd(oi)
     tobj(oi) = rs(0)
     tobjd(oi) = rs(1)     
     oi = oi + 1
  rs.movenext
  loop
  end if

  call ofv.cerrarconsulta(rs)

'xxxxxxxxxxxx
redim preserve sPath(0)
i = 1
do

  strsql = "select * from csbarr where codojasoc = " & sCodigo & " and tipobjasoc = '" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    redim preserve sPath(i)
    sCodigo = rs(1)
    sTipo = rs(2)
    sPath(i) = ""
    if not isnull(isnull(rs(0))) then
       for oi=0 to ubound(tobj)
           if rs(2) = tobj(oi) then  
              sPath(i) = tobjd(oi) & ": "
              exit for
           end if
       next  
      sPath(i) = sPath(i) & rs(0)
    else
      i = i + 1
      exit do
    end if
    i = i + 1
  else
    exit do
  end if
  call ofv.cerrarconsulta(rs)
loop

    strsql = "SELECT PRODSERV.descri, TEMA.descri FROM TEMA, RELAOBJ, "
    strsql = strsql & "PRODSERV where RELAOBJ.codobjmaster = PRODSERV.codprodserv and "
    strsql = strsql & " RELAOBJ.tipobjmaster = PRODSERV.marcaps and "
    strsql = strsql & "TEMA.codtema = PRODSERV.codtema and codprodserv = " & sCodigo


set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
if not rs.eof then
  stSubtema = rs(0)
  stTema = rs(1)
  call ofv.cerrarconsulta(rs)
end if

for h = UBound(sPath) to 1 step - 1
  if trim(spath(h)) <> "" then
    if h > 1 then
      sRoot = sRoot & ofv.convertircar(sPath(h)," ","&nbsp;") & "\"
    else
      sRoot = sRoot & ofv.convertircar(sPath(h)," ","&nbsp;")
    end if
  end if
next
if right(sRoot,1) = "\" then sRoot = left(sRoot,len(sRoot)-1)

'xxxxxxxxxx


  Strsql="Select * From Objetos Where codobjeto=" & nCodigo
  set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof Then
    sModifica = session("Modificar")
    vuelta= ofv.vueltaasp2("../")

    wdescri = rs(8)  

    sHTML = ofv.FormHeader(session("FormN"))

'''''''''''''''''
      sImagenEM  = "<FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT>"
      sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"

      shtml   = shtml & "<script language=javascript>" & chr(13)
      shtml   = shtml & "function shwOrig(xusr,XUSRN){" & chr(13)
      shtml   = shtml & "document.all.cx1.style.display='block';" & chr(13)
      shtml   = shtml & "}" & chr(13)

            shtml=shtml & "function chequear(){" & chr(13)
            shtml=shtml & "alert(window.top.CorMtopm.name + ' '  + window.top.CorMtopm.top.name);" & chr(13)
            shtml=shtml & "}" & chr(13)

      shtml   = shtml & "</script>" & chr(13)
'----------------------------------------------

      sinput2 = "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=corflaticon size=1   >r</font>"

  sHTML = sHTML & "<DIV id='cx1' class='DINDM' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;"
  sHTML = sHTML & "position:absolute;left:100;top:160;width:600px;height:250px;overflow:auto;'>"

'''''''''''''''''''''

  sHTML = sHTML & OFV.gentabla("", "TWIND", "", "100%", "", "0", "", "-1", "-2", "", "", "", "no")
  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td CLASS=mbt width=99%  >&nbsp;&nbsp;ORIGEN&nbsp;DE&nbsp;LA&nbsp;DOCUMENTACION</td>"
  sHTML = sHTML & "<td CLASS=mbt width=1% ONCLICK=""document.all.cx1.style.display='none';"" title='Ocultar.' >"
  sHTML = sHTML & sinput2
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"
  sHTML = sHTML & OFV.genFtabla()
  sHTML = sHTML & "<center>"
  sHTML = sHTML & OFV.gentabla("", "aaa ", "", "98%", "", "0", "", "0", "0", "", "", "", "no")
  sHTML = sHTML & "<tr height=22 ><td width=40% "
  sHTML = sHTML & "class=ttt  >Origen</td>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=ttt  >Instrumento</td>"
  sHTML = sHTML & "<td width=10% "
  sHTML = sHTML & "class=ttt  >L.Instrumento</td>"
  sHTML = sHTML & "<td width=10% "
  sHTML = sHTML & "class=ttt  >L.Analisis</td></tr>"
 
 '' sHTML = sHTML & "<tr><td  "
 '' sHTML = sHTML & "class=UT colspan=4  >&nbsp;</td></tr>"


  Strsql="SELECT OBJPRTY.*, Estados.estado as estdes FROM OBJPRTY, Estados "
  Strsql=Strsql & "where OBJPRTY.estado = Estados.objid and codobjeto=" & nCodigo

  set rsx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx.eof Then

     nestado = rsx(9)
     sTipo   = rsx(2)

    svienemain = "NO"
    sOrigen1 = "&nbsp;"
    if rsx(3) <> "" then
      if clng("0" & rsx(3)) > 0 then
        svienemain = "SI"
        Strsql = "Select descri From Origen Where objid = " & rsx(3)
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen1 = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
      end if
    end if


    sInstrumento1 = "&nbsp;"
    if rsx(4) <> "" then
        Strsql = "Select descri From Instrumen Where objid = " & rsx(4)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento1 = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
    end if

    if rsx(5) <> "" then
       sInstiden1 =  rsx(5)
    else
       sInstiden1 = "&nbsp;"
    end if 

    if rsx(18) <> "" then
       wUrl = rsx(18) & ""

       if not isnull(wUrl) then
          slinkD = instr(wUrl,"://")
          if slinkD > 0 then
             slinkD = wUrl
          else
             slinkD = "HTTP://" & wUrl
          end if
       else
          slinkD = "javascript:void(0);"
       end if

       sInstlink1 = slinkD
            sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href='" & sInstlink1 & "' target=_blank "
	    sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
            sInstlink1 = sinput2
    else
       sInstlink1 = "---"
    end if 

    if rsx(19) <> "" then

       wUrl = rsx(19) & ""

       if not isnull(wUrl) then
          slinkD = instr(wUrl,"://")
          if slinkD > 0 then
             slinkD = wUrl
          else
             slinkD = "HTTP://" & wUrl
          end if
       else
          slinkD = "javascript:void(0);"
       end if

       sInstlink2 = slinkD


            sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href='" & sInstlink2 & "' target=_blank "
	    sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
            sInstlink2 = sinput2
    else
       sInstlink2 = "---"
    end if 

    xTrColor = "#FFFFFF" 

    sHTML = sHTML & "<tr height=22 bgcolor='" & xTrColor & "' >"

    sHTML = sHTML & "<td  class=ut>"
      'Texto  origen


    SArma = "" & "&nbsp;" & sOrigen1 & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "<td  class=ut>"
      'Texto  instrumento
 
    SArma = "" & sInstrumento1 & "&nbsp;" & sInstiden1 & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"



    sHTML = sHTML & "<td  class=ut align=center >"
      'link consulta

    SArma = "" & sInstlink1 & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "<td  class=ut align=center >"
      'link consulta

    SArma = "" & sInstlink2 & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "</tr>"



  end if

  svienemas = "NO"


  Strsql = "Select * From Objprtyadv Where codobjeto=" & nCodigo & " Order By secuen"

  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx.eof then   svienemas = "SI"

  do until rsx.eof

     if  xTrColor = "#FFFFFF" then
         xTrColor = "#F9F9F9" 
     ELSE
         xTrColor = "#FFFFFF" 
     END IF

    sHTML = sHTML & "<tr height=22 bgcolor='" & xTrColor & "' >"

    sHTML = sHTML & "<td  class=ut>"
      'Texto  origen
      sOrigen = ""
      if cstr(rsx(3) & "") <> "" then
        Strsql = "Select descri From Origen Where objid = " & rsx(3)
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
      end if

    SArma = "" & "&nbsp;" & sOrigen & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "<td  class=ut>"
      'Texto  origen
      sInstrumento = ""
      if cstr(rsx(4) & "") <> "" then
        Strsql = "Select descri From Instrumen Where objid = " & rsx(4)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
    SArma = "" & sInstrumento & "&nbsp;" & rsx(5) & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    if rsx(7) <> "" then

       wUrl = rsx(7) & ""

       if not isnull(wUrl) then
          slinkD = instr(wUrl,"://")
          if slinkD > 0 then
             slinkD = wUrl
          else
             slinkD = "HTTP://" & wUrl
          end if
       else
          slinkD = "javascript:void(0);"
       end if

       sInstlink = slinkD
            sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href='" & sInstlink & "' target=_blank "
	    sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
            sInstlink = sinput2
    else
       sInstlink = "---"
    end if 

    sHTML = sHTML & "<td  class=ut align=center >"
      'Texto  identificacion

    SArma = "" & sInstlink & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    if rsx(8) <> "" then

       wUrl = rsx(8) & ""

       if not isnull(wUrl) then
          slinkD = instr(wUrl,"://")
          if slinkD > 0 then
             slinkD = wUrl
          else
             slinkD = "HTTP://" & wUrl
          end if
       else
          slinkD = "javascript:void(0);"
       end if

       sInstlink = slinkD

            sinput2 = "<a  class=nt style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	    sinput2 = sinput2  & "none" & chr(34) & ";' href='" & sInstlink & "' target=_blank "
	    sinput2 = sinput2  & "  >" & sImagenEM & "</a>"
            sInstlink = sinput2
    else
       sInstlink = "---"
    end if 

    sHTML = sHTML & "<td  class=ut align=center >"
      'Texto  identificacion

    SArma = "" & sInstlink & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "</tr>"

    rsx.movenext
  loop

  call ofv.cerrarconsulta(rsx)


'''''''''''''''''''




  sHTML = sHTML & OFV.genFtabla()
  sHTML = sHTML & "</center>"
  sHTML = sHTML & "</DIV>"

 '----------------------------------------------


    'Encabezado empresa

 if SESSION("EMPMUN") = "S" then

    codempre  = rs(7)
    codempreL = "N"
    h1        = ""

      if SESSION("EMPMUN") = "S" and SESSION("GLBLOG") = 2 and codempre = 1 and session("un") > 1 THEN
         codempre = SESSION("CLIEMP")
         codempreL = "&nbsp;&nbsp;<A  href='javascript:void(0);' title='Documento Corporativo.' ><B><FONT FACE=corflaticon COLOR=#993333 SIZE=5>4</FONT></B></a>"
      END IF

    'Texto empresa
      Strsql = "Select * From clienteplan where id = " & codempre
      set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rsx.eof then
         sInput1 = rsx(1)
         sInput2 = rsx(5)
      end if

    sinput = "<img  src=../COREss20/Empresa/" & sInput2 & " >" & "&nbsp;&nbsp;&nbsp;<span style='position:relative;top:-2;' >" & sinput & "</span>"
    h1 = h1 & ofv.GenCelda("","ttt valign=middle ","","1%","","","",sinput,"si")

    sinput = sinput1
    h1 = h1 & ofv.GenCelda("","ttt valign=middle ","","50%","","","",sinput,"si")

    if codempreL <> "N" then
       h1 = h1 & ofv.GenCelda("","vt ","right","5%","","","",codempreL,"si")
    end if

    fila = ofv.GenRow("","","","","","","",h1,"si")

  sHTML = sHTML & ofv.GenTabla("","aaa  ","","98%","",0,0,2,2,"","",Fila,"si")

 end if


    sHTML = sHTML & "<center>"

  sHTML = sHTML & ofv.GenTabla("","","","98%","",0,0,2,2,"","","","no")


    'Encabezado Descripcion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","ver:","si")
    'Texto descri
           sLink   = "<button value='ver' onclick='chequear();' >"
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  sLink,"si")

  '  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")





      if Nestado=5 then 


               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)  & "&ID2=" & rs(0)  & "&ID3=" & rs(0)
      
      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","center","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")


              'Encabezado Descripcion
               h1 = ofv.GenCelda("","tt  ","","20%","","","","Observaciones:","si")

               h1 = h1 & Celdaobs
               h1 = h1 & ofv.GenCelda("","uti","","80% colspan=2","","","",  "&nbsp;","si")
               sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


      end if







    'Encabezado Descripcion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Descripcion:","si")
    'Texto descri
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  rs(3),"si")


    'Encabezado Identificacion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Identificacion:","si")
    'Texto ident
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(1),"si")



    'Encabezado Identificacion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","PRODUCTO:","si")

    'Texto ident
     sInstrumento = "S/D"
      if cstr(rs(9) & "") <> "" then
        Strsql = "Select descri From TIPOPRODUCTO Where objid = " & rs(9)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","", sInstrumento,"si")

    if  sInstrumento <> "S/D" then
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if


    'Encabezado Tema
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Tema:","si")
'    h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    'Texto Tema
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",stTema,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    'Encabezado Subtema
    h1 = ofv.GenCelda("","tt  ","","20%","","","","SubTema:","si")
'    h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    'Texto subTema
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",stSubtema,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    if sRoot <> "" then
      'Encabezado ruta
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Carpeta:","si")
'   h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
    'Texto ruta
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRoot,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    end if


    sRem = rs(4)
    sRdo = rs(5)

  end if

  call ofv.cerrarconsulta(rs)



''''''''''''''''
  sImagenI3 = "<FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>&#0232;</FONT>"

  Strsql="SELECT OBJPRTY.*, Estados.estado as estdes FROM OBJPRTY, Estados "
  Strsql=Strsql & "where OBJPRTY.estado = Estados.objid and codobjeto=" & nCodigo

  set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof Then

     if svienemas = "SI" then

       SMASINS = ""


        sinput2 = "<a class=nt  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
	sinput2 = sinput2  & " underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
	sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "shwOrig('"& rs(0) &"','" & rs(1) & "');" & chr(34)
   	sinput2 = sinput2  & " TITLE='Ver origenes del documento.' >" & sImagenI3 & "&nbsp;VER&nbsp;LISTA&nbsp;...&nbsp;</a>"
        SMASINS = sinput2

     else

        SMASINS = sOrigen1 & ":&nbsp;" & sInstrumento1 & "&nbsp;" & sInstiden1 & "&nbsp;&nbsp;" & sInstlink1  & "&nbsp;&nbsp;" & sInstlink2  & ""


     end if

    'Encabezado Origen 
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Origen:","si")
    if  svienemain <> "SI" then
        SMASINS = "S/D"
    end if
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",SMASINS,"si")


    if  SMASINS <> "S/D" then
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if


    'Encabezado Nro de acta
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Nro.de Acta:","si")
    'combo/texto estado
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(17),"si")

    if  rs(17) & "" <> "" then
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if

    'Encabezado Estado
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Estado:","si")
    'combo/texto estado
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(21),"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    'Encabezado version
    h1 = ofv.GenCelda("","tt  ","","20%","","","","version:","si")
    'combo/texto estado
    sVersion = "ORIGINAL"
    if rs(14) = "V" then
       sVersion = rs(15)
    end if  
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sVersion,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    'Encabezado Reemplazo
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Reemplaza:","si")
    'combo/texto estado
    sRemDes = "S/D"
    if sRem <> "" then
        Strsql = "Select * From csObjetosRdo Where codobjeto = " & srem
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then 
           sRemDes = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6)
        end if
        call ofv.cerrarconsulta(rsOrigen)
    end if
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")

    if  sRemDes <> "S/D" then
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if

   if cint(rs("estado")) = 4 then

    'Encabezado Reemplazada
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Reemplazada por:","si")
    'combo/texto estado
    sRemDes = "S/D"
    if srdo <> "" then
        Strsql = "Select * From csObjetosRdo Where codobjeto = " & srdo
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then 
           sRemDes = rsOrigen(1) & ": " & rsOrigen(2) & " - " & rsOrigen(3) & " - " & rsOrigen(5) & " - " & rsOrigen(6)
        end if
        call ofv.cerrarconsulta(rsOrigen)
    end if
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")

    if  sRemDes <> "S/D" then
        sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if

   end if

    'Encabezado  Fechas
    h1 = ofv.GenCelda("","ttt colspan=5 ","","20%","","","","Fechas","si")
    sHTML = sHTML & ofv.GenRow("","","","","25","","",h1,"si")




    h1 = ""
    h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Creacion","si")
    h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Revision","si")
    h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Publicacion","si")
    h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Vigencia","si")
    h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Baja","si")
    h1 = h1 & ofv.GenCelda("","ttr  ","","17%","","","","Acta","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")








    h1 = ""
    'texto fecha de creacion
    if clng(rs(6)) = 0 then
      dFecha = "-"
    else
      dFecha = cdate(clng(rs(6)) + cdate("01-01-1900"))
    end if
    h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

    'texto fecha de Vencimiento

    if clng(rs(13)) = 0 then
      dFecha = "-"
    else
      dFecha = cdate(clng(rs(13)) + cdate("01-01-1900"))
    end if
    h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

    'texto fecha de publicacion

    if clng(rs(20)) = 0 then
      dFecha = "-"
    else
      dFecha = cdate(clng(rs(20)) + cdate("01-01-1900"))
    end if
    h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

    'texto fecha de vigencia

    if clng(rs(7)) = 0 then
      dFecha = "-"
    else
      dFecha = cdate(clng(rs(7)) + cdate("01-01-1900"))
    end if
    h1 = h1 & ofv.GenCelda("","vt style='color:#990000;' ","","","","","",fechadorCor(dFecha,""),"si")

    'texto fecha de eliminacion

    if clng(rs(8)) = 0 then
      dFecha = "-"
    else
      dFecha = cdate(clng(rs(8)) + cdate("01-01-1900"))
    end if
    h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")

    'texto fecha de acta

    if clng(rs(16)) = 0 then
      dFecha = "-"
    else
      dFecha = cdate(clng(rs(16)) + cdate("01-01-1900"))
    end if
    h1 = h1 & ofv.GenCelda("","ut","","","","","",fechadorCor(dFecha,""),"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")




    sHTML = sHTML & ofv.GenFTabla()



  Strsql = "Select * From ObjTSEGAGR Where codobjeto=" & nCodigo & " Order By OBJID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS.EOF THEN

    sHTML = sHTML & ofv.GenTabla("","","","98%","",0,0,2,2,"","","","no")
      Celda = ofv.GenCelda("","TTr","","40%","","","","AGRUPAMIENTO","si")
      Celda = Celda & ofv.GenCelda("","TTr","","40%","","","","SEGMENTO","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

  do until rs.eof

      sInput = "S/D"
      if cstr(rs(5)&"") <> "" then
         Strsql = "Select descri From TIPOAGRUPA where OBJID = " & rs(5)
         set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rsInstrumento.eof then sInput = rsInstrumento(0)
         call ofv.cerrarconsulta(rsInstrumento)
      end if

      Celda = ofv.GenCelda("","ut","","40%","","","",sInput,"si")

    'SEGMENTO

     sInput = "S/D"
      if cstr(rs(4)&"") <> "" then
         Strsql = "Select descri From TIPOSEGMENTO where OBJID = " & rs(4)
         set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if not rsInstrumento.eof then sInput = rsInstrumento(0)
         call ofv.cerrarconsulta(rsInstrumento)
      end if

    Celda = Celda & ofv.GenCelda("","ut","","40%","","","",sInput,"si")

    sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

   RS.MOVENEXT
   LOOP

   sHTML = sHTML & ofv.GenFTabla()

  END IF




  end if




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</center>"

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML



%>