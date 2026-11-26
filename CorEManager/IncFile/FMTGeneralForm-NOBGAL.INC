
<%

 FUNCTION GENFMTGENERALFORM(FNCODOBJ,FNTIPOBJ)


  FNSHTML = ""




redim tobj(0)
redim tobjd(0)

CLASEFILA="istt"' bgcolor=#f05f23 style='font-weight:bold;color: #ffffff;'"
CLASEFILAT="ut"' style='font-size:12px;'"
CLASEFILATF="xxxc style='color: #000000;'"
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
  'response.write strsql
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

'response.write stSubtema
  Strsql="Select * From Objetos Where codobjeto=" & nCodigo
  set rs=ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof Then
    sModifica = session("Modificar")
    vuelta= ofv.vueltaasp2("../")

    wdescri ="<i>" & rs(3) & "</i>"

'----------------------------------------------

 sinput2 = "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=webdings size=1   >r</font>"

  FNSHTML = FNSHTML & "<DIV id='cx1' class='DINDM' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;"
  FNSHTML = FNSHTML & "position:absolute;left:100;top:160;width:600px;height:250px;overflow:hidden;'>"

'''''''''''''''''''''

  FNSHTML = FNSHTML & OFV.gentabla("", "TWIND", "", "100% ", "", "0", "", "-1", "-2", "", "", "", "no")
  FNSHTML = FNSHTML & "<tr>"
  FNSHTML = FNSHTML & "<td CLASS=mbt width=99%  >&nbsp;&nbsp;ORIGEN&nbsp;DE&nbsp;LA&nbsp;DOCUMENTACION</td>"
  FNSHTML = FNSHTML & "<td CLASS=mbt width=1% ONCLICK=""document.all.cx1.style.display='none';"" title='Ocultar.' >"
  FNSHTML = FNSHTML & sinput2
  FNSHTML = FNSHTML & "</td>"
  FNSHTML = FNSHTML & "</tr>"
  FNSHTML = FNSHTML & OFV.genFtabla()
  FNSHTML = FNSHTML & "<center>"
  FNSHTML = FNSHTML & OFV.gentabla("", "aaa ", "", "98%", "", "0", "", "0", "0", "", "", "", "no")
  FNSHTML = FNSHTML & "<tr height=22 ><td width=40% "
  FNSHTML = FNSHTML & "class=ttt  >Origen</td>"
  FNSHTML = FNSHTML & "<td width=40% "
  FNSHTML = FNSHTML & "class=ttt  >Instrumento</td>"
  FNSHTML = FNSHTML & "<td width=10% "
  FNSHTML = FNSHTML & "class=ttt  >L.Instrumento</td>"
  FNSHTML = FNSHTML & "<td width=10% "
  FNSHTML = FNSHTML & "class=ttt  >L.Analisis</td></tr>"

 '' FNSHTML = FNSHTML & "<tr><td  "
 '' FNSHTML = FNSHTML & "class=UT colspan=4  >&nbsp;</td></tr>"

  tipocpt="D"
  Strsql="SELECT OBJPRTY.*, Estados.estado as estdes,tipocpt,tipobjeto.descri as tobj FROM OBJPRTY "
  strsql=Strsql  & " inner join Estados on (OBJPRTY.estado = Estados.objid) "
   strsql=Strsql  & " inner join tipobjeto on (tipobjeto.tipobjeto = OBJPRTY.tipobjeto) "
  Strsql=Strsql & "where OBJPRTY.codobjeto=" & nCodigo
'response.write strsql
  set rsx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx.eof Then
   sVersion = "ORIGINAL"
    if rsx(14) &"" = "V" then
       sVersion = " VERSION (" & rsx(15) & ")"
    end if
     nestado = rsx(9)
     sTipo   = rsx(2)
     tipocpt=rsx(22)
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

    FNSHTML = FNSHTML & "<tr height=22 bgcolor='" & xTrColor & "' >"

    FNSHTML = FNSHTML & "<td  class=ut>"
      'Texto  origen


    SArma = "" & "&nbsp;" & sOrigen1 & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

    FNSHTML = FNSHTML & "<td  class=ut>"
      'Texto  instrumento

    SArma = "" & sInstrumento1 & "&nbsp;" & sInstiden1 & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"



    FNSHTML = FNSHTML & "<td  class=ut align=center >"
      'link consulta

    SArma = "" & sInstlink1 & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

    FNSHTML = FNSHTML & "<td  class=ut align=center >"
      'link consulta

    SArma = "" & sInstlink2 & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

    FNSHTML = FNSHTML & "</tr>"



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

    FNSHTML = FNSHTML & "<tr height=22 bgcolor='" & xTrColor & "' >"

    FNSHTML = FNSHTML & "<td  class=ut>"
      'Texto  origen
      sOrigen = ""
      if cstr(rsx(3) & "") <> "" then
        Strsql = "Select descri From Origen Where objid = " & rsx(3)
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
      end if

    SArma = "" & "&nbsp;" & sOrigen & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

    FNSHTML = FNSHTML & "<td  class=ut>"
      'Texto  origen
      sInstrumento = ""
      if cstr(rsx(4) & "") <> "" then
        Strsql = "Select descri From Instrumen Where objid = " & rsx(4)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
    SArma = "" & sInstrumento & "&nbsp;" & rsx(5) & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

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

    FNSHTML = FNSHTML & "<td  class=ut align=center >"
      'Texto  identificacion

    SArma = "" & sInstlink & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

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

    FNSHTML = FNSHTML & "<td  class=ut align=center >"
      'Texto  identificacion

    SArma = "" & sInstlink & ""
    FNSHTML = FNSHTML & SArma

    FNSHTML = FNSHTML & "</td>"

    FNSHTML = FNSHTML & "</tr>"

    rsx.movenext
  loop

  call ofv.cerrarconsulta(rsx)


'''''''''''''''''''




  FNSHTML = FNSHTML & OFV.genFtabla()
  FNSHTML = FNSHTML & "</center>"
  FNSHTML = FNSHTML & "</DIV>"

 '----------------------------------------------


    'Encabezado empresa

 if SESSION("EMPMUN") = "S" then

    codempre  = rs(7)
    codempreL = "N"
    h1        = ""

      if SESSION("EMPMUN") = "S" and SESSION("GLBLOG") = 2 and codempre = 1 and session("un") > 1 THEN
         codempre = SESSION("CLIEMP")
         codempreL = "&nbsp;&nbsp;<A  href='javascript:void(0);' title='Documento Corporativo.' ><B><FONT FACE=WingDINGS COLOR=#CCCCFF SIZE=5>4</FONT></B></a>"
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

  FNSHTML = FNSHTML & ofv.GenTabla("","aaa  ","","98%","",0,0,2,2,"","",Fila,"si")

 end if


    FNSHTML = FNSHTML & "<center>"

  FNSHTML = FNSHTML & ofv.GenTabla("","","","98%","",0,0,2,2,"","","","no")


    'Encabezado Descripcion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","ver:","si")
    'Texto descri
           sLink   = "<button value='ver' onclick='chequear();' >"
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  sLink,"si")

  '  FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")





      if Nestado=5 then


               sAccion = "genObsdoc.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(0)  & "&ID2=" & rs(0)  & "&ID3=" & rs(0)

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='block';"" "
      	       sHref   = sAccion
      	       sInput  = "<div style='border-style:solid;border-color:#CCCCFF;border-width:1;' >" & "OBS" & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut" & xsffun & "  title='Agregar/consultar observaciones para el elemento.' ","center","","1","","",sInput,"si")
               Celdaobs  = ofv.GenLink("","ut" & "  ",sHref,"WK4","","",Celdaobs,"si")


              'Encabezado Descripcion
               h1 = ofv.GenCelda("","tt  ","","20%","","","","Observaciones:","si")

               h1 = h1 & Celdaobs
               h1 = h1 & ofv.GenCelda("","uti","","80% colspan=2","","","",  "&nbsp;","si")
               FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")


      end if







    'Encabezado Descripcion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Descripcion:","si")
    'Texto descri
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",  rs(3),"si")


    'Encabezado Identificacion
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Identificacion:","si")
    'Texto ident
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(1),"si")






    sRem = rs(4)
    sRdo = rs(5)

  end if

  call ofv.cerrarconsulta(rs)



''''''''''''''''
  sImagenI3 = "<FONT FACE=WingDINGS COLOR=#f7aa22 SIZE=2>&#0232;</FONT>"


 Strsql="SELECT OBJPRTY.*, Estados.estado as estdes,tipocpt ,tipobjeto.descri as tobj,autor ,isnull(TIPOPRODUCTO.descri,'---') as producto FROM OBJPRTY "
 strsql=Strsql  & " inner join objetos on (OBJPRTY.codobjeto = objetos.codobjeto) "

 strsql=Strsql  & " inner join Estados on (OBJPRTY.estado = Estados.objid) "
 strsql=Strsql  & " inner join tipobjeto on (tipobjeto.tipobjeto = OBJPRTY.tipobjeto) "
 strsql=Strsql  & " left join TIPOPRODUCTO on (TIPOPRODUCTO.objid = objetos.tsegmento) "
  Strsql=Strsql & "where OBJPRTY.codobjeto=" & nCodigo

'response.write strsql
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
        FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if


    'Encabezado Nro de acta
    h1 = ofv.GenCelda("","tt  ","","20%","","","","Nro.de Acta:","si")
    'combo/texto estado
    h1 = h1 & ofv.GenCelda("","uti","","80% colspan=3","","","",rs(17),"si")

    if  rs(17) & "" <> "" then
        FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
    end if

	    if clng(rs(7)) = 0 then
	      dFecha = "-"
	    else
	      dFecha = cdate(clng(rs(7)) + cdate("01-01-1900"))
    end if



	    estadoDoc="<i>" & rs("estdes") & " </i>"
        tobjdsc="<i>" & rs("tobj") & "</i>"
        sproducto=rs("producto")
         sautor=""

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
        FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",h1,"si")
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

   h1 =ofv.GenCelda("","uti","","80% colspan=3","","","",sRemDes,"si")



end if
if rs("tipocpt")="K" and  session("APLMDCLI")="BGAL" then


  strsql=" select "
  strsql= strsql & " b.descripcion as guarda,"
  strsql= strsql & " c.descripcion  as categoria,"
  strsql= strsql & " d.descripcion as claseform ,"
  strsql= strsql & "  a.codigoform ,"
  strsql= strsql & "  f.descripcion as firmas , "

  strsql= strsql & "  A.MODOIMPRESION ,"
  strsql= strsql & "  A.IMPRESION ,"
   strsql= strsql & "  g.descripcion as COPIAS,"
   strsql= strsql & "  f1.descripcion,A.DESCRIPCION,"
   strsql= strsql & "  g1.descripcion as proceso,"
   strsql= strsql & "  g2.descripcion as BANCA"
  strsql= strsql & "  from objpropiedades a "
  strsql= strsql & "  left join FORMADMTGUARDA b on ( a.tiempoguarda=b.objid) "
  strsql= strsql & "   left join FORMADMCATG c on ( a.catform=c.objid) "
  strsql= strsql & "   left join FORMADMTFORMU d on ( a.claseform=d.objid) "
  'strsql= strsql & "   left join FORMADMCATG e on ( a.codigoform=e.objid) "
  strsql= strsql & "   left join FORMADMTFIRMA f on ( a.firmasform=f.objid) "
   strsql= strsql & "   left join FORMADMPROCESO f1 on ( a.proceso=f1.objid) "
  strsql= strsql & "   left join FORMADMTCOPIAS g on ( a.copias=g.objid) "
  strsql= strsql & "   left join FORMADMPROCESO g1 on ( a.proceso=g1.objid) "
  strsql= strsql & "   left join FORMADMBANCA g2 on ( a.BANCA=g2.objid) "

  strsql= strsql & "   where a.codobjeto=" & nCodigo & " and tipobjeto='" & stipo & "'"
'response.write strsql
set rsprop=ofv.crearconsultaEx(StrSql,cn,1,parametros)

 if not rsprop.eof then

       codigoform=rsprop("codigoform")
       h1 = ofv.GenCelda("",CLASEFILA,"","10%","","","","Versión: ","si")
       h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sVersion,"si")


           h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","","Banca: ","si")

           SDATO="---"
            if not isnull(rsprop("BANCA")) then
            SDATO=rsprop("BANCA")
            END IF

   		   h1 =  h1 &  ofv.GenCelda("",CLASEFILAT,"","45%","","","",SDATO,"si")
		   FNSHTMLP =   ofv.GenRow("","","","","","","",h1,"si")

            h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","","Modo  de Impresion: ","si")
             sdato = "---"
             if not isnull(rsprop("MODOIMPRESION")) then
			   			 IF rsprop("MODOIMPRESION")="SF" THEN
			   			  sdato = "SIMPLE FAZ"
			   			 ELSE
			   			  sdato = "DOBLE FAZ"
			   			 END IF


   		    end if
   		    h1 =  h1 &  ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")

  		    h1 =  h1 &  ofv.GenCelda("",CLASEFILA,"","","","","","Firmas: ","si")
  		     sdato="---"
			             if not isnull(rsprop("firmas")) then

						   	 sdato = replace(rsprop("firmas")," ","&nbsp;")


						  END IF


			   		    end if
   		    h1 =  h1 &  ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")


           FNSHTMLP =  FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

		   sdato = "---"


          h1 =  ofv.GenCelda("",CLASEFILA,"","5%","","","","Tiempo&nbsp;Guarda:","si")

          sdato = "---"

          if not isnull(rsprop(0)) then
            sdato = rsprop(0)
          end if

           h1 =h1 & ofv.GenCelda("",CLASEFILAT ,"","40%","","","", sdato ,"si")


           h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","","Producto:","si")

           h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sproducto,"si")
	       FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")


           h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Tipo&nbsp;de&nbsp;Formulario:","si")

            sdato = "---"

		      if not isnull(rsprop("claseform")) then
		   		 sdato = rsprop("claseform")
              end if
           h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")


		    h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","", "Proceso","si")
		   	sdato = "---"

			 if not isnull(rsprop("proceso")) then
			   sdato = rsprop("proceso")
			 end if

            h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")

           FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")

	  	    sdato = "---"
  		    h1 = ofv.GenCelda("",CLASEFILA,"","5%","","","", "Categoria","si")
  		   sdato = "---"

			 if not isnull(rsprop("categoria")) then
			   sdato = rsprop("categoria")
			 end if

            h1 = h1 & ofv.GenCelda("",CLASEFILAT,"","45%","","","",sdato,"si")

            h1 = h1 & ofv.GenCelda("",CLASEFILA,"","","","","","&nbsp;","si")
             h1 = h1 & ofv.GenCelda("","","","45%","","","","&nbsp;","si")

   		   FNSHTMLP = FNSHTMLP & ofv.GenRow("","","","","","","",h1,"si")
           FNSHTMLP = ofv.GenTabla("","aaa border style='border-style:solid;border-width:1;border-color:#f05f23;' cellspacing=2 ","","50%","","0","0","0","0","","",FNSHTMLP,"si")


   'Encabezado Tema

      descri="<b>UBICACION INDICE:<br>" & "Tema: " & stTema & "&nbsp;&nbsp;&nbsp;&nbsp;Subtema: " & stSubtema & "</b>"
      h1 = ofv.GenCelda("",CLASEFILA,"","","","","",descri,"si")


      FNSHTMLC2 = FNSHTMLC2 & ofv.GenRow("","","","","","","",h1,"si")

      TABLATEMA = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23; '","","50%","","0","0","0","0","","",FNSHTMLC2,"si")




      FNSHTMLC2=""


		   if not isnull(rsprop("DESCRIPCION")) then
			 sdato=REPLACE(rsprop("DESCRIPCION"),CHR(13),"<BR>")
			 'DESCRIPCION
			 	        h1 = ofv.GenCelda("","cc bgcolor=#CCCCFF ","","20%","","","","DESCRIPCION:","si")
			 	    '    h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
			 	        'Texto subTema
			 	        FNSHTMLC2 = FNSHTMLC2 &   ofv.GenRow("","","","","","","",h1,"si")



			h1 = ofv.GenCelda("","vt","","","","","", SDATO  ,"si")
			FILAZ =  ofv.GenRow("","","","","","","",h1,"si")
            TABLADESCRI = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;'","","50%","","0","0","0","0","","",FILAZ,"si")
		   END IF



  '    if sRoot <> "" then
        'Encabezado ruta
  '    h1 = ofv.GenCelda("","cc bgcolor=#CCCCFF","","20%","","","","Carpeta:","si")
  '   h1 = h1 & ofv.GenCelda("","","","","","","","&nbsp;","si")
      'Texto ruta
  '    FNSHTMLC2 = FNSHTMLC2 &   ofv.GenRow("","","","","","","",h1,"si")
   '   h1 =  ofv.GenCelda("",CLASEFILA,"","80% colspan=3","","","",sRoot,"si")
   '   FNSHTMLC2 = FNSHTMLC2 & ofv.GenRow("","","","","","","",h1,"si")
   '   END IF

  '     FNSHTMLC2 = ofv.GenTabla("","aaa cellspacing=0 cellpadding=5px ","","50%","","0","0","0","0","","",FNSHTMLC2,"si")




  end if

















  dfecha="&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Fecha de Vigencia : <i>" & fechadorCor(dFecha,"") & "</i>"
 ' FNSHTML= "<br><TABLE  style='border-style:solid;border-width:1;border-color:black; '  WIDTH=80% cellspacing=2px cellpadding=4px >"
 ' FNSHTML=  FNSHTML & " <TR ><TD class=cc xalign=center  COLSPAN=2><font size=2 >"& tobjdsc & "&nbsp;:&nbsp;"  & 	wdescri  &  "</font></TD></TR>"

  'FNSHTML=  FNSHTML & " <TR 	 ><TD class=CC align=center valign=middle COLSPAN=2>CODIGO DEL FORMULARIO :" & sdatoc& "&nbsp;</TD></TR>"
  celda= ofv.GenCelda("",clasefila,"","","","","","Estado :" & estadoDoc & dfecha ,"si")
  FNSHTML = ofv.GenRow("","","","","","","",celda,"si")
  tablaestado  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;' ","","50%","","0","0","0","0","","",FNSHTML,"si")

if rs(2) <> "9" then
   nom= "<h5>" & codigoform & " - " &  wdescri & "</h5>"
else
   nom= "<h5>" & wdescri & "</h5>"
end if
   ' FNSHTML= "<br><TABLE  style='border-style:solid;border-width:1;border-color:black; '  WIDTH=80% cellspacing=2px cellpadding=4px >"
   ' FNSHTML=  FNSHTML & " <TR ><TD class=cc xalign=center  COLSPAN=2><font size=2 >"& tobjdsc & "&nbsp;:&nbsp;"  & 	wdescri  &  "</font></TD></TR>"

    'FNSHTML=  FNSHTML & " <TR 	 ><TD class=CC align=center valign=middle COLSPAN=2>CODIGO DEL FORMULARIO :" & sdatoc& "&nbsp;</TD></TR>"
    celda= ofv.GenCelda("",CLASEFILA,"","","","","", nom ,"si")
    FNSHTML = ofv.GenRow("","","","","","","",celda,"si")
   tablanombreform  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;' ","","50%","","0","0","0","0","","",FNSHTML,"si")



  'FNSHTML=    " <TR><TD class=" & clasefila & " xalign=center valign=middle COLSPAN=2>"


  FNSHTML=""

'  DESCARGAR
   Strsql = "Select DESCRI ,  ARCHIVO ,DIRECTORIO,tipo  From OBJHTM Where codobjeto=" & nCodigo & " Order By OBJID"
    'response.write strsql
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS.EOF THEN

      FNSHTML =  ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#f05f23;' ","","50%","",00,0,10,0,"","","","no")
      Celda = ofv.GenCelda("",clasefila & "  COLSPAN=3 ALIGN=CENTER","","","","","","<font size=2> DESCARGA EL FORMULARIO</font>","si")

      FNSHTML = FNSHTML & ofv.GenRow("","","","","","","",Celda,"si")
      POS=1
   DO WHILE NOT rs.EOF

         If xTrColor = "#FFFFFF" Then
			xTrColor = "#FFFFFF" '"#F0F0F0"
		 Else
			xTrColor = "#FFFFFF"
         End If

         Celda = ofv.GenCelda("","ut","center","5%","","","","&nbsp;","si")

         if cstr(rs(0)&"") <> "" then
	        sInput= "<font size=2> " & REPLACE(RS(0) &"",CHR(13),"<BR>") & "</font>"
	     end if
		 Celda = CELDA & ofv.GenCelda("",CLASEFILATF,"","90%","","","", sInput,"si")



        sInput = "S/D"
		if cstr(rs(1)&"") <> "" then
		   sImagenP = "<img width=100px src='imagenes/bajada2.gif'" '"<b><font face=WINGdings size=4 color=#F7AA22>&#052;</font></b>"
           if rs("tipo")=3 then
		    sinput = "<a class=nt href='" & RS("archivo")
		   else
		    sinput = "<a class=nt href='CoArchivos3.asp?Archivo=" &  RS("DIRECTORIO") & "/" & rs(1)
		   end if
		   sinput = sinput  & "' title='Descargar' target='_blank' >" &  sImagenP & "</a>"

        end if

        Celda = CELDA & ofv.GenCelda("","ut","center","5%","","","",sInput,"si")

        FNSHTML = FNSHTML & ofv.GenRow("","ut  bgcolor='" & xTrColor & "'  ","","","","","",Celda,"si")

     RS.MOVENEXT
     POS=POS +1
     LOOP

    FNSHTML =FNSHTML & "</table>"


  END IF
  END IF


  FNSHTML=  tablanombreform  & "<br><br>" & tablaestado & "<br><br>" & FNSHTML  &  "<br><br>" & FNSHTMLP  & "<br><br>" & TABLATEMA  &"<br><br>" & TABLADESCRI

   FNSHTML = FNSHTML & ofv.GenFTabla()
  GENFMTGENERALFORM = "<div id=zglos><center>" &  FNSHTML  & "</center></div>"


 END FUNCTION




%>