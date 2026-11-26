<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

'response.write "22"
  sTipo = request.querystring("Tipo")

  if sTipo="Consulta" then response.redirect "ConsTema.asp?" & request.querystring


  session("xorigen") = "tema"

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
  LNombre = request.querystring("Nombre")
  sNombre = request.querystring("Nombre")
  sNombre = ofv.convertircar(sNombre,"_"," ")
  if request.querystring("Menu") <> "" then Session("Menu") = request.querystring("Menu")
  sID = request.querystring("ID")
  if sNombre = "" Then sNombre = UCase(sTipo)
  sitem = ofv.convertircar(request.querystring("sitem"),"_"," ")
  stitem = ofv.convertircar(request.querystring("stitem"),"_"," ")
  sticono = request.querystring("sticono")
  stnivel = request.querystring("stnivel")

  xwbody = "bgcolor=#ccbbaa"
  If BrowserName() = "IE" OR BrowserName() = "CHROME"  then xwbody = "bgcolor=#dddddd"

 ''' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"


  ssTipo = sTipo
  session("sEstadon") = "1"
  session("sNivelCero") = sTipo
  if sTipo="Temas" then
    session("sEstado") = "APROBADA"
    session("sEstadon") = "3"
  elseif sTipo="Desarrollo" then
    session("sEstado") = "DESARROLLO"
    session("sEstadon") = "1"
  elseif sTipo="ParaAprobar" then
    session("sEstado") = "PARA APROBAR"
    session("sEstadon") = "2"
  elseif sTipo="Eliminadas" then
    session("sEstado") = "ELIMINADA"
    session("sEstadon") = "4"
  elseif sTipo="Consulta" then
    session("sEstado") = "APROBADA"
    session("sEstadon") = "3"
  end if
  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "Select estado from estados where objid = " & session("sEstadon")
 ' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     ssTipo = rs(0)
  end if
  call ofv.cerrarconsulta(rs)


'ofv.CorFormHeader = CorFormHeader2(session("FormN"),"N")
'sHTML = ofv.formheader("")

  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
  sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
  sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
  sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"

funcion = ""

if recarga = "S" AND sTipo <> "Consulta" then
   funcion = " onload='document.recarga.submit();' "
end if

sHTML = ofv.MENUheader("",funcion)




  sNombre = ofv.convertircar(sNombre," ","_")

  HrfAsp = "Objetos.asp?"
  HrfTrg = "fesquemas"

  if sTipo="Consulta" then
    sHTML = sHTML & "<center>"
    sstipo = session("FormN")
    HrfAsp = "ObjTreeSubTemaCons.asp?"
    HrfTrg = ""
  end if

  Fila = ""



     SINPUT3 = "&nbsp;<A title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='mostrarmenu();' >" & sImagenEM & "</a>"

  sHref = ""
  sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
  sLink = ofv.GenLink("","nt",sHref,"ftemas","","",sImage,"si")
  Celda = ofv.GenCelda("","aaa colspan=3 ","center","2%","","","",sLink,"si")
 ' sLink = ofv.GenLink("","VT",sHref,"ftemas","","","<b><u>" & ATREMP(0,4) & "</u></b>","si")
  Celda = Celda & ofv.GenCelda("","aaa ","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","VT colspan=3 ","","","","","","<b><u>" & ATREMP(0,4) & "</u></b>","si")
  Celda = Celda & ofv.GenCelda("","","","","","","",SINPUT3,"si")
  Fila = fila & ofv.GenRow("","","top","","","","",Celda,"si")

  Celda = ofv.GenCelda("","VT colspan=8 STYLE='BORDER-TOP-STYLE:RIDGE;BORDER-TOP-COLOR:#F7AA22;BORDER-TOP-WIDTH:2;' ","","","","","","&nbsp;","si")
  Fila = fila & ofv.GenRow("","","top","","","","",Celda,"si")


   Tabla = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")


           shtml=shtml & "<script language=javascript>" & chr(13)
            shtml=shtml & "function mostrarmenu(){" & chr(13)
            shtml=shtml & "document.getElementById('x1').style.display='';" & chr(13)
            shtml=shtml & "}" & chr(13)
            shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
            sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.getElementById('x1').style.display='none';" & chr(34)
            sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
            TABLA = Tabla & "<DIV id='x1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:10px;width:250px;height:250px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


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

      sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
      sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
      TABLAX = TablaX & sLink
      TABLAX = TablaX & "</td>"
      TABLAX = TablaX & "</tr>"
     END IF

  NEXT

  TABLAX = TablaX & OFV.genFtabla()

  TABLA = Tabla & TABLAX

  TABLA = Tabla & "</DIV>"



  TABLA = Tabla & "<DIV id='x2' STYLE='display:BLOCK;z-index:1;margin-top:1;margin-left:1;overflow:auto;'>"
  Fila = ""

  sHref = ""
  sImage = ofv.GenImage("","","","","Iconos\ICorMMenu.jpg","0","")
  sLink = ofv.GenLink("","nt",sHref,"ftemas","","",sImage,"si")
  Celda = ofv.GenCelda("","aaa colspan=3 rowspan=2  style='" & Session("CBRDICO") & "'","center","2%","","","",sLink,"si")
  sLink = ofv.GenLink("","cc",sHref,"ftemas","","","<b><u>" & Ucase(sstipo) & "</u></b>","si")
  Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2","","","","","",sLink,"si")
  Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
  Fila = fila & ofv.GenRow("","","top","","","","",Celda,"si")


' end if

  stitem = "Tema"
  sIcono = "ICorTema.jpg"
  sNuevoTipo = "SUBTEMAS"
  nNivel = 0

  sID = 0
  nv = 0
  cborderB = Session("CBRDBOT")
  cborderl = Session("CBRDLFT")
  borderB = ""
  borderl = ""
  Strsql = "Select * From Tema order by descri"
 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
          Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
    Celda = ""

          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")


    Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2","","2%","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
    Celda = Celda & ofv.GenCelda("","aaa  rowspan=2","","","","","","No hay elementos asociados","si")
    Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
  else

'    Fila = ""
    do until rs.eof


       WTVER = 1

       Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
    ''  response.write strsql & "<br>"
       set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rs1.eof then
          if not RS1(4) then  WTVER = 0
         ' response.write strsql & " " & WTVER & "<br>"
       end if

       if WTVER = 1 then

       if nv = 0 then
          Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
       else
          borderl = cborderl
          borderb = ""
          Celda = ""
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

       end if



          borderl = cborderl
          borderb = cborderb
          Celda = ""
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")

      sHref = HrfAsp & "Tipo=" & sNuevoTipo & "&ID=" & rs(0) & "&sitem=TEMA:_"
      sHref = sHref & ofv.convertircar(rs(1)," ","_") & "&stitem="
      sHref = sHref & ofv.convertircar(stitem," ","_") & "&sticono=" & sIcono
      sHref = sHref & "&stnivel=" & nNivel
      sImage = ofv.GenImage(rs(0) & "-" & rs(2),"","","","Iconos\" & sIcono,"0","")
      sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","",sImage,"si")
      Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2 style='" & Session("CBRDICO") & "'","center","2%","","","",sLink,"si")
      sLink = ofv.GenLink("","nt ",sHref,HrfTrg,"","",rs(1),"si")
      Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","aaa bgcolor=" & xTrColor & "   rowspan=2","","","","","",sLink,"si")
      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
      nv = 1


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F0F0F0"
                         ELSE
                            xTrColor = "#FFFFFF"
                         END IF

      end if

      rs.movenext
    loop

  end if
    if session("sEstadon") = "4" then
       if nv = 0 then
          Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
       else
          borderl = cborderl
          borderb = ""
          Celda = ""
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

       end if


          borderl = cborderl
          borderb = cborderb
          Celda = ""
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")


      sHref = HrfAsp & "Tipo=" & sNuevoTipo & "&ID=xxx&sitem=RECICLAJE"
      sHref = sHref & "&stitem="
      sHref = sHref & ofv.convertircar(stitem," ","_") & "&sticono=ICorRecycle.jpg"
      sHref = sHref & "&stnivel=" & nNivel
      sImage = ofv.GenImage("xxx-RYC","","","","Iconos\ICorRecycle.jpg","0","")
      sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","",sImage,"si")
      Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2 style='" & Session("CBRDICO") & "'","center","2%","","","",sLink,"si")
      sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","","RECICLAJE","si")
      Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","",sLink,"si")
      Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
      Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
    end if
          Celda = ""
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

          Celda = ""
          Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","89%","","","","&nbsp;","si")
          Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
          Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  if sTipo="Consulta" then
     Tabla = Tabla & ofv.GenTabla("","aaa  topmargin=4 leftmargin=4 style='"  & Session("CBRDABL") & ";border-style=solid;border-width:1;' ","","95% ","","0","","-1","-2","","",Fila,"si")
  else
     Tabla = Tabla & ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")

     if recarga = "S" then
     Formulario = OFV.GenForm("recarga", "", "nulo2.asp", "fesquemas", "post", "","<input type=hidden value='si' >", "si")
     Tabla = Tabla & Formulario
     end if


  end if

  TABLA = Tabla & "</DIV>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & Tabla & "</body></html>"
end if

response.write sHTML

%>