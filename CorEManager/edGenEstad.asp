<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  OPC = request.querystring("OPC")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")

  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"

  set cn2 = ofv.conectar(ofv.strconn0)
  Strsqlx = "Select * From NivelConfiden order by NCCOD"
  set rs = ofv.crearconsultaEx(StrSqlx,cn2,1,parametros)
  redim preserve CNIV(0)
  redim preserve DNIV(0)
  ii = 0
  if not rs.eof Then
     do until rs.eof
        redim preserve CNIV(ii)
        redim preserve DNIV(ii)
        CNIV(ii) = rs(0)
        DNIV(ii) = rs(1)
        ii = ii + 1
     rs.movenext
     loop

  end if
  call ofv.cerrarconsulta(rs)

  set cn = ofv.conectar(ofv.strconn2)

  IF OPC = "SI" THEN
     ESQ = request.querystring("ESQ")
''     IF ESQ > 0 THEN
        STRSQL = "UPDATE OBJETOS SET OBJREM = " & ESQ & " WHERE CODOBJETO = " & nCodigo
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

''     END IF
  END IF

      sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0143;</FONT></B>"
      sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=2>4</FONT></B>"
      sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"


  Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof Then
 ''   sHTML = ofv.FormBody

    sHTML = ofv.FormHeader(session("FormN"))

    SVIENE = "N"

    Strsql="SELECT * FROM OBJALOG "
    Strsql=Strsql & "where codobjeto=" & nCodigo & " AND TIPOBJETO = '" & STIPO & "' "

    set rsx=ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsx.eof Then


            SVIENE = "S"

            shtml=shtml &  chr(13) & "<script type=" & chr(34) & "text/javascript"& chr(34) & " src='IncFile/dragdrop.js' ></script>" & chr(13)

            shtml=shtml & "<script language=javascript>" & chr(13)
            shtml=shtml & "function mostrarmenu(){" & chr(13)
            shtml=shtml & "document.getElementById('x1').style.display='block';" & chr(13)
            shtml=shtml & "}" & chr(13)
            shtml=shtml & "</script>" & chr(13)
'----------------------------------------------
            sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
            sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.getElementById('x1').style.display='none';" & chr(34)
            sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
            sHTML = sHTML & "<DIV id='x1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:100px;top:10px;width:600px;height:250px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"

'''''''''''''''''''''


  sHTML = sHTML & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "600", "", "0", "", "0", "0", "", "", "", "no")
  sHTML = sHTML & "<tr><td  "
  sHTML = sHTML & "class=vT align=left colspan=2 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;ESTADISTICAS DE ACCESO</td></tr>"
  sHTML = sHTML & "<tr><td colspan=2 "
  sHTML = sHTML & " >&nbsp;</td></tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=vt  >FECHA ULTIMO ACCESO:</td>"

  sHTML = sHTML & "<td class=vt  >"
  sHTML = sHTML & Fechador( RSX(3) - clng(cdate("01-01-1900")) )
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr><td colspan=2 "
  sHTML = sHTML & " >&nbsp;</td></tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=vt  >EDICIONES:</td>"

  sHTML = sHTML & "<td class=vt  >"
  sHTML = sHTML & RSX(4)
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=vt  >CONSULTAS:</td>"

  sHTML = sHTML & "<td class=vt  >"
  sHTML = sHTML & RSX(5)
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=vt  ><FONT COLOR=#993333>TOTAL ACCESOS:</FONT></td>"

  sHTML = sHTML & "<td class=vt  ><FONT COLOR=#993333>"
  sHTML = sHTML & RSX(5) + RSX(4)
  sHTML = sHTML & "</FONT></td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr><td colspan=2 "
  sHTML = sHTML & " >&nbsp;</td></tr>"

  TTOTAL = ROUND(RSX(7) / (RSX(5) + RSX(4)),0)

  sHTML = sHTML & "<tr>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=vt  >TIEMPO PROMEDIO DE ACCESO:</td>"

  sHTML = sHTML & "<td class=vt  >"
  sHTML = sHTML & TTOTAL & "&nbsp;S."
  sHTML = sHTML & "</td>"
  sHTML = sHTML & "</tr>"

  sHTML = sHTML & "<tr><td  "
  sHTML = sHTML & "class=UT align=center colspan=2 style={border-TOP-style:RIDGE;border-TOP-width:2;BORDER-TOP-COLOR:#F7AA22;} >&nbsp;</td></tr>"






'''''''''''''''''''





     sHTML = sHTML & OFV.genFtabla()
     sHTML = sHTML & "</DIV>"

     SINPUT3 = "&nbsp;<A title='Estadisticas de acceso' HREF='javascript:void(0);' onclick='mostrarmenu();' >" & sImagenEM & "</a>"

     END IF
     call ofv.cerrarconsulta(rsx)

 '----------------------------------------------






    wdescri = rs(3)

   Fila = ""

   IF  Session("EMPMUN") = "S" THEN

    'empresa
    Celda = ofv.GenCelda("","tt","","","","","","Empresa","si")
      Strsql = "Select * From clienteplan where id = " & rs(7)
      set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not rsx.eof then
         sInput = rsx(1)
         sInput2 = rsx(5)
      end if


    sinput = "<img src=../COREss20/Empresa/" & sInput2 & " >" & "&nbsp;" & sinput

    SINPUT2 = ""

    IF rs(7) = 1 THEN
       sImagenAL = "<B><FONT FACE=corflaticon COLOR=#008000 SIZE=3>&#0252;</FONT></B>"

       SINPUT2 = "&nbsp;<A href='edgeneralEMPALC.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo  & "' "
       SINPUT2 = SINPUT2 & " TITLE='ESTABLECER ALCANCE PARA EMPRESAS' >" & sImagenAL & "</A>"
    END IF

    Celda = Celda & ofv.GenCelda("","vt valign=top ","","","","","",sInput & SINPUT2,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

   end if

    'Descripcion
    Celda = ofv.GenCelda("","tt","","","","","","Descripcion","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text","edGeneral","80","dt","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput & SINPUT3,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    'Identificacion
    Celda = ofv.GenCelda("","tt","","","","","","Identificacion","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","edGeneral","20","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    'PRODUCTO
    Celda = ofv.GenCelda("","tt","","","","","","PRODUCTO","si")


        sinputK = ""
        sinputt = ""






    if bModificar then
      Strsql = "Select * From TIPOPRODUCTO Order By descri"
      sInput = ofv.GenerarCombo("edGeneral",rs(9).name,rs(9).value,Strsql,cn,0,1,"")
    else
      if cstr(rs(9) & "") <> "" then
        Strsql = "Select descri From TIPOPRODUCTO Where objid = " & rs(9)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
      sInput = ofv.GenerarInput("",sInstrumento,"readonly","","90","dt","")
    end if






    Celda = Celda & ofv.GenCelda("","UT","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    'AGRUPAMIENTO
    Celda = ofv.GenCelda("","tt","","","","","","AGRUPAMIENTO y segmento","si")


        sinputK = "<a style='text-decoration:none;color:#ffffff;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
        sinputK = sinputK  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
        sinputK = sinputK  & "none" & chr(34) & ";' "
        sinputK = sinputK  & "  href='edgeneraladv.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "'  "
        sinputK = sinputK  & "  TITLE='Agregar/Modificar'  ><b>" & sImagenI & "</b></a>"



    Celda = Celda & ofv.GenCelda("","ut","","","","","",sinputK,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


    'Nivel de Confidencialidad
    Celda = ofv.GenCelda("","tt","","","","","","Nivel","si")





    if bModificar then
      sInput = ofv.GenerarCombo("edGeneral",rs(6).name,rs(6).value,Strsqlx,cn2,0,1,"")
    else
      if cstr(rs(6)) <> "" then
         sNiv = rs(6)
         for ii=0 to ubound(CNIV)
             if cstr(CNIV(ii)) = cstr(rs(6)) then
                sNiv = DNIV(ii)
                exit for
             end if
         next
      end if
      sInput = ofv.GenerarInput("",sNiv,"readonly","","90","dt","")
    end if



    Celda = Celda & ofv.GenCelda("","UT","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



sInstrumento = ""
    'Reemplazo
    Celda = ofv.GenCelda("","tt","","","","","","Reemplaza","si")
 '   if bModificar then
 '     Strsql = "Select * From csGenObjetos order by xdescri"
 '     sInput = ofv.GenerarCombo("document.edGeneral",rs(4).name,rs(4).value,Strsql,cn,0,1,"sb")
 '   else
      if rs(4) <> "" then
      Strsql = "Select xdescri From csGenObjetos where codobjeto = " & rs(4)
     '' response.write strsql & "<br>"
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
      sInput = ofv.GenerarInput("",sInstrumento,"readonly","","90","dt","")
  '  end if

  xestad = 1

  Strsql = "SELECT estado FROM OBJPRTY "
  Strsql = Strsql & "where  codobjeto=" & nCodigo
  set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsx.eof then xestad = rsx(0)
  call ofv.cerrarconsulta(rsx)

    if bModificar and cint(xestad) = 1 then
       sImagen = ofv.GenImage("","aaa valign=bottom ","","","Iconos/ICorEsq1.jpg","0","")
       sHref = "EdgeneralBUS.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SLetra=A&Sopcion=A&VASP=CM"
       sLink = ofv.GenLink("","",sHref,"","","",sImagen,"si")
       sinput = sinput & "&nbsp;" & sLink
    end if

    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:9pt;background-color:#f7eedd;} "
    'Encabezado Referencia
    Celda = ofv.GenCelda("","aaa colspan=2","","100%","","","","&nbsp;","si")
    Fila =  Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","tt colspan=2","","100%","","","","Palabras para Busqueda","si")
    Fila =  fila & ofv.GenRow("","","","","","","",Celda,"si")

      nCol = 120

  if len(wdescri) = 0 then wdescri = "Sin Descripcion"

  if not  isnull(rs(8)) then
     wdescri = rs(8)
  end if




    t3b = "<font class=at><textarea class=dt "
    t3b = t3b & "name=" & rs(8).name & " cols=" & nCol & " rows=6 "
    if bModificar then
      t3b = t3b & " onblur=edGeneral.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 950) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = ofv.GenCelda("","abc colspan=2","center","100%","","","",t3b,"si")
    Fila =  fila & ofv.GenRow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objetos&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    Formulario = ofv.GenForm("edGeneral","",sAccion,"","post","",Fila,"si")


  call ofv.cerrarconsulta(rs)




    sHTML = sHTML & "<br><br><center>"
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
    sHTML = sHTML & "<center>"
  end if

  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn)
  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>