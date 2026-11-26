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

  set cn = ofv.conectar(ofv.strconn2)

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  sopcion = request.querystring("sopcion")

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

  if sopcion <> "" then

     sevento = request.querystring("sevento")
     sukey = GTkey(nCodigo)

     strsql = "insert into objcontrol  (codobjeto,tipobjeto,frecuencia) values ("
     strsql = strsql & nCodigo & ",'" & stipo & "','" & sevento & "')"
     call ofv.crearconsultaEx(StrSql,cn,1,volver)

     response.redirect "edcontroles.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


  else

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)



  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style={display:none;xheight:350px;overflow:auto;" & styl & "] >"

'-------------------------------------------------


  strsql = "Select Version, estado From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    nestado = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  siconmas = "ICorAcciones.jpg"
  Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sEsquema = rs(1)
     siconmas = rs("archivo")
  end if

  call ofv.cerrarconsulta(rs)

  Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof THEN


  bstitulox = sEsquema & ": " & rs(1) & " - " & rs(3)
  h1 = "<IMG SRC=Iconos\" & siconmas & " border=0>"
  END IF
  call ofv.cerrarconsulta(rs)

  sImageobs = "&nbsp;<font face='corflaticon' color='#990000' size=3 >&#095;</font>"

  Formulario = ""

  Celda      = ofv.GenCelda("","ttt COLSPAN=6 ","","","","","","<big>LISTA DE CONTROLES</BIG>","si")
  Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

  Strsql = "Select * From frecuencia  Order By descri"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs1.eof


   'sicK = "<font face='CorFlaticon' size=4 color='" & session("cmfmenu") & "' >&#0158;</font>&nbsp;"
   sicK = "<i class='las la-plus'></i>"


   if bModificar then
      sAccion = "edcontroles.asp?Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&sopcion=A&sevento=" & rs1(0)

      sinput  = "<a class='nt'  href='"
      sinput  = sinput  & sAccion
      sinput  = sinput  & "' title='Agregar Control para " & rs1(1) & "' >" & sicK & "</a>"
      sicK = sinput

   end if

  Agregar = "&nbsp;"

   Celda = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
   Celda = Celda & ofv.GenCelda("","","","1%","","","",Agregar,"si")
   Celda = Celda & ofv.GenCelda("","uts","","1%","","","",sicK,"si")
   Celda = Celda & ofv.GenCelda("","UT COLSPAN=2 ","","","","",""," <U>" & rs1(1) & "</U>","si")

   Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")



  Strsql = "SELECT Objcontrol.objid, Objcontrol.codsector, Objcontrol.codsector, "
  Strsql = Strsql & "OBJCONTROL.frecuencia, OBJCONTROL.codperfil, OBJCONTROL.referencia, b.codinx1 FROM OBJCONTROL "
  'strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objcontrol.objid = b.codinx1) "
  'strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (Objcontrol.objid = b.codinx1) "
  strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where TIPO='OBS' AND codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on ( cast( Objcontrol.objid as nvarchar)= b.codinx1) "
  Strsql = Strsql & "Where Objcontrol.codobjeto = " & nCodigo & " and OBJCONTROL.frecuencia = '" & rs1(0) & "' "
  Strsql = Strsql & " order by Objcontrol.objid"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    do until rs.eof

       'sicK = "<font face='CorFlaticon' size=3 color=#993333 >2</font>"
	   sicK = "<i class='las la-minus' style=font-size:12pt;color:" & session("cmfmenu") & " ></i>"


      Celda = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
      Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")



      sSector = "S/D "
        Strsql = "Select descri From Eco, Objsector where "
        Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.codsector = " & rs(1)
        set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsSector.eof then sSector = rsSector(0)
        call ofv.cerrarconsulta(rsSector)

      sPerfil = "S/D "
      if rs(4) <> "" then
        Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(4) & "'"
        set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsPerfil.eof then sPerfil = rsPerfil(0)
        call ofv.cerrarconsulta(rsPerfil)
      end if

      sfrec = "S/D"
      sfrecn = "S/D"
      if not isnull(rs(3))  then
        Strsql = "Select descri From frecuencia Where codfrec = '" & rs(3) & "' "
        set rsfrec = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsfrec.eof then
           sfrec = "<I>(" & rsfrec(0) & ")</I>"
           sfrecn = "(" & rsfrec(0) & ")"
        end if
        call ofv.cerrarconsulta(rsfrec)
      end if

      Celdax = " "  & sSector & " - " &  sPerfil & " - " & sfrec
      Celdan =   sSector & " - " &  sPerfil & " - " & sfrecn


    'Eliminar
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objcontrol&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      Celdax1 = ofv.BotonEliminar(ofv.convertircar(Celdan," ","_"),sHref)
   '' celdax1 = replace(celdax1,"<TD","<TD style={border-style:outset;border-width:1;} ")
    else
      Celdax1 = ofv.GenCelda("","","","","","","","","si")
    end if
    if sHref <> "" then
      Celdax1 = ofv.BotonEliminar("Eliminacion",sHref)
      Celdax1 = replace(Celdax1,"<TD ","<TD title='Eliminacion' ")
    else
      Celdax1 = ofv.GenCelda("","","","","","","","","si")
    end if

      Celda = Celda & Celdax1


      sinput = "<a class=NT href=" & chr(34) & "edcontrolDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sinput = sinput & "&XEVEN=" & rs(0) & " "
      sinput = sinput &  chr(34) & " title=" &  chr(34) & "ver controles de " & celdan &  chr(34) &  " >" & Celdax & "</a>"

    Celda = Celda & ofv.GenCelda("","utis","","1%","","","",sicK,"si")
    Celda = Celda & ofv.GenCelda("","UT","","96%","","","",sinput,"si")

     ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")


      if (Nestado=5  or Nestado=6) and (rs(6) & "") <> "" and bModificar then

               stitle = "Hay observaciones para este grupo."
      	       sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & sImageobs & "</div>"
      	       Celdaobs  = ofv.GenCelda("","ut title='" & stitle & "'  valign=top ","","","1","","",sInput,"si")
      end if


      Celda = Celda & Celdaobs

    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")


    rs.movenext
    loop
    call ofv.cerrarconsulta(rs)




   rs1.movenext
   loop

 ' sHTML = sHTML & "<CENTER>"
  sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","-1","2","","",Formulario,"si")
'  sHTML = sHTML & "</CENTER>"

  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") ' ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if








'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no","no","no",Listar)


  call ofv.cerrarconsulta(rs1)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

 end if

end if

response.write sHTML


%>