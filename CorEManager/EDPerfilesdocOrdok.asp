<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sOpc    = request.querystring("opc")
  vuelta  = ofv.vueltaasp2("../")

  welem = request.querystring("welem")

  if len(welem) = 0 then welem = 0

     sfuncion = ""

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'Encabezados
  Formulario = ""


  dim stabla(15)
  dim stablaE(15)
  dim stablaS(15)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar   = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir  = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0

  Strsql = "Select * From Objperfildoc Where codobjeto =" & nCodigo & " and tipobjeto ='" & stipo & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if  not rs.eof then

      fila   = ""

      Celda  = ofv.GenCelda("","tt","","10%","","","","Perfil: ","si")
      sInput = "&nbsp;&nbsp;" & rs(0) & "&nbsp;-&nbsp;" & rs(1)

      Celda  = celda & ofv.GenCelda("","ttt","","80%","","","",sInput,"si")
      sInput = ofv.GenerarInput("bt3","volver","submit","","10","bt","")
      Celda  = Celda & ofv.gencelda("","ut","right","10%","","","",sInput,"si")
      Fila   = fila & ofv.genrow("","","","","","","",Celda,"si")
      Fila   = ofv.GenTabla("","","","70%","","0","","-2","-1","","",Fila,"si")
      sHTML  = sHTML & OFV.GenForm("defa2", "", "EDperfilesdocProp.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&opc=SI", "_self", "post", "", Fila, "si")

      Formulario = ""
      fila   = ""

      sInput = "&nbsp;Orden&nbsp;de&nbsp;Solapas"

      Celda = ofv.GenCelda("","ttr colspan=3 ","","","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")


      Formulario = Fila


    for xx = 2 to 16
        stabla (xx - 1) = rs(xx + 30)
    next


    yy = 0
    for xx = 1 to 15
        if cint(stabla (xx)) < 99 then
           yy = yy + 1
           stablaE (stabla (xx)) = xx
        end if
    next
    tyy = yy

    if  welem = 15 then welem = 0

    if  welem > 0 then
  '      if  cint(stabla(welem - 1)) < 3 then welem = 0
    end if

    if  welem > 0 then


       welemix  = stabla(welem - 1)
       welemixa = cint(welemix) - 1
       welemixh = cint(welemix) - 2

       for xx = 1 to welemixh

           stablaS(xx) =  stablaE(xx)

       next

       stablaS(welemix) =  stablaE(welemixa)
       stablaS(welemixa) =  welem - 1

       for xx = (cint(welemix) + 1) to cint(tyy)

           stablaS(xx) =  stablaE(xx)

       next

       Strsql = "UPDATE OBJPERFILdoc SET "

       for xx = 1 to tyy

           Strsql = strsql & rs(stablaS(xx) + 1 + 30).name & " = " & xx
           if xx <> tyy then Strsql = strsql & ", "

       next

       Strsql = strsql & " where codobjeto =" & nCodigo & " and tipobjeto ='" & stipo & "' "
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)


    else

       for xx = 1 to tyy

           stablaS(xx) =  stablaE(xx)

       next

    end if


    xx = 0

    for yy = 1 to tyy

      fila   = ""
      xx     = xx + 1

      sInput = rs(stablaS(yy) + 1 + 15)

      if  rs(stablaS(yy) + 1) then
          Celda  = ofv.GenCelda("","vt style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00bb00;' ","right","","","","",xx & ".&nbsp;&nbsp;","si")
          clatr  = " bgcolor=#f5f5f5 "
          if  xx = 1 or xx = 82 then

              Celda  = celda & ofv.GenCelda("","vti style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00bb00;' ","","","","","",sInput,"si")
              Celda  = Celda & ofv.gencelda("","ut style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00bb00;' ","center","","","","","&nbsp;","si")

          else

              Celda  = celda & ofv.GenCelda("","vti style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00bb00;' ","","","","","",sInput,"si")
              sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
              sinput = "<a title='Subir Solapa.' class='ft1n' " & sFuncAnc & " href=EDperfilesdocOrd.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo
              sinput = sinput  & "&Welem=" & (stablaS(yy) + 1) & " ><font class=btic>5</font></a>"
              Celda  = Celda & ofv.gencelda("","btbn style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00bb00;' ","center","","","","",sInput,"si")

          end if
          Fila = fila & ofv.genrow("","aaa " & clatr,"","","25","","",Celda,"si")

      else
          Celda  = ofv.GenCelda("","ut","right","","","","",xx & ".&nbsp;&nbsp;","si")
          clatr  = ""
          Celda  = celda & ofv.GenCelda("","uti ","","","","","",sInput,"si")
          Celda  = Celda & ofv.gencelda("","ut","","","","","","&nbsp;","si")
          Fila = fila & ofv.genrow("","aaa " & clatr,"","","","","",Celda,"si")

      end if

      Formulario = Formulario & Fila

     next

  end if

          clatr  = ""
          Celda  = ofv.GenCelda("","ut","right","10%","","","","&nbsp;","si")
          Celda  = celda & ofv.GenCelda("","uti ","","85%","","","","&nbsp;","si")
          Celda  = Celda & ofv.gencelda("","ut","","","5%","","","&nbsp;","si")
          Fila   = ofv.genrow("","aaa " & clatr,"","","","","",Celda,"si")

  Formulario = Formulario & Fila

  sHTML = sHTML & ofv.GenTabla("","","","70% ","","0","","-2","0","","",Formulario,"si")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>