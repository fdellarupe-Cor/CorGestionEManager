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
  if  session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sOpc    = request.querystring("opc")
  vuelta  = ofv.vueltaasp2("../") & "?opc=SI"

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'Encabezados
  Formulario = ""



  dim stabla(12)


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

      fila = ""


      Celda  = ofv.GenCelda("","ttt ","","","","","",sOpc  & " Perfil&nbsp;Default: ","si")
      sInput = ofv.combosino("defa1","0","dato")
      Celda  = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
      Fila   = ofv.genrow("","","","","","","",Celda,"si")
      Formulario = ofv.GenTabla("","","","","","0","","-2","-1","","",Fila,"si")
      sHTML  = sHTML & OFV.GenForm("defa1", "", "EDperfilesdocDef.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo  , "_self", "post", "", Formulario, "si")

      Fila   = ""

      Celda  = ofv.GenCelda("","tt","","10%","","","","Perfil: ","si")
      sInput = "&nbsp;&nbsp;" & rs(0) & "&nbsp;-&nbsp;" & rs(1)

      Celda  = celda & ofv.GenCelda("","ttt","","80%","","","",sInput,"si")
      sInput = ofv.GenerarInput("bt3","Orden","submit","","10","bt","")
      Celda  = Celda & ofv.gencelda("","ut","right","10%","","","",sInput,"si")
      Fila   = fila & ofv.genrow("","","","","","","",Celda,"si")
      Fila   = ofv.GenTabla("","","","95%","","0","","-2","-1","","",Fila,"si")
      sHTML  = sHTML & OFV.GenForm("defa2", "", "EDperfilesdocOrd.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&Welem=0", "_self", "post", "", Fila, "si")


      Formulario = ""
      fila  = ""

      Celda = ofv.GenCelda("","ttr ","","","","","","SOLAPA","si")
      Celda = celda & ofv.GenCelda("","ttr","","","","","","TITULO","si")
      Celda = celda & ofv.GenCelda("","ttr","","","","","","ACTIVA","si")
      Celda = celda & ofv.GenCelda("","ttr","","","","","","DESCRIPCION","si")
      Fila  = fila & ofv.genrow("","","","","20","","",Celda,"si")
    '''  Formulario = Fila

dim atab(15)

    for xx = 1 to 15
           atab(xx) = 99
    next

    for xx = 2 to 16

        if cint(rs(xx + 30)) < 99 then

           atab(rs(xx + 30)) = xx

        end if


    next

           sAccion = ""
           if bModificar then
              sAccion = "asp/ActualizarEN.asp?ID=" & rs("objid") & "&Tabla=Objperfildoc&volver=" & vuelta
              sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
           end if


    for yy = 1 to 15

        if atab(yy) = 99 then exit for
           xx = atab(yy)
       '''    fila = ""
           sForm = "form" & posi

           if rs(xx)  then
              clase = "ut style='border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#00bb00;' "
              clatr = " bgcolor=#f5f5f5 "
           else
              clase = "ut"
              clatr = ""
           end if

           sInput = rs(xx + 30) & ".&nbsp;&nbsp;"
           Celda = ofv.GenCelda("",clase,"right","1%","","","",sInput,"si")

           sInput = ofv.GenerarInput(rs(xx + 15).name,rs(xx + 15),"text",sForm & "X1","120","dt","")
           sInput = ofv.genform(sForm & "X1","",sAccion,"","post","",sInput,"si")
           Celda = celda & ofv.GenCelda("",clase,"","40%","","","",sInput,"si")

           if yy = 1 then
              sInput = "<b>SI</b>"
           else
              sInput = ofv.combosino(sForm & "X2",rs(xx),rs(xx).name)
           end if
           sInput = ofv.genform(sForm & "X2","",sAccion,"","post","",sInput,"si")

           Celda = celda & ofv.GenCelda("",clase,"center","5%","","","",sInput,"si")



           sInput = "&nbsp;"

           Strsql = "Select descri From Objperfildsc where dato = '" & rs(xx).name & "' "
         '  response.write strsql
           set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not rs2.eof then
              sInput = rs2(0)
           end if
           call ofv.cerrarconsulta(rs2)

           Celda = celda & ofv.GenCelda("","ut ","","40%","","","","<i>" & sInput & "</i>","si")

           Fila = fila & ofv.genrow("","aaa " & clatr,"","","","","",Celda,"si")



           posi = posi + 1



    next
    sHTML = sHTML & ofv.GenTabla("","","","95% ","","0","","-2","-1","","",fila,"si")

  else

     fila = ""

      Celda = ofv.GenCelda("","ttt colspan=2  ","","","","","",sOpc  & " Perfil Default: ","si")
      sInput = ofv.combosino("defa","1","dato")
      sInput = OFV.GenForm("defa", "", "EDperfilesdocDef.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo  , "_self", "post", "", sInput, "si")

      Celda = celda & ofv.GenCelda("","ut colspan=6 ","","90%","","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")

      Celda = ""
      for ii=1 to 7
          Celda = celda & ofv.GenCelda("","aaa ","","10%","","","","&nbsp;","si")
      next
      sInput = OFV.GenerarInput("boton", "", "hidden", "", "20", "it", "")
      sInput = OFV.GenForm("nulo", "", "BarraEN.asp?Codigo=" & nCodigo & "&tipo=" & stipo & "&Funcion=E  " , "Barra", "post", "", sInput, "si")
      Celda  = Celda & OFV.gencelda("", " ", "center", "", "", "", "", sInput & "&nbsp;", "si")

      Fila   = fila & ofv.genrow("","","","","","","",Celda,"si")
      Formulario = Fila
      sHTML = sHTML & ofv.GenTabla("","","","95% ","","0","","-2","-1","","",Formulario,"si")

  end if

  sHTML1 = ""

  if  sopc = "xSI" then

            sHTML1 = "<script language=JavaScript>" & Chr(13)
            sHTML1 = sHTML1 & " window.top.rbarr(); " & Chr(13)
            sHTML1 = sHTML1 & "</script>" & Chr(13)

  end if



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
 sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & sHTML1
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>