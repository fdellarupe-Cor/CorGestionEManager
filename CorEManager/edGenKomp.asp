<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  nCodigo = request.querystring("Codcomp")
  sTipo   = request.querystring("Tipocomp")
  vuelta  = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select descri,archivo From Tipocomp Where tipo = '" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then
      sComplemento = rs(0)
      sImagen = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'sHTML  = sHTML & Tabla & "<br><br><br><center>"
  'sHTML  = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","","","no")


  Strsql = "Select * From Complementos Where codcomp=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof Then


    'Descripcion
    Celda = ofv.GenCelda("","tt","","30%","","","","Descripcion","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text","edGeneral2","80","dt","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","80","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","uti","","70%","","","",sInput,"si")
    Fila  = ofv.GenRow("","","","","25","","",Celda,"si")

    'Identificacion
    Celda = ofv.GenCelda("","tt","","30%","","","","Identificacion","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","edGeneral2","25","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","25","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
    Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarC.asp?ID=" & rs(0) & "&Tabla=Complementos&volver=" & vuelta
      sAccion = sAccion & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
    end if
    xtabla= ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila,"si")
    Formulario = ofv.GenForm("edGeneral2","",sAccion,"","post","",xtabla,"si")


    call ofv.cerrarconsulta(rs)

    Strsql = "SELECT * FROM COMPPRTY Where codcomp=" & nCodigo
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    nID = 0
    sASP = "asp/AgregarC.asp?"
    Valor3 = "0"
    Valor4 = "0"
    if  not rs.eof Then
        nID = rs(0)
        sASP = "asp/ActualizarC.asp?"
        Valor3 = rs(3)
        Valor4 = rs(4)
    end if

        if valor3 = "" then Valor3 = "0" & rs(3)
        if valor4 = "" then Valor4 = "0" & rs(4)



  'Tipo de Generacion
    Celda = ofv.GenCelda("","tt","","30%","","","","Tipo de Generacion","si")
    if bModificar then
        Strsql = "Select * From Tipogen where estado=3 Order By tipogen"
        'sInput = ofv.GenerarCombo("edPropiedades2",rs(3).name,Valor3,Strsql,cn,0,0,"sb")
      Set rsOrigen1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      scomK = ""'"null" & Chr(9)
      scomD = "" '"" & Chr(9)
      SCOMSEP =""

      do while not rsOrigen1.eof
       scomK=scomK & SCOMSEP & rsOrigen1(0)
       scomd=scomd & SCOMSEP & rsOrigen1(0)
       SCOMSEP = chr(9)
       rsOrigen1.movenext
      loop

      sInput =  ofv.gencomboFX("edPropiedades2",rs(3).name,Valor3, scomK, scomD, "")
    else
        sInput = ofv.GenerarInput("",Valor3,"readonly","","60","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","uti","","70%","","","",sInput,"si")
    Fila  = ofv.GenRow("","","","","25","","",Celda,"si")

  'Sistema Generador
    Celda = ofv.GenCelda("","tt","","30%","","","","Sistema Generador","si")
    if bModificar then
        Strsql = "Select * From Sisgen where estado=3 Order By sisgen"
        'sInput = ofv.GenerarCombo("edPropiedades2",rs(4).name,Valor4,Strsql,cn,0,0,"sb")
        Set rsOrigen1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

      scomK = ""'"null" & Chr(9)
      scomD = "" '"" & Chr(9)
      SCOMSEP =""

      do while not rsOrigen1.eof
       scomK=scomK & SCOMSEP & rsOrigen1(0)
       scomd=scomd & SCOMSEP & rsOrigen1(0)
       SCOMSEP = chr(9)
       rsOrigen1.movenext
      loop

      sInput =  ofv.gencomboFX("edPropiedades2",rs(4).name,Valor4, scomK, scomD, "")
    else
        sInput = ofv.GenerarInput("",Valor4,"readonly","","60","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","uti","","70%","","","",sInput,"si")
    Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
     xtabla= ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila,"si")

      sAccion = ""
       if bModificar then
        sAccion = sASP & "ID=" & nID & "&Tabla=Compprty&volver=" & vuelta & "&Codcomp=" & nCodigo
        sAccion = sAccion & "&Tipocomp=" & sTipo
      end if


    Formulario = Formulario & ofv.GenForm("edPropiedades2","",sAccion,"","post","",xtabla,"si")



  'Consultar Componente
    Celda = ofv.GenCelda("","tt","","30%","","","","Visualizar","si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='EditarImagen2.asp?ID=" & nCodigo
    sinput = sinput  & "' title='visualizar Componente' target='_new' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft style=text-align:left;","","","15%","","",sInput,"si")

    Fila  = ofv.GenRow("","","","","25","","",Celda,"si")
    xtabla= ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila,"si")



  end if

    sHTML = sHTML & Formulario & xtabla
    'sHTML = sHTML & ofv.GenFTabla()




' BLOQUE INFO
    sInput  = ofv.GenerarInput("boton","Informacion del documento","submit","","10","bietx","")
    sAccion = "Edgeneralinfo.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&family=K&tipocomp=" & sTipo

    bt7     = ofv.GenForm("bt7f","",sAccion,"","post","",sInput,"si")

    celdas  = ofv.GenCelda("","ut","center","","","","",bt7,"si")
    filas   = ofv.GenRow("","","","","30","","",celdas,"si")

    sHTML   = sHTML & "<br>" & ofv.GenTabla("","aaa","","90%","","0","","2","0","","",filas,"si")

    sHTML = sHTML & "</center>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)



  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>