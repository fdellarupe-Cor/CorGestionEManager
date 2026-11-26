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
  sOPCION = request.querystring("OPCION")
  vuelta  = ofv.vueltaasp2("../")

  set  cn = ofv.conectar(ofv.strconn2)

  sHTML   = ofv.FormHeader(session("FormN"))
 shtml = shtml & CorPnlFnc("N","")

'----------------------------------------------


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

 IF SOPCION = "SI" THEN

  Tabla= "OBJMEMPALC"

  XINCD   = request.querystring("FILTRO")

  strsql  = "Select * From OBJMEMPALC Where codobjeto = " & nCodigo
  set rs  = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sForm   = "Agregarb"

  sInput  = "AGREGAR DATOS"
  Celda   = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
  Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  Celda   = ofv.gencelda("","tt","","20%","","","","EMPRESA","si")

  'Combo EMPRESA
  Strsql  = "Select * From CLIENTEPLAN WHERE ID NOT IN (" & XINCD & ") "
  STRSQL  = STRSQL & " order by CLIENTE "
  sInput  = ofv.GenerarCombo2(rs(3).name,"",Strsql,cn,0,1,"Agregar")
  Celda   = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

  Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sFuncion = ""
  sAccion  = "asp/AgregarEN.asp?Tabla=" & Tabla
  sAccion  = sAccion & "&volver=" & vuelta & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

  tabla=ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")

  'ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")
  Formulario = ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,tabla,"no")

  sHTML = sHTML &Formulario ' ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botones Aceptar y Cancelar
  sAccion = "edGENERALEMPALC.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML   = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  Agregar = ""
  Buscar  = ""
  Celda   = ""
  sHTML = sHTML & "</div>"



    Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

    sHTML = sHTML & CorPnlIndx("N","")

    sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  call ofv.cerrarconsulta(rs)


 ELSE


  'Encabezados

  Formulario = ""

  Formulario = Formulario & "<tr ><td class=ttt "
  Formulario = Formulario & " COLSPAN=3 VALIGN=BOTTOM width=30% ><big>EMPRESAS ALCANZADAS</big></td></tr>"

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si") & "<br>"



  celdas = ofv.GenCelda("","ttr","","2%","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","ttr","","98%","","","","EMPRESA","si")
  Formulario = ofv.GenRow("","","","","20","","",celdas,"si")


 '' Formulario = Formulario & ofv.Encabezados("EMPRESA")

  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"

  set cn = ofv.conectar(ofv.strconn2)

  XINCD = ""
  XINCS = ""

  Strsql = "Select * From OBJMEMPALC Where codobjeto=" & nCodigo & " Order By CODEMPRE "
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF  NOT RS.EOF THEN
  do  until rs.eof
      sForm = "edP" & rs(0)

      XINCD = XINCD & XINCS & RS(3)
      XINCS = ","

    'Eliminar
    if  bModificar then
        sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=OBJMEMPALC&volver=" & vuelta
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if

    if  sHref <> "" then
        Celda = ofv.BotonEliminar(rs(0),sHref)
    else
        Celda = ofv.GenCelda("","","","","","","","","si")
    end if


 '   sInput = rs(3) & ".&nbsp;"
 '   Celda  = Celda & ofv.GenCelda("","ttt","right","2%","","","",sInput,"si")

    'EMPRESA

      sInput = "S/D"
      if  cstr(rs(3)&"") <> "" then

          Strsql = "Select CLIENTE From CLIENTEPLAN where ID = " & rs(3)
          set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsInstrumento.eof then sInput = rsInstrumento(0)
          call ofv.cerrarconsulta(rsInstrumento)

      end if


    sInput = "&nbsp;&nbsp;" & sInput
    Celda  = Celda & ofv.GenCelda("","ttt","","80%","","","",sInput,"si")



    Fila   = ofv.GenRow("","","","","","","",Celda,"si")

    if  bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJMEMPALC&volver=" & vuelta
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
    else
        Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
    end if

    rs.movenext
  loop

 ELSE

  Formulario = Formulario & "<tr height=30 ><td class=vt >&nbsp;</TD><td WIDTH=30% class=vt >&nbsp;</TD><td WIDTH=70% class=vt >&nbsp;</TD> "
  Formulario = Formulario & "</tr>"

 END IF



  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

 ' response.write XINCD

''  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",XINCD,"si")

    Agregar = ""

  IF XINCD <> "" THEN
     XINCD = XINCD & ",1"
     XINCS = XINCD
  ''   Strsql = "Select * From OBJMEMPALC Where codobjeto=" & nCodigo & " AND CODEMPRE NOT IN (" & XINCD & ") "
     Strsql = "Select id From CLIENTEPLAN Where  id NOT IN (" & XINCD & ") "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF NOT RS.EOF THEN
        XINCD = "OK"
     ELSE
        XINCD = "NO"
     END IF
  ELSE
     XINCD = "OK"
     XINCS = "1"
  END IF

  'Botonera
  if bModificar  AND  XINCD = "OK" then
    sAccion = "edGENERALEMPALC.asp?OPCION=SI&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo & "&FILTRO=" & XINCS
    Agregar = ofv.BotonAgregar(sAccion)

  end if


  'Boton Volver
  sAccion   = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  sRetornar = ofv.BotonVolver(sAccion)

  'Botonera
  Buscar  = ""


   sHTML = sHTML & "</span></div>"



  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  sHTML = sHTML & CorPnlIndx("N","")
  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,sRetornar)


  ''  sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Cons,"")

  call ofv.cerrarconsulta(rs)

 END IF

  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>

