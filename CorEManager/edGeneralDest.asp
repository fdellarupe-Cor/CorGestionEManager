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
  bModificaresp = ofv.AModificar
  if session("Nomodificar") = "S" then

     if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False

  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sOPCION = request.querystring("OPCION")
  vuelta  = ofv.vueltaasp2("../")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'Encabezados
 ' Formulario = ofv.Encabezados("AGRUPAMIENTO","SEGMENTO")

  celdas = ofv.GenCelda("","ttr","","2%","","","","&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","ttr","","49%","","","","Destinos","si")
  Formulariocabe = ofv.GenRow("","","","","20","","",celdas,"si")

  sImagenI   = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"

  set cn = ofv.conectar(ofv.strconn2)

  IF  SOPCION = "SI" THEN
      Strsql = "insert into envioformdet2  (idlote,iddestino,codobjeto,tipobjeto) values ("
      Strsql = Strsql & "'0','0','" & nCodigo & "','" & sTipo & "')"
      call ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF

  Strsql = "Select * From envioformdet2 Where idlote=0 and codobjeto=" & nCodigo & " and tipobjeto='" & stipo & "' Order By OBJID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do  until rs.eof
      sForm = "edP" & rs(0)

        if  bModificar then
      sAccion    = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=envioformdet2&volver=" & vuelta
      sAccion    = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      end if

    'Eliminar
     if  bModificar then
         sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=envioformdet2&volver=" & vuelta
         sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
     end if

     if  sHref <> "" then
         Celda = ofv.BotonEliminar(rs(0),sHref)
     else
         Celda = ofv.GenCelda("","","","","","","","","si")
     end if

    'AGRUPAMIENTO

    Strsql = "Select id,destino From destinosurl where estado=3 order by id"
  ''  sInput = ofv.GenerarCombo(sForm,rs(5).name,rs(5).value,Strsql,cn,0,1,"")
      Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
       scomK = ""'"null" & Chr(9)

      scomD = "" '"" & Chr(9)
      SCOMSEP =""
      do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
       scomd=scomd & SCOMSEP & rsperfil(1)
       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop

      sInput =  ofv.gencomboFX(sForm & "w1", rs(2).name, rs(2).value, scomK, scomD, "")
      sInput = ofv.GenForm(sForm & "w1","",sAccion,"","post","",sInput,"si")
    'Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",sinput,"si")
    'Celda  = Celda & ofv.GenCelda("","VT","","40%","","","",formulario,"si")
    Celda  = Celda & ofv.GenCelda("","VT","","40%","","","",sInput,"si")


     Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")


    rs.movenext
  loop

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulariocabe & Formulario,"si")
 ''  tabla= ofv.GenTabla("","","","70%","","0","","0","0","","",Formulariocabe & fila,"si")
''  sHTML = sHTML & tabla
 '' sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","-2","-1","","",Formulario,"si")
 '' sHTML = sHTML & formulario

    Agregar = ""

  'Botonera
  if  bModificar then
      sAccion = "edgeneraldest.asp?OPCION=SI&Codobjeto=" & nCodigo
      sAccion = sAccion & "&Tipobjeto=" & sTipo
      Agregar = ofv.BotonAgregar(sAccion)
  end if

  sAccion = "edGENERAL.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  sRetornar = ofv.BotonVolver(sAccion)



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML   = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,sRetornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>

