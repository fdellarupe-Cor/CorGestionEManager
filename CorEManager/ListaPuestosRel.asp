<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  ID=request.querystring("ID")

  CantRegMoverC=request.querystring("CantRegMoverC")

  if ID = "" then
          ID = session("elID")
       else
          session("elID") =cdbl(ID)
     end if

  IF session("Regisvuel")="" THEN
   session("Regisvuel")=CantRegMoverC
  ELSE
   CantRegMoverC=session("Regisvuel")
  END IF
  sUN = request.querystring("UN")
  sOpcion = request.querystring("Opcion")

  IF sUN <> "" THEN
     SESSION("SUN") = SUN
  ELSE
    SUN = SESSION("SUN")
  END IF

  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  if sOpcion = "SI" then

   usrexiste=false

   USRID=REQUEST.FORM("XXX") & ""
   strsql="select * from PUESTOREL where cod='" & USRID & "'"
  ' response.write strsql
   set rsusr= ofv.crearconsultaEx(StrSql,cn,1,volver)
   if  rsusr.eof then

    strsql = "Insert into PUESTOREL "
    strsql = strsql & " (IDPST,tipo,cod)"
    strsql = strsql & " Values ("&ID&",'USUARIO','"& usrid &"')"
    call ofv.crearconsultaEx(StrSql,cn,1,volver)
   else
    usrexiste=true
   end if
  end if

      strsql =  "SELECT  DESCRI FROM LISTADEPUESTOS WHERE OBJID=0" & ID

	  SET RP= ofv.crearconsultaEx(StrSql,cn,1,volver)
	  IF NOT RP.EOF THEN
	   DESCRIP=RP(0)
	  ELSE
	   DESCRIP="&nbsp;"
	  END IF

  call ofv.cerrarconsulta(rp)


  'sImagen4 = "<B><FONT FACE=CorFlaticon COLOR=#ffffff SIZE=2>4</FONT></B>"
  sImagen4 = "<B><i class='las la-angle-right' STYLE=COLOR:#ffffff ></i></B>"
  sImagenP = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  Celda = ofv.gencelda("","ttt","","","","","","<BIG>" & sImagen4 & " Puesto : " & DESCRIP & "</BIG>","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","aaa","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"

  'Encabezados


 '  Formulario = ofv.Encabezados("Id","Descripcion del Puesto","segmento","Tipo","codigo","#","DOTACION","continuador","GRAFICA","Propiedades")

  Formulario = ofv.Encabezados("TIPO","CODIGO","DESCRI")

  Strsql = "Select * From PUESTOREL WHERE IDPST=0" & ID &  " Order By TIPO,cod"
  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=PUESTOREL&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    sTIPO = rs(3) & ""

  '  if ssino = "" then ssino = "R"
    scomK = "USUARIO" & chr(9)' & "PUESTO"
    scomD = "USUARIO" & chr(9)' & "PUESTO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(3).name,sTIPO,scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
    end if

    Celda = Celda & ofv.GenCelda("","vt ","","10%","","","",rs(3),"si")

    if bModificar then
       sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"80","dt","")
    else
       sInput = rs(1)
    end if


    Celda = Celda & ofv.GenCelda("","vt ","","15%","","","",sInput,"si")


      if bModificar AND RS("TIPO") <> "USUARIO" then
	 sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"30","dt","si")
       else

         IF RS("TIPO")="USUARIO" THEN
         XX=GETINFOUSER(rs(1)&"")
         SINPUT=session("SGU_nombre")
         ELSE
	 sInput = rs(2)
	 END  IF
      end if

     Celda = Celda & ofv.GenCelda("","vt","","75%","","","",sInput,"si")


          'sImagenP = "&#0232;"
    '        sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    '        sinput = "<a class='ft1n' " & sFuncAnc & " href='BSQUSER.asp?NID=" & rs(0) & "&BSQ=PORNOMBRE" '& CantRegMoverx
    '     '   sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    '        sinput = sinput  & "' title='BUSCAR USUARIOS' >" & sImagenP & "</a>"


   ' Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=PUESTOREL&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","-1","1","","",Formulario,"si")


  retornar = ""



  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "BSQUSER.asp?BSQ=PORNOMBRE&CantRegMover=" & CantRegMoverx & "&Opcion=SI"  '"listapuestosREL.asp?CantRegMover=" & CantRegMoverx & "&Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


 ' str1 = "Select descri  From PUESTOREL Order  By TIPO,descri"  ' orden pagina
 ' str2 = "Select descri  From PUESTOREL Order  By TIPO,descri"  ' orden alfabetico

  Buscar =""' ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

     Listar  = ""



     'Listar Organigrama

 '    xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Listar el organigrama", "NO", "Organigrama")
  '   sAccion = "listadepuestosOrganigrama.asp"
   '  Listar  = "" 'ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")



  'consulta

 ' sAccion = "listadepuestosCons.asp?" & request.querystring

  Celda = ""'ofv.BotonConsultar(sAccion)

  sAccion = "LISTADEPUESTOS.asp?CantRegMover=" & CantRegMoverC

  retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Listar,celda & retornar)

  call ofv.cerrarconsulta(rs)

  shtml=SHTML & 	ValidarInput()

   if usrexiste then

    XX=getInfouserPst(USRID,0,"USUARIO")
     sql="SELECT b.descri from  puestorel a inner join ListadePuestos b on (b.objid=a.IDPST) where a.cod='" & USRID & "'"

	 set rsf = ofv.crearconsultaEx(sql,cn,1,parametros)
	 spuestof=""
	 if not rsf.eof then spuestof= rsf(0)
     call ofv.cerrarconsulta(rsf)

    shtml=shtml & "<script>alert('El usuario/a "& session("SGU_nombre") & " ya esta asignado al puesto " & spuestof & "');</script>"

   end if

  sHTML = sHTML & "</body></html>"
  call ofv.cerrarconn(cn)
end if

Response.Write sHTML

%>