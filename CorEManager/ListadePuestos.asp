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
  session("Regisvuel")=""
  sUN = request.querystring("UN")
  sOpcion = request.querystring("Opcion")

  IF sUN <> "" THEN

     SESSION("SUN") = SUN

  ELSE

    SUN = SESSION("SUN")
    SESSION("SUN")=SESSION("CLIEMP")

  END IF
'response.write SESSION("CLIEMP")
  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

'response.write SESSION("CLIEMP")
  if sOpcion = "SI" then


        strsql = "Insert into ListadePuestos "
        strsql = strsql & " (descri,nivelj,codigo,secuen,segmento,DOTACION,CONTINUADOR,sector,posy,GRAFICA,niveljreal,site,tipo)"
        strsql = strsql & " Values ("
        strsql = strsql &  "'Nuevo Puesto','1','','0','0','1','NO'"
        strsql = strsql & ",0,0,'N','1','" & SESSION("un") & "','R')"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if




    sImagen4 = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>4</FONT></B>"
    sImagenP = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  'Celda = ofv.gencelda("","ttt","","","","","","<BIG>" & sImagen4 & " " & SNOMEMPRE & "</BIG>","si")
  'Fila = ofv.genrow("","","","","","","",Celda,"si")
  'sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"

  'Encabezados


 '  Formulario = ofv.Encabezados("Id","Descripcion Puesto","segmento","Tipo","codigo","#","DOTACION","continuador","GRAFICA","Propiedades")
  Formulario = ofv.Encabezados("Id","Descripcion&nbsp;del&nbsp;Puesto","Pos.Grafico","Perfil&nbsp;del&nbsp;puesto","segmento","Tipo","#","Dotacion","Pertenece&nbsp;A","Grafica","Propiedades","CodRel")
  Strsql = "Select * From listadepuestos where site='" & SESSION("un") & "' Order By continuador,descri"
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=listadepuestos&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name &" maxlength=70",rs(1),"text",sForm,"50","dtbs","")
    Celda  = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    sql3   = "Select * From NIVELpuesto Order By nivelp"
    sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2).value,sql3,cn,1,2,"")
    celda  = celda & ofv.GenCelda("","vti ","","","","","",sInput ,"si")

    sql3   = "Select * From NIVELpuesto Order By nivelp"
        sInput = ofv.GenerarCombo(sForm,rs(12).name,rs(12).value,sql3,cn,1,2,"")
    celda  = celda & ofv.GenCelda("","vti ","","","","","",sInput ,"si")


    sql3   = "Select * From SEGMENTO Order By descri"
    sInput = ofv.GenerarCombo(sForm,rs(6).name,rs(6).value,sql3,cn,0,1,"")
    celda  = celda & ofv.GenCelda("","vti ","","","","","",sInput,"si")


    ssino = rs(4) & ""

    if ssino = "" then ssino = "R"
    scomK = "R" & chr(9) & "G" & chr(9) & "S"
    scomD = "REAL" & chr(9) & "GENERICO" & chr(9) & "STAFF"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(4).name,ssino,scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
    end if

    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Descripcion
  '  sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"10","dtbs","")
  '  Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    if bModificar then
       sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"20","dtn ","SI")
    else
       sInput = rs(5)
    end if


     Celda = Celda & ofv.GenCelda("","vt","","","","","",sInput,"si")




      'scomK = "A" & chr(9) & "B" &  chr(9) & "C" & chr(9) & "D"
	  'scomD = "A" & chr(9) & "B"  & chr(9) & "C" & chr(9) & "D"

     'ssino = rs(7) & ""

	 '   if bModificar then
	 '       sInput = ofv.gencomboFX(sForm,rs(7).name,ssino,scomK,scomD," ")
	 '   else
	 '       sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
     '  end if

      if bModificar then
	        sInput = ofv.GenerarInput(rs(7).name,rs(7),"text",sForm,"20","dtn ","SI")
	     else
	        sInput = rs(7)
      end if

     Celda = Celda & ofv.GenCelda("","vt  ","","","","","",sInput,"si")

  '   scomK = "SI" & chr(9) & "NO"
	' scomD = "SI" & chr(9) & "NO"

	' ssino = rs(8) & ""

	'  if bModificar then
	'       sInput = ofv.gencomboFX(sForm,rs(8).name,ssino,scomK,scomD," ")
	'   else
	' 	    sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
   '    end if
   '  Celda = Celda & ofv.GenCelda("","vt  ","","","","","",sInput,"si")



         scomK = "0" & chr(9)
	     scomD = "SIN DEFINIR" & chr(9)

	     strsql="select * from eco where eco.estado=3 and ecoun='" &  SESSION("un") & "' order by eco.DESCRI"
'response.write strsql
	     set rseco = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	     do while not rseco.eof
	     scomK=scomK & rseco(0) & chr(9)
	     scomd=scomd & rseco(1) & chr(9)
	     rseco.movenext
	     loop

	      ssino = rs(9) & ""

	     if ssino = "" then ssino = "0"
	     if bModificar then
	 	        sInput = ofv.gencomboFX(sForm,rs(9).name,ssino,scomK,scomD," ")
	 	    else
	 	        sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
	     end if

     Celda = Celda & ofv.GenCelda("","vt  ","","","","","",sInput,"si")



  	'sql3   = "Select * From NIVELpuesto Order By nivelp"
    '  sInput = ofv.GenerarCombo(sForm,rs("posy").name,rs("posy").value,sql3,cn,1,2,"")
    '  celda  = celda & ofv.GenCelda("","vti ","","","","","",sInput,"si")


      'if bModificar then
	 '	        sInput = ofv.GenerarInput(rs(10).name,rs(10),"text",sForm,"3","dtn","SI")
	 '	     else
	 '	        sInput = rs(7)
      'end if

   '  Celda = Celda & ofv.GenCelda("","vt  ","","","","","",sInput,"si")
    'Propiedades
        ssino = rs(11) & ""
       if ssino = "" then ssino = "N"
    '
        scomK = "S" & chr(9) & "N"
        scomD = "SI" & chr(9) & "NO"
    '
    '
        if bModificar then
            sInput = ofv.gencomboFX(sForm,rs(11).name,ssino,scomK,scomD," ")
        else
            sInput = ofv.GenerarInput("",ssino,"readonly","","10","dt","")
        end if
    '
       ''''' sInput = ofv.combosino(sForm,ssino,rs(2).name)
       Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='ListaPuestosFUN.asp?PUESTO=" & rs(0) & "&CantRegMover=" & CantRegMoverx
    sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"


    Celda = Celda & ofv.gencelda("","ft","center","","5%","","",sInput,"si")



    'sImagenP = "&#0232;"
        'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
        sinput = "<a href='ListaPuestosrel.asp?ID=" & rs(0) & "&CantRegMoverC=" & CantRegMoverx
        sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
        sinput = sinput  & "' title='Relaciones de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"


    Celda = Celda & ofv.gencelda("","ft","center","","5%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=listadepuestos&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  'sHTML = sHTML & ofv.GenTabla("","","","800px","","0","","-1","1","","",Formulario,"si")
  
    ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")


  retornar = ""



  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "listadepuestos.asp?CantRegMover=" & CantRegMoverx & "&Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri  From listadepuestos where site='" & SESSION("un") & "' Order  By continuador,descri"  ' orden pagina
  str2 = str1 ' "Select descri  From listadepuestos Order  By continuador,descri"  ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

     Listar  = ""


     'Listar Organigrama

     xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Listar el organigrama", "NO", "Organigrama")
     sAccion = "listadepuestosOrganigrama.asp"
     Listar  = "" 'ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")



  'consulta

  sAccion = "listadepuestosCons.asp?" & request.querystring
  Celda = ""'ofv.BotonConsultar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Listar,celda & retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>