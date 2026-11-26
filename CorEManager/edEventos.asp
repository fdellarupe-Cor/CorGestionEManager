<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
CONT = 0
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")



  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  sHTML  = ofv.FormHeader(session("FormN"))
  vuelta = ofv.vueltaasp2("../")

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

%>
<!-- #INCLUDE FILE="IncFile/CorVinculadorPL.inc" -->
<%


  IF WKSIGUE = "S" THEN

  'Encabezados

  if wkaplic = "PL" then

     Formulario = ofv.Encabezados("Descripcion","Secuencia")

  else
		If session("APLMDCLI") = "BGAL" then
			Formulario = ofv.Encabezados("Descripcion","Secuencia","Proceso")
			Else
				Formulario = ofv.Encabezados("Descripcion","Secuencia")
		End If
  end if



  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False


  Strsql = "Select * From Objeven Where codobjeto=" & nCodigo & " Order By secuen"
 '' response.write Strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
  CONT  = CONT + 1
    sForm = "xxedEventos" & rs(0)

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objeven&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if



    'Eliminar
    if bModificar and (rs(13) & "") <> "PL" then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objeven&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
 ''   Celda = ofv.BotonEliminar(rs(2),sHref)
      Celda = ofv.BotonEliminar("Eliminacion",sHref)
      Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Descripcion
    if bModificar and (rs(13) & "") <> "PL" then

      xfunc = " onblur=" & chr(34) & " if (this.value.length > 10) {alert('El texto es demasiado"
      xfunc = xfunc & " extenso'); return false;} else {"
      xfunc = xfunc & " document." & sform  & "r2" & ".submit() ;}" & chr(34)

      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm & "X1","200","dt maxlength=70","")
      sInput = REPLACE(sInput,"{ document." & sform  & "X1" & ".submit(); return true;}","{ if (this.value.length > 70) {alert('El texto es demasiado extenso'); return false;} else { document." & sform  & "X1" & ".submit(); return true;}}")
      sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")

    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","60","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","vt","","24%","","","",sInput,"si")

    'Secuencia
    if bModificar and (rs(13) & "") <> "PL" then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm & "X2","5","dtn","SI")
      sInput = ofv.GenForm(sForm & "X2","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenerarInput("",rs(4),"readonly","","5","dtn","SI")
    end if
    Celda = Celda & ofv.GenCelda("","vt","","5%","","","",sInput,"si")

   if wkaplic = "PL" then

   else

    'Proceso
    if  bModificar then
        Strsql = "Select * From CSPROCESOS Order By DSCSECUEN"
         IF LEN(rs(11).value &"" )=0 THEN
          VALOR=""
        ELSE
         VALOR=rs(11).value
        END IF
        '' sInput = ofv.GenerarCombo(sForm,rs(11).name,VALOR,Strsql,cn,0,10,"")

         Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
       scomK = ""'"null" & Chr(9)

      scomD = "" '"" & Chr(9)
      SCOMSEP =""
      do while not rsperfil.eof
       scomK=scomK & SCOMSEP & rsperfil(0)
       scomd=scomd & SCOMSEP & rsperfil(10)
       SCOMSEP = chr(9)
       rsPerfil.movenext
        loop

      sInput =  ofv.gencomboFX(sForm & "X3", rs(11).name, valor, scomK, scomD, "")
      sInput = ofv.GenForm(sForm & "X3","",sAccion,"","post","",sInput,"si")

    else
        if  cstr(rs(11) & "") <> "" then
            Strsql = "Select DSCPROC From CSPROCESOS Where CODPROCESO = '" & rs(11) & "' "
            set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
            call ofv.cerrarconsulta(rsInstrumento)
        end if
        sInput = ofv.GenerarInput("",sInstrumento,"readonly","","90","dt","")
    end if

If session("APLMDCLI")="BGAL" then
    Celda = Celda & ofv.GenCelda("","vt","","50%","","","",sInput,"si")
End If

   end if

    formu =""


    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
  loop

   'sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","0","0","","",Formulario,"si")
   if CONT = 0 then
       	    sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes eventos agregados.</div>"
       	else
       	    sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","0","0","","",Formulario,"si")
       	    sHTML = sHTML & tabla
	end if


  END IF


  'Botonera
  if bModificar then
    sAccion = "AgregarEvento.asp?Tabla=Objeven&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
    Agregar = ofv.BotonAgregar(sAccion)

  end if

  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

        ' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if




'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

    sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>