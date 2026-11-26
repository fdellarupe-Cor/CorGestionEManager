<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

 ' RESPONSE.WRITE REQUEST.QUERYSTRING("XWIN")
  CantRegMostrarx = 8
  XWIN			 = request.querystring("XWIN")
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_XWIN=" & XWIN
  sOPCION       = request.querystring("OPCION")
  SCODINDI       = request.querystring("ID")
  volver=request.querystring("volver")
  xfuncion=request.querystring("xfuncion")
  xcantvar=request.querystring("cantvar")

  if SCODINDI<>"" then
   session("SCODINDI") =SCODINDI
  else
   SCODINDI = session("SCODINDI")
  end if




  if volver<> "" then
   session("volver")=volver
  else
   volver=session("volver")
  end if
  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  dim arrVar(6)
  DIM ArrPorcentaje(1)

  arrVar(0)="A"
  arrVar(1)="B"
  arrVar(2)="C"
  arrVar(3)="D"
  arrVar(4)="E"
  arrVar(5)="F"
  arrVar(6)="G"




     Strsql = "select * FROM orindicadores WHERE OBJID=0" & SCODINDI
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


      Strsql = "select * FROM orrelaindiriesgo WHERE idindicador=0" & SCODINDI
      'response.write strsql
      set rst = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	   if not rst.eof and rs("estado")=2 then
	     TieneRel="SI"
	   else
	     TieneRel="NO"
	   end if
	  call ofv.cerrarconsulta(rst)

     '  response.write TieneRel



     'strsql= "select * from orusrlst  where indicador=" & SCODINDI & " and usrresp='USR-" & rs("usuario") & "'"
	 strsql= "select * from orusrlst  where indicador=" & SCODINDI
     set rsresp = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     if rsresp.eof then
      strsql2= "insert into orusrlst (indicador,usrresp)  values(" & SCODINDI & ",'USR-" & rs("usuario") & "')"
      call ofv.crearconsultaEx(StrSql2,cn,1,parametros)
     end if
     call ofv.cerrarconsulta(rsresp)


     xmodifica=cbool(Fmodifica(rs("estado"),"",""))
     'response.write rs("estado")
     if not xmodifica OR (XWIN="ADDRELIND" and rs("estado")=2) then
	     smod=" readonly "
	    else
	     smod=""
     end if


     formula=rs("formula")
     cantvarz=rs("cantvar")

     if cint(formula) = cint(xfuncion) then
       FUNCAMBIO=cint("0" & cantvarz) <> cint("0" & xcantvar)
       xcantvar=rs("cantvar")
     else
       FUNCAMBIO=  (cint(formula) <> cint(xfuncion) or  cint("0" & cantvarz) <> cint("0" & xcantvar))
     end if


     ' tipo de indicador
     TIPO=RS("TIPO")

     STRSQL2="SELECT  CANTVAR,tipof,DEF FROM csORFORMULAS WHERE OBJID=0" & RS("FORMULA")
     set rsF = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
	 if not RSF.eof  then
	    TIPOFORMULA=RSF(1)
	    if rsf(1) <> 4 then
	      xcantvar=rsf(0)
	      IF RSF("DEF") ="PORCENTAJE" THEN xcantvar=2
	      STRSQLu= "UPDATE ORINDICADORES SET CANTVAR=0" & xcantvar & " WHERE OBJID=0" &  RS(0)
	     call ofv.crearconsultaEx(StrSqlU,cn,1,parametros)

	   end if

	 end if

	 'RESPONSE.WRITE FUNCAMBIO & "2222"
     IF FUNCAMBIO THEN
      STRSQL2= "DELETE FROM ORVARIABLES WHERE  IDINDICADOR=" & SCODINDI
      CALL  ofv.crearconsultaEx(StrSql2,cn,1,parametros)
     END IF

    Strsql = "select * FROM orindicadores WHERE OBJID=0" & SCODINDI
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    cantvar=rs("cantvar")
    strsql2="select * from orvariables where idindicador=" & SCODINDI
    set rsx = ofv.crearconsultaEx(StrSql2,cn,1,parametros)

     if rsx.eof then
      call ofv.cerrarconsulta(rsx)

      for i=0 to cantvar-1
       strsqlv=" insert into orvariables (tipocarga,activa,grupo,idindicador,VARNOMBRE)"
       strsqlv=strsqlv & "  VALUES (1,0,1," & SCODINDI &",'" & arrVar(i) & "')"
       call ofv.crearconsultaEx(strsqlv,cn,1,parametros)
      next
      strsqlv=" insert into orvariables (activa,grupo,idindicador,VARNOMBRE)"
      strsqlv=strsqlv & "  VALUES (0,1," & SCODINDI &",'Z')"
      call ofv.crearconsultaEx(strsqlv,cn,1,parametros)

     end if



   Celda = ""

   wkproccod = "&nbsp;"

   Celda =  ofv.GenCelda("","uti  ","","","","","","TABLA DE INDICADORES","si")

   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND  style='border-style:OUTSET;border-width:1;'  ","","98% ","30","0","","-1","1","","",Fila,"si")




  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""




  if rs.eof then

    Celda = ofv.gencelda("","vT  style='border-style:SOLID;border-width:1;border-color:#f93333;' ","center","","","","","<big>NO HAY CARGADOS INDICADORES</big>","si")
    Formulario = ofv.genrow("","","","","60","","",Celda,"si")

  ELSE

    Celda = ""

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","2%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","10%","","","","ACCESO","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","10%","","","","LISTA&nbsp;USR","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","20%","","","","UNIDAD MEDIDA&nbsp;DATOS","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","20%","","","","UNIDAD MEDIDA&nbsp;RESULTADO","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","20%","","","","FRECUENCIA","si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","20%","","","","CLASIFICACION","si")


 Celda=""

  posI = 0
 ' Strsqleco="select * from eco order by descri"

  scomD = "PUBLICO" & chr(9) & "PRIVADO" & CHR(9)
  scomK = "1" & chr(9) & "2" & CHR(9)

  Do While (Not rs.EOF)
    sForm = "form" & posi

    Celda = ""
    'Tipo
    sInput = rs(0)

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","ACCESOS","si")
    ' ACCESO
    if bModificar and XWIN <> "ADDRELIND" then
	     sInput = ofv.combosino(sForm ,rs(15),rs(15).name)
	     sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=orindicadores&volver=" & vuelta
	     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	else
	      sinput = ofv.gencelda("","ut","","","","","",rs(6),"si")
    end if

     IF rs(15)= 1 then
     acceso="PRIVADO &nbsp;&nbsp;"
     else
     acceso="PUBLICO&nbsp;&nbsp;"
     end if

     if  bmodificar and XWIN <> "ADDRELIND" then
      sImagenP = "&#0232;"
	  sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
	  sinput = sinput & "&nbsp;&nbsp;&nbsp;&nbsp;<a class='ft1n' " & sFuncAnc & " href='orasgusracc.asp?cod=4&volver=" & vuelta2
	  sinput = sinput  & "&Codobjeto=" & rs(0) & "&Tipobjeto=" & sTipo & "&circuito=" & rs(0)
	  sinput = sinput  & "' title='Lista de autorizados.' >" & sImagenP & "</a>"
     else
      sinput="&nbsp;"
     end if
     Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",ACCESO & sInput,"si")


     Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","","","","","UN. DE MED. DATOS","si")

     STRSQL="SELECT * FROM ORMEDIDAS ORDER BY DESCRI"

     r18val = rs(5) & ""
	 if r18val = "" then r18val = "0"
     if  xmodifica then
	  sInput = ofv.GenerarCombo(sForm,rs(5).name,r18val,Strsql,cn,0,1,"") 'replace(ofv.GenerarCombo(sForm,rs(5).name,r18val,Strsql,cn,0,1,""),"NAME=",ACTUALIZAR & " NAME=")
	 else
	  sInput = ofv.GenerarCombo("",rs(5).name,r18val,Strsql,cn,0,1,"") 'replace(ofv.GenerarCombo(sForm,rs(5).name,r18val,Strsql,cn,0,1,""),"NAME=",ACTUALIZAR & " NAME=")
	 end if

     Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' "  ,"","","","","", sInput,"si")

     Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","","","","","UN. DE MED. RESUL","si")

     STRSQL="SELECT * FROM ORMEDIDAS ORDER BY DESCRI"

      r18val = rs(11) & ""
	  if r18val = "" then r18val = "0"

     if  xmodifica then
	  sInput = ofv.GenerarCombo(sForm,rs(11).name &  actualizar,r18val,Strsql,cn,0,1,"")
	 else
	  sInput = ofv.GenerarCombo("",rs(11).name &  actualizar,r18val,Strsql,cn,0,1,"")
	 end if
     Celda = Celda &  ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","", sInput,"si")



    Filax  = FILAX & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    if  xmodifica then
     SINPUT = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"30","dt","")
    else
     SINPUT = ofv.GenerarInput(rs(1).name,rs(1),"text readonly",sForm,"30","dt","")
    end if

    Celda =  ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","NOMBRE","si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")




     Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","LISTA RESPONSABLES","si")
   ' LST USR
	 if bModificar  and XWIN <> "ADDRELIND" then
		   sInput = ofv.combosino(sForm ,rs(16),rs(16).name)
		   'sInput = ofv.gencelda("","uti","","","","","",sInput,"si")
		   sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=orindicadores&volver=" & vuelta
		   sAccion = sAccion & "&Codobjeto=" & rs(0) & "&Tipobjeto=" & sTipo
		   'sInput =  ofv.GenForm(sForm,"",sAccion,"","post","",sInput,"si")
		else
		   sinput = ofv.gencelda("","ut","","","","","",rs(6),"si")
	 end if



	  IF rs(16)= 1 then
	    acceso="&nbsp;&nbsp;LISTA&nbsp;&nbsp;"
	  else
	    acceso="SIN&nbsp;LISTA&nbsp;&nbsp;"
	  end if
      if bmodificar and XWIN <> "ADDRELIND" then
      sImagenP = "&#0232;"
	  sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
	  sinput = sinput & "&nbsp;&nbsp;&nbsp;&nbsp;<a class='ft1n' " & sFuncAnc & " href='orasgusracc.asp?cod=1&volver=" & vuelta2
	  sinput = sinput  & "&Codobjeto=" & rs(0) & "&Tipobjeto=" & sTipo & "&circuito=" & rs(0)
	  sinput = sinput  & "' title='Lista de autorizados.' >" & sImagenP & "</a>"
	  else
	   sinput="&nbsp;"
	  end if
      'Celda = Celda & ofv.gencelda("","cc","","5","","","",ACCESO &"&nbsp;&nbsp;&nbsp;" &sInput,"si")
      Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",ACCESO & sInput,"si")




     Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","","","","","FRECUENCIA","si")

	 'STRSQL="SELECT * FROM ORFRECUENCIA ORDER BY VALOR"  'Original
	 STRSQL="SELECT * FROM ORFRECUENCIA WHERE ACTIVO = 'S' ORDER BY VALOR"

	  r18val = rs(12) & ""

	  if r18val = "" then r18val = "0"

     if  xmodifica then
       sInput = ofv.GenerarCombo(sForm,rs(12).name ,r18val,Strsql,cn,0,1,"")
      else
       sInput = ofv.GenerarCombo("",rs(12).name ,r18val,Strsql,cn,0,1,"")
     end if
	 Celda = Celda &  ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","", sInput,"si")



     'Celda = Celda & ofv.gencelda("","","","","30%","","",sInput,"si")
     Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","","","","","CLASIFICACION","si")
	 STRSQL="SELECT * FROM ORCLASIFICAINDI ORDER BY DESCRI"
	 r18val = rs(18) & ""

	 if r18val = "" then r18val = "0"
     IF    XWIN <> "ADDRELIND" THEN
	  sInput = ofv.GenerarCombo(sForm,rs(18).name,r18val,Strsql,cn,0,1,"")
	 ELSE
	  sInput = ofv.GenerarCombo("",rs(18).name,r18val,Strsql,cn,0,1,"")
	 END IF
	 Celda = Celda &  ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","", sInput,"si")




     Filax  = FILAX & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
    celda=""
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","","","","","ESTADO","si")

    sinput="<select   name="&rs(8).name &"   onchange=validar('"&sForm&"','"&TieneRel&"');>"
    if rs("estado") =1 then
	 sinput=sinput & "<option value=1>NUEVO</OPTION>"
	end if
	sinput= sinput & "<option value=2>INICIADO</OPTION>"
	if tienerel <> "SI" THEN
	 sinput=  sinput &"<option value=3>NO VIGENTE</OPTION>"
	END IF
	sinput=sinput & "</SELECT>"

    sinput=replace(sinput,"value=" & rs(8),"value=" & rs(8) & " selected")


	    Celda = Celda &  ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","", sInput,"si")


   ' IF TIPO=0 THEN
      Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","","","","","FORMULA","si")
      Strsql = "Select OBJID,DESCRI  From ORFORMULAS ORDER BY DESCRI"

      if xmodifica then
        sInput = ofv.GenerarCombo(sform,rs(13).name,rs(13).value,Strsql,cn,0,1,"")
       else
        sInput = ofv.GenerarCombo("",rs(13).name,rs(13).value,Strsql,cn,0,1,"")
      end if
    '  if not xmodifica then sinput=replace(sinput,"onclick","xonclick")

      IF TIPOFORMULA=4 and xmodifica THEN
       SINPUT2="<input name="&rs(14).name &" align=right value=" & rs(14) & " type=text class=dt onchange="&chr(34)&"valnum(this.value,'"&sForm&"');"&chr(34) & ">" 'ofv.GenerarInput(rs(14).name,rs(14),"textsb",sForm,"5","dt  onkeypress="&chr(34)&"valnum(this.value,"&sForm&");"&chr(34) ,"")
      ELSE
       if TIPOFORMULA=4 then
        SINPUT2="<input name="&rs(14).name &" align=right value=" & rs(14) & " type=text class=dt readonly>"
       else
        sinput2=""
       end if
      END IF
      Celda = Celda &  ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","", sInput & SINPUT2,"si")
'
	'STRSQL_Owner = "SELECT * FROM dbo.Usuarios WHERE ID = '" & rs("usuario") & "'"

	''set owner = ofv.crearconsultaEx(STRSQL_Owner,cn,1,parametros)

	xxx=getinfouser(rs("usuario"))

	Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' align=left","","10","","","","CREADOR","si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",session("SGU_nombre"),"si")

    Filax  = FILAX & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    sAccion = ""

    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORINDICADORES&volver=" & vuelta &"_XFUNCION=" & RS("FORMULA") & "_CANTVAR=" & xcantvar 'RS("CANTVAR")
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


    posi = posi + 1
    rs.Movenext
  Loop

  Celda = ""

  ' muestro las variables
  Celda = ofv.gencelda("","VTI  COLSPAN=6 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big> CARGA DE VARIABLES</BIG>","si")
  FILAC=ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
  Celda = ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big></BIG>","si")
  Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>NOMBRE</BIG>","si")
  Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>DESCRIPCION</BIG>","si")
  Celda = CELDA &ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>TIPO&nbsp;DE&nbsp;CARGA</BIG>","si")

  FilaC  = FILAC & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

   Strsql = "select * FROM csorvariables2 WHERE VARNOMBRE <> 'Z' AND idindicador=0" & SCODINDI & " order by  varnombre"
 '  response.write strsql
   Formulario2=""
   set rsvar = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   POSI=0
   resul=0

   scomK = "1" & chr(9) & "2"
   scomD = "MANUAL" & chr(9) & "AUTOMATICO"

   do while not rsvar.eof



   sForm = "Vform" & posi
   Celda = ofv.gencelda("","VTI  ","","","","","",rsvar("varnombre")& "=","si")




   sInput = ofv.GenerarInput(rsVAR(9).name  ,rsVAR(9),"textsb",sForm,"70","dt" & smod,"")
   Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


   sInput = ofv.GenerarInput(rsVAR(4).name  ,rsVAR(4),"textsb",sForm,"70","dt" & smod,"")
   Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
   if xmodifica then
    sInput = ofv.gencomboFX(sForm,rsVAR(7).name  ,rsVAR(7)&"",scomK,scomD," ")
   else
    sInput = ofv.gencomboFX("",rsVAR(7).name  ,rsVAR(7)&"",scomK,scomD," ")
   end if
   Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")




   Filax  =  ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda  ,"si")


   if bModificar  then
         sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rsVAR(0)," ","*")
         sAccion = sAccion & "&tabla=ORVARIABLES&volver=" & vuelta  &"_XFUNCION=" & formula & "_CANTVAR=" & xcantvar
   end if



   Formulario2 = Formulario2 & ofv.GenForm(sForm,"",sAccion,"","post","",FilaX ,"si")

   rsvar.movenext
   POSI=POSI + 1
   loop
   call ofv.cerrarconsulta(rsvar)
   sInput = resul

  tabla2="<BR>" &  ofv.GenTabla("","","","98% ","","0","","0","0","","",filac & Formulario2 & Filaxz,"si")
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98%","","0","","0","0","","",Formulario,"si") & chr(13) & tabla2 & UMBRALES(SCODINDI) & RIESGOSASOCIADOS(SCODINDI)



  'Botonera
  Agregar = ""
  if bAgregar  then
    sAccion = "ORCARGAINDICADORES.asp"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
  Buscar  = ""
  sAccion = volver & "?&XWIN=" & XWIN 'request.querystring '"orCARGAUSRINDI.asp"

 'response.write xwin
  Celda = ofv.BotonVolver(sAccion)
  'esponse.write request.querystring
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","",Buscar,celda)

  call ofv.cerrarconsulta(rs)




  call ofv.cerrarconn(cn)
	   sHTML = sHTML & "<script type='text/javascript'> "

	  sHTML = sHTML & "function disp_confirm(disp_msg)"
      sHTML = sHTML & "  {"

      sHTML = sHTML & " var i; var r=confirm(disp_msg);"
      sHTML = sHTML & "  if (r==true){"
      sHTML = sHTML &  "      "
      sHTML = sHTML & "    }else{"
      sHTML = sHTML & "    alert('Se cancelo la accion!');"
      sHTML = sHTML & "    }"
      sHTML = sHTML & "  }"

	   sHTML = sHTML & "function valnum(num,disp_frm)"
	   sHTML = sHTML & "  {"

	   sHTML = sHTML & "  if (num<=5)"
	   sHTML = sHTML & "    {"
	   sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"

	   sHTML = sHTML & "    fr.submit();"
	   sHTML = sHTML & "    }"
	   sHTML = sHTML & "  else"
	   sHTML = sHTML & "    {"
	   sHTML = sHTML & "    alert('No se Adminte màs de 5 Variables!'); return false;	"
	   sHTML = sHTML & "    }"
   	   sHTML = sHTML & "  }" & chr(13)



       sHTML = sHTML & "function valnumUmbral(num,disp_frm,val2,col)"
   	   sHTML = sHTML & "  { num.value=parseInt(num.value.replace(',','.'));val2=parseInt(val2.replace(',','.'));"
       'sHTML = sHTML & "  if (col =='1') {result= num.value < val2;} else {result= num.value > val2} ;"
       sHTML = sHTML & "  if (col =='1') {result= true;} else {result= num.value > val2} ;"
   	   sHTML = sHTML & "  if (result) "
   	   sHTML = sHTML & "    {  "


   	   sHTML = sHTML & "     var fr= document.getElementById(disp_frm);" & chr(13)

   	   sHTML = sHTML & "     fr.submit();" & chr(13)
   	   sHTML = sHTML & "    } "
   	   sHTML = sHTML & "  else " & chr(13)
   	   sHTML = sHTML & "    {"
       sHTML = sHTML & "    if  (col !='1') { alert('debe ser mayor que '.concat(val2.toString()));} else { "
   	   sHTML = sHTML & "    alert('debe ser menor que '.concat(val2.toString())); } return false;	"  & chr(13)
   	   sHTML = sHTML & "    }" & chr(13)
       sHTML = sHTML & "  }"
'alert('num=' + num + '   val2=' + val2 + '  val3=' + val3 + ' col='+ col);
       sHTML = sHTML & "function valnumUmbral2(num,disp_frm,val2,col,val3)"
   	   sHTML = sHTML & "  {var result;   num=parseInt(num.replace(',','.'));val2=parseInt(val2.replace(',','.'));val3=parseInt(val3.replace(',','.'));"
       sHTML = sHTML & "  if (col != '3' ) { result= num > val3 ;} else {result= val2 > val3 ;};"
   	   sHTML = sHTML & "  if (num > val2 && result ) "
   	   sHTML = sHTML & "    { "
'alert('num=' + num + '   val2=' + val2 + '  val3=' + val3 + ' col='+ col);

   	   sHTML = sHTML & "     var fr= document.getElementById(disp_frm);" & chr(13)

   	   sHTML = sHTML & "     fr.submit();" & chr(13)
   	   sHTML = sHTML & "    } "
   	   sHTML = sHTML & "  else " & chr(13)
   	   sHTML = sHTML & "    {"

       sHTML = sHTML & "    alert('Error');"
   	   sHTML = sHTML & "    }" & chr(13)
   	   sHTML = sHTML & "  }"



       sHTML = sHTML & "function valnumUmbralN(num,disp_frm,val2,tope)"
	   sHTML = sHTML &    "  {num.value=parseFloat(num.value.replace(',','.')); val2=parseFloat(val2.replace(',','.'));"
       sHTML = sHTML &    "   var tope=parseFloat(tope.replace(',','.')); var topeV=false;"
       shtml = shtml &    " if (tope.toString()!='NaN'){"
       shtml = shtml &       "  if (num.value < tope  || tope==0 ) {topeV=true;topeM='';}"
        shtml = shtml &     " else{ topeV=false; topeM=' y menor que ' + tope ;}}"

       shtml = shtml &     " else{"
       shtml = shtml &    " topeV=true; topeM='';} "

	   sHTML = sHTML & "  if ( (num.value > val2) && topeV ) "
	   sHTML = sHTML & "    {  "


	   sHTML = sHTML & "     var fr= document.getElementById(disp_frm);" & chr(13)

	   sHTML = sHTML & "     fr.submit();" & chr(13)
	   sHTML = sHTML & "    } "
	   sHTML = sHTML & "  else " & chr(13)
	   sHTML = sHTML & "    {"

	   sHTML = sHTML & "    alert('debe ser mayor que ' + val2 + topeM);  return false;	"  & chr(13)
	   sHTML = sHTML & "    }" & chr(13)
   	   sHTML = sHTML & "  }"



   	   sHTML = sHTML & "function validar(disp_frm,tienerel)"
	   sHTML = sHTML & "  {"
	   shTML = sHTML & "  var fr= document.getElementById(disp_frm); r=confirm('Confirma los Umbrales');"
	'   shTML = sHTML & "  if(tienerel=='SI') {alert('Tiene Riesgos relacionados!')} else {"
	   sHTML = sHTML & "    if ( fr.CLASIFICACION.value!='0' && fr.UNIDADFRECUEN.value!='0' && fr.UNIMEDIDADATOS.value!='0' "
	   sHTML = sHTML & " && fr.UNIMEDIDARESUL.value!='0'&& r )    {fr.submit();} else {if (r==true) {alert('Debe cargar todos los datos!');};fr.ESTADO.selectedIndex=0; return false;} "
	   sHTML = sHTML & "  }"


       sHTML = sHTML & "</script>"

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

'--------------------------------------------------------------------------------------------------------

FUNCTION UMBRALES(SCODINDI)
DIM RS
DIM CELDA
DIM SHTML
DIM Formulario
DIM sForm

  WKSETTABLA = "ORUMBRALES"
 ' formulario="222"
  Strsql = "Select * From " & WKSETTABLA & " where idindicador="& SCODINDI & " Order By ORDEN"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  IF RS.EOF THEN

   Strsql = "Select * From " & WKSETTABLA & " where idindicador=0 Order By ORDEN"
   set rsux = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   do while not rsux.eof

    STRSQL=" Insert into ORUMBRALES (valor1,valor2,IDINDICADOR,ORDEN,COLOR,DESCRIPCION) VALUES (0,0," & SCODINDI & "," & RSUX("ORDEN")  &",'" &RSUX("COLOR") &"','" & RSUX("DESCRIPCION")&"')"
   ' RESPONSE.WRITE STRSQL
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
    rsux.movenext
   loop

   CALL OFV.CERRARCONSULTA(RSUX)

  END IF

  CALL OFV.CERRARCONSULTA(RS)

  Strsql = "Select a.* ,(select valor2 from orumbrales b where b.IDINDICADOR=a.idindicador and b.orden=a.orden-1) as valprox From orumbrales a  WHERE a.IDINDICADOR="&SCODINDI&" Order By a.ORDEN"
  ' Strsql = "Select * From " & WKSETTABLA & " WHERE IDINDICADOR="&SCODINDI&" Order By ORDEN"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 ' if not rs.eof then rs.move(CantRegMoverx)
  posI = 1
Do While (Not rs.EOF)
    sForm = "Uform" & posi

      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")

     if posi=1 then
      vAntdesde=rs("valor1")
      vAnthasta=rs("valor2")
    end if
    sInput = RS(1) 'ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")


    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text"& actualizar,sForm,"30","dt" & smod,"")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

     nCol = 40


           wdescri=""
	       if  not  isnull(rs(6)) then
	           wdescri = TRIM("" &rs(6))
	       end if
'
	       if  len(wdescri) = 0 then wdescri = " "
'
	       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' " & smod
	       t3b = t3b & "name=" & rs(6).name & " cols=" & nCol & " rows=2 "
	       if  bModificar  then
	           t3b = t3b & " onblur=" & sForm & ".submit() "
	           t3b = t3b & " onkeypress=" & chr(34)
	           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
	           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
	       else
	           t3b = t3b & " readonly >"
	       end if
	       xdatos =  wdescri
	       t3b    = t3b & xdatos & "</textarea></font>"
	       SINPUT = t3b
'
'
	    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")



        if posi = 1 then
  		SINPUT="<input name="&rs(4).name  & smod &" align=right value=" & rs(4) & " type=text class=dt onchange="&chr(34)&"valnumUmbral(this,'"&sForm&"','"& vAnthasta &"','1');"&chr(34) & ">"
  		else
  		SINPUT="<input name="&rs(4).name  & smod &" align=right value=" & rs(4) & " type=text class=dt onchange="&chr(34)&"valnumUmbral(this,'"&sForm&"','"& vAnthasta &"','2');"&chr(34) & ">"
  		end if


		Celda = Celda & ofv.gencelda("","cc","","10","","","",sInput,"si")
'
       if rs("valprox")&"" ="" then
         valorProx=-1000
       else
        valorProx=rs("valprox")
       end if
		SINPUT="<input name="&rs(5).name   & smod &" align=right value=" & rs(5) & " type=text class=dt onchange="&chr(34)&"valnumUmbral2(this.value,'"&sForm&"','"& vAntdesde&"','3','"&valorProx&"');"&chr(34) & ">" 'ofv.GenerarInput(rs(14).name,rs(14),"textsb",sForm,"5","dt  onkeypress="&chr(34)&"valnum(this.value,"&sForm&");"&chr(34) ,"")

        Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")



      '   SINPUT="<input name="&rs(4).name  & smod &" align=right value=" & rs(4) & " type=text class=dt onchange="&chr(34)&"valnumUmbral(this,'"&sForm&"','"& vhasta&"','"&rs("valprox")&"');"&chr(34) & ">"


     '   Celda = Celda & ofv.gencelda("","cc","","10","","","",sInput,"si")
     '   if posi=1 then
      '    SINPUT="<input name="&rs(5).name  & smod &" align=right value=" & rs(5) & " type=text class=dt onchange="&chr(34)&"valnumUmbral2(this,'"&sForm&"','"& vdesde &"','"&rs("valprox")&"');"&chr(34) & ">"
      '  else
      '  end if


       ' Celda = Celda & ofv.gencelda("","cc","","10","","","",sInput,"si")


       sInput = "&nbsp;"
       Celda = Celda & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & RS(7) & ";' ","CENTER","5%","","","",sInput,"si")



       Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta &"_XFUNCION=" & formula & "_CANTVAR=" & xcantvar
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1


    vAnthasta=rs(5)

    rs.Movenext
    if not rs.eof then
       vAntdesde=rs(4)
    else
       vAntdesde=0
    end if

  Loop
  'sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

    Celda = "<BR>" & ofv.gencelda("","VTI  COLSPAN=9 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>CARGA DE UMBRALES</BIG>","si")
    FILAC=ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
    Celda = ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","","si")
    Celda = CELDA & ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","1%","","","","<big>ORDEN</BIG>","si")
    Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","<big>DESCRIPCION</BIG>","si")
    Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","30%","","","","<big>OBSERVACION</BIG>","si")
    Celda = CELDA &ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","10%","","","","<big>DESDE</BIG>","si")
    CELDA = CELDA &  ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","10%","","","","<big>HASTA</BIG>","si")
    CELDA = CELDA &  ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>COLOR</BIG>","si")
    FilaC  = FILAC & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  UMBRALES = ofv.GenTabla("","","","80% ","","0","","0","0","","",FILAC & Formulario,"si")

END FUNCTION

FUNCTION RIESGOSASOCIADOS(SCODINDI)
DIM RS
DIM CELDA
DIM SHTML
DIM Formulario
DIM sForm

' Strsql = "Select * From CSRELPATHDESCRI  where COBJ
  Fila=""
  WKSETTABLA = "ORRELAINDIRIESGO"

  Strsql = "Select CODRIESGO  From " & WKSETTABLA & " where idindicador="& SCODINDI
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 1
  FILA=""
  Do While (Not rs.EOF)

    on error resume next
    RUTA=""
    STRSQL="SELECT * FROM CSORBUSQRIESGO WHERE CODRIESGO='" & RS(0) & "'"
  '  response.write strsql & "<br>"
    SET RSRIESGO=ofv.crearconsultaEx(StrSql,cn,1,parametros)
   ' on error resume next
    if  not RSRIESGO.eof then

    Celda =ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RSRIESGO("RIESGONOMBRE"),"si")
    Celda = CELDA & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RSRIESGO("DESCRIACT"),"si")
    Celda =CELDA & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RSRIESGO("PERIODO"),"si")
    Strsql = "Select * From CSRELPATHDESCRI  where COBJ=0" & RSRIESGO("CODOBJETO")
   ' response.write strsql
    SET RSRIESGOPATH=ofv.crearconsultaEx(StrSql,cn,1,parametros)
    DELIMX=""
	FOR I=1 TO 7
	  IF RSRIESGOPATH("EXPR" & I)&"" <> "" THEN
	    RUTA=RUTA & DELIMX & RSRIESGOPATH("EXPR" & I)
	    DELIMX="/"
	  END IF
    NEXT
    CELDAX=ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;' ","LEFT","","","","",RUTA,"si")
    Fila = FILA & ofv.genrow("","","","","","","",CELDAX & Celda,"si")
    CALL OFV.CERRARCONSULTA(RSRIESGO)
    CALL OFV.CERRARCONSULTA(RSRIESGOPATH)
    end if
    rs.Movenext

  Loop

  IF FILA <> "" THEN
    Celda = "<BR>" & ofv.gencelda("","VTI  COLSPAN=4 STYLE='background-color:#bbbbbb;color:#ffffff;' ","","","","","","<big>LISTA DE RIESGOS ASOCIADOS</BIG>","si")
    FILAC=ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda = ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","40% ALIGN=CENTER","","","","<big>PROCESO/SUBPROCESOS</BIG>","si")
	Celda = CELDA & ofv.gencelda("","VTI   STYLE='background-color:#bbbbbb;color:#ffffff;' ","","35% ALIGN=CENTER","","","","<big>RIESGO</BIG>","si")
	Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","15% ALIGN=CENTER","","","","<big>ACTIVIDAD</BIG>","si")
	Celda = CELDA & ofv.gencelda("","VTI  STYLE='background-color:#bbbbbb;color:#ffffff;' ","","10% ALIGN=CENTER","","","","<big>PERIODO</BIG>","si")
    FilaC  = FILAC & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    RIESGOSASOCIADOS = ofv.GenTabla("","","","98%% ","","0","","0","0","","",FILAC & Fila,"si")
 ELSE
    RIESGOSASOCIADOS=""
 END IF
   CALL OFV.CERRARCONSULTA(rs)
END FUNCTION


%>