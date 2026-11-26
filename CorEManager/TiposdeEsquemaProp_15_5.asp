<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  redim nivel(9)
  nivel(0) = "C"
  nivel(1) = "D"
  nivel(2) = "E"
  nivel(3) = "F"
  nivel(4) = "G"
  nivel(5) = "H"
  nivel(6) = "I"
  nivel(7) = "J"
  nivel(8) = "K"
  nivel(9) = "L"
  sImagenCampo = "ARCHIVO"
  CantRegMostrarx = cdbl(request.querystring("CantRegMostrar"))
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  wPerfil = request.querystring("wPerfil")
  wNivel = request.querystring("wNivel")

  if isnull(wNivel) then wNivel = "NO"
  if len(wNivel) = 0  then wNivel = "NO"


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wPerfil=" & wPerfil

  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 '' Response.Write bModificar & " bModificar "
 '' Response.Write bAgregar   & " bAgregar   "
 '' Response.Write bSuprimir  & " bSuprimir  "
 '' Response.Write bConsultar & " bConsultar "


  set cn = ofv.conectar(ofv.strconn2)
  set cn3 = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  Strsql = "Select * From Tipobjeto where tipobjeto = '" & wperfil & "' "
 'Response.Write Strsql & " Strsql "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posI = 0

    sForm = "form" & posi

     Fila = ""


 ' Encabezados 1

    'Descripcion

    sInput = rs(0) & " - " & rs(1)
    Celda = ofv.gencelda("","TTT COLSPAN=3 ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","20%","","","","Perfil","si")

  '  sInput = rs(3)


    Sql1 = "Select Tipoperfil,DESCRI From Objperfil Order By Tipoperfil"
    sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,Sql1,cn,0,1,"")

    ''  sInput = gencomboX(sForm,rs(3),rs(3).name,"0",Sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Nivel
     Celda = ofv.gencelda("","tt","","","","","","Nivel","si")

     SCOMSEP = ""
     scomK   = ""

     For i=0 to 9
         scomK = scomK & SCOMSEP & nivel(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

  ''  sInput = gencomboX(sForm,rs(4),rs(4).name,"1","N","",scomK,scomD,"")

    sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD,"")

    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Archivo
     Celda = ofv.GenCelda("","tt","","","","","","Orden","si")

    sInput=ofv.GenerarInput(rs(15).name,rs(15),"text",sForm,"10","dtn","si")
    'sImagenValor = rs(15)
    'sInput = ofv.GenerarCombo3(sForm,ofv.Getfiles("Iconos"),sImagenCampo,sImagenValor)

  ''  sInput = gencomboX(sForm,rs(15),rs(15).name,"0","A","Iconos",0,1,"")
    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Avisos de Revision
    '' AGREGADO 10102014 PARA DOCUMENTOS A VENCER

    Celda = ofv.GenCelda("","tt","","","","","","AVISO REVISION","si")

    SCOMSEPx = chr(9)
    scomKk = "0"
    scomDd = "NO"
    scomKk = scomKk & SCOMSEPx & "5"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 5 DIAS"
	scomKk = scomKk & SCOMSEPx & "10"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 10 DIAS"
	scomKk = scomKk & SCOMSEPx & "20"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 20 DIAS"
	scomKk = scomKk & SCOMSEPx & "30"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 30 DIAS"
    scomKk = scomKk & SCOMSEPx & "60"
    scomDd = scomDd & SCOMSEPx & "ANTES DE LOS 60 DIAS"

''' ORREVISION  = dFechan
''' SORREVISION = "NO"
''' sInput = ofv.gencomboFX(sForm,"ORREVISION",SORREVISION,scomKk,scomDd,"")
''' sInput = ofv.gencomboFX(sForm,rs(1).name,rs(1),scomKk,scomDd,"")
''' Celda = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
''  StrSql="SELECT * FROM Tipobjeto order by descri "
''  set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
''
''  SCOMSEP = ""
''  scomK   = ""
''  scomD   = ""
''
''   svalor1 = rs(24) & ""
''  'Response.write svalor1 & " aviso revision " & chr(13) & "<br>"
''	  do while not rsx1.eof
''
''  	 SCOMSEP = chr(9)
''     scomK = scomK & SCOMSEP & rsx1(0)
''     scomD = scomD & SCOMSEP & rsx1(1)
''
''  rsx1.movenext
''  loop
''' sInput = ofv.gencomboFX(sForm,"VALOR1",svalor1,scomK,scomD,"")
''' response.write " --> " & sInput & " sInput "
'''  Celda = Celda & ofv.gencelda("","vti","","50%","","","",sInput,"si")

  if bModificar then
       sInput = ofv.gencomboFX(sForm,rs(24).name,rs(24),scomKk,scomDd,"")
  else
       sInput = ofv.GenerarInput("",rs(24),"readonly","","50","dt","")
  end if

    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","50","","",Celda,"si")

'' FIN AGREGADO 10102014 PARA DOCUMENTOS A VENCER

      Celda = ofv.GenCelda("","tt","","","","","","TEXTO DE IMPRESIÓN","si")

      sInput = "<textarea name=TEXTOIMPRESION onkeypress=" & chr(34)
	        sInput = sInput & " if (this.value.length > 450) {alert('El texto es demasiado"
	        sInput = sInput & " extenso'); return false;}" & chr(34)
	   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs("TEXTOIMPRESION") & "","''","'") & "</textarea>"

      Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
      Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
    'filtro
   Celda = ofv.GenCelda("","tt","","","","","","ACEPTA COMENTARIOS","si")


    scomK = "S" & chr(9) & "N"
    scomD = "SI" & chr(9) & "NO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(26).name,rs(26),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(26),"readonly","","10","dt","")
    end if
    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


   ' IF rs(26)="S"	 THEN
   ' Celda = ofv.GenCelda("","tt","","","","","","AUTORIZANTES","si")




	'    if bModificar  then
	'        sImagenP = "&#0232;"
	'           sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
	'	       sinput = "<a class='ft1n' " & sFuncAnc & " href='ESQUEMASdkdet.asp?TIPOBJETO=" &wperfil & "&CantRegMover=" & CantRegMoverx
	'	       sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
   '       sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"
	'    else
	'        sInput = "&NBSP;" 'ofv.GenerarInput("",rs(26),"readonly","","10","dt","")
	'    end if
	'    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
'
   ' Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")
   'END IF
    Celda = ofv.GenCelda("","tt","","","","","","Filtro","si")


    scomK = "S" & chr(9) & "N"
    scomD = "SI" & chr(9) & "NO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(22).name,rs(22),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(22),"readonly","","10","dt","")
    end if
    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



 ' Encabezados 2

    'xxx

     Celda = ofv.GenCelda(""," ttt colspan=5","","","","","","     CIRCUITOS ","si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   ' Para Circuitos de Autorizacion y Observacion Default

     Celda = ofv.GenCelda("","tt","","","","",""," CIRCUITO DEFAULT ","si")

     if bModificar  then
  	    sImagenP = "&#0232;"
        sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
        sinput = "<a class='ft1n' " & sFuncAnc & " href='ESQUEMAScircdet.asp?TIPOBJETO=" &wperfil & "&CantRegMover=" & CantRegMoverx
        sinput = sinput & "&CantRegMostrar=" & CantRegMostrarx
        sinput = sinput & "' title='Circuitos Autorizacion Default' >" & sImagenP & "</a>"
     else
    	''sInput = "&NBSP;" 'ofv.GenerarInput("",rs(26),"readonly","","10","dt","")
     end if

     Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    'CIRCUITO
     Celda = ofv.GenCelda("","tt ","","20%","","","","AUTORIZACION","si")

    'SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = 'GEN'  )  OR (OBJTIPO = 'COR')) "
   	 SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = 'COR')) "


    if bModificar then

        StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
        set rsx1 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)

        SCOMSEP = ""
        scomK   = ""
        scomD   = ""
        SCOMSEP = chr(9)

        do while not rsx1.eof
           scomK = scomK & SCOMSEP & rsx1(0)
           scomD = scomD & SCOMSEP & rsx1(1)
           SCOMSEP = chr(9)
        rsx1.movenext
        loop

 ''     sInput = ofv.gencomboFX(sForm,"ORCARACTER",SORCARACTER,scomK,scomD,"")
        sInput = ofv.gencomboFX(sForm,rs(19).name,rs(19),scomK,scomD,"")
 ''     sInput = ofv.GenerarCombo(sForm,rs(19).name,rs(19),strsql,cn3,0,1,"")

    else

        StrSql = "Select * From WKFCKTMASTER WHERE WKFCOD = '" & rs(19) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

  'sForm1 = sForm
  'sForm1 = request.form("sForm1") & ""
  ''sInput1  = ofv.GenerarInput("sForm1",sForm1,"hidden","","1","dt","")

   Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")

   Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

 'CIRCUITO
     Celda = ofv.GenCelda("","tt","","","","","","OBSERVACION","si")

    SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = 'COR')) "

    if bModificar then
        StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
        set rsx1 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)

        SCOMSEP = ""
        scomK   = "NO"
        scomD   = "NO"
        SCOMSEP = chr(9)

        do while not rsx1.eof
           scomK = scomK & SCOMSEP & rsx1(0)
           scomD = scomD & SCOMSEP & rsx1(1)
           SCOMSEP = chr(9)
        rsx1.movenext
        loop

        sInput = ofv.gencomboFX(sForm,rs(20).name,rs(20),scomK,scomD,"")
  ''    sInput = ofv.GenerarCombo(sForm,rs(20).name,rs(20),strsql,cn3,0,1,"")

    else
        StrSql = "Select * From WKFCKTMASTER WHERE WKFCOD = '" & rs(20) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celda = CELDA & ofv.GenCelda("","vti ","","","","","",sInput,"si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



    'NIVEL
     Celda = ofv.gencelda("","tt","","","","","","NIVEL","si")

    if bModificar then
      strsql = "Select * From NivelConfiden Order By NCcod"
      sInput = ofv.GenerarCombo(sForm,rs(18).name,rs(18),strsql,cn2,0,1,"")
   '' sInput = gencomboX(sForm,rs(18),rs(18).name,"21",strsql,cn2,0,1,"")
    else
      strsql = "Select * From NivelConfiden where NCcod = " & rs(18)
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celda = CELDA & ofv.GenCelda("","vti ","","","","","",sInput,"si")




     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")






 ' Encabezados 3

    'xxx
     Celda = ofv.GenCelda("","ttt colspan=3 ","","","","","","Permiso de impresion","si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    'acceso
     Celda = ofv.GenCelda("","tt","","","","","","acceso","si")


    scomK = "S" & chr(9) & "N" & chr(9) & "L"
    scomD = "Todos" & chr(9) & "No" & chr(9) & "Solo acceso Local"

    'acceso
    if bModificar then
       '' sInput = gencomboX(sForm,rs(16),rs(16).name,"2","N","",scomK,scomD,"")
        sInput = ofv.gencomboFX(sForm,rs(16).name,rs(16),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(16),"readonly","","90","dt","")
    end if
    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'resol
     Celda = ofv.gencelda("","tt","","","","","","resolucion","si")

    scomK = "S" & chr(9) & "A"
    scomD = "Todos" & chr(9) & "Solo Alta resolucion"

    if bModificar then
     '''   sInput = gencomboX(sForm,rs(17),rs(17).name,"1","N","",scomK,scomD,"")
        sInput = ofv.gencomboFX(sForm,rs(17).name,rs(17),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(17),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")


    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Tipobjeto&volver=" & vuelta
    end if
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","-1","2","","",Formulario,"si") & "<br>"
  ''sHTML = sHTML & ofv.GenTabla("","","","80% border","","0","","-1","2","","",Formulario,"si") & "<br>"




 ' Encabezados 4






    'Acepta niveles
     Celda = ofv.gencelda("","ttt colspan=10","","","","","","Acepta niveles","si")


     Fila = ofv.genrow("","","","","","","",Celda,"si")





    Celda =  ""



    'Acepta linea 1

    for ii = 0 to 4

        sinput = "<a  href=TiposdeEsquemaProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
        sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx  & "&wNivel=" & nivel(ii) & " ><b>"
        sinput = sinput  & "<font color=#000080>" & nivel(ii) & "</font></b></a>"

        Celda = Celda & ofv.gencelda("","cc","center","2%","","","",sinput,"si")

        sInput = ofv.combosino(sForm & "A",rs(ii + 5),rs(ii + 5).name)


        Celda = Celda & ofv.gencelda("","vt","","5%","","","",sInput,"si")


    next

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    Celda = ""

    'Acepta linea 2

    for ii = 5 to 9

        sinput = "<a  href=TiposdeEsquemaProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
        sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx  & "&wNivel=" & nivel(ii) & " ><b>"
        sinput = sinput  & "<font color=#000080>" & nivel(ii) & "</font></b></a>"

        Celda = Celda & ofv.gencelda("","cc","center","2%","","","",sinput,"si")

        sInput = ofv.combosino(sForm & "A",rs(ii + 5),rs(ii + 5).name)

        Celda = Celda & ofv.gencelda("","vt","","5%","","","",sInput,"si")


    next

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","aaa colspan=11","","","","","","&nbsp;","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & rs(0)
      sAccion = sAccion & "&tabla=Tipobjeto&volver=" & vuelta
    end if

    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","tbp","","30% ","","0","","-1","2","","",Formulario,"si")

    posi = posi + 1
  end if



  if wNivel <> "NO" then

  call ofv.cerrarconsulta(rs)

    'Tipo
     Celda = ofv.GenCelda("","ttT colspan=3 ","center","100%","","","","Tipos de esquemas con nivel: " & wNivel,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")

    'Tipo
     Celda = ofv.GenCelda("","tt","","10%","","","","Tipo","si")

    'Descripcion
     Celda = Celda & ofv.gencelda("","tt","","90%","","","","Descripcion","si")

     Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

     Strsql = "Select * From Tipobjeto where nivel = '" & wNivel & "' order by descri "
  '''response.write strsql
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     if not rs.eof then

        do until rs.eof
           'Tipo
           sInput = rs(0)
           Celda =  ofv.gencelda("","cc","","","","","",sInput,"si")

          'Descripcion
           sInput = rs(1)
           Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")
           Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

           rs.movenext
        loop

    else

           sInput = "No hay esquemas para este nivel."
           Celda = ofv.gencelda("","cc colspan=3 ","","","","","",sInput,"si")
           Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    end if



    sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","2","","",Fila,"si") & "<br>"



  end if


  Agregar = ""
  Buscar  = ""

  sAccion = "TiposdeEsquema.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & table

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>