<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  'response.write session("codSol")
  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
 ' if session("Nomodificar") = "S" then bModificar = False

  sHTML = sHTML & "<HTML>"


 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN
    sHTML = sHTML & "<BODY onload='window.top.LDFOBJRSK();' >"
 ELSE
    sHTML = sHTML & "<BODY onload='window.top.ldfobj();' >"
 END IF


  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")

  OP      = request.querystring("OP")
  sobjid  = request.querystring("objid")

  ID      = request.querystring("ID")
  ID2     = request.querystring("ID2") & ""
  ID3     = request.querystring("ID3") & ""


'response.write request.querystring
  sform = "modtxt"
  sFormX = "filtroCombo"

  wposi = 0
  wpossal = 0
  if id2 <> "" then
     session("id2")=id2 & ""
  else
     id2=session("id2") & ""
  end if

  if id3 <> "" then
     session("id3")=id3 & ""
  else
     id3=session("id3") & ""
  end if

  vuelta = "../genCOMENTARIO.asp?id=" & ID



  eststt = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  estst1 = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;} "
  estst2 = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#88bbee;} "
  estst3 = " ;font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#77aadd;} "
  estst4 = " ;font-family:tahoma,arial;color:#000000;font-size:8pt;background-color:#ffffff;} "
  estst5 = " ;font-family:tahoma,arial;color:#eeeeee;font-size:12pt;background-color:#5588bb;} "
  esttxt = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#ffffff;border-style:outset;border-width:2;} "


  set cn = ofv.conectar(ofv.strconn2)


   set cnx = ofv.conectar(ofv.strconn5)

   strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
   set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
   if not rss.eof then idcircuito=rss("segckt")

   call ofv.cerrarconsulta(rss)
   call ofv.cerrarconn(cnx)




   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
   xxxx      = GETFECH()
   vHORAHOY  = XLHORDAT



  if OP = "INS" then

    strsql= "insert into OBJOBSCMT (codsol,codobjeto,tipobjeto,codusr,codinx1,codinx2,codinx3,codcircuito,fecha,hora,TIPO) values('9999','" & nCodigo & "','" & sTipo & "','" & session("usuario") & "','" & ID2 & "','" & id & "','" & id3 & "','" & idcircuito & "','" & vFECHAHOY & "','" & vHORAHOY & "','CMD')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if

  if OP = "DEL" then

    strsql= "delete from OBJOBSCMT where objid = 0" & sobjid
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if


    Fila = ""

    wkproccod = ""
    wkproccodA = ""

    sImage = "<font face='corflaticon' size=5 color='#000000'  >&#0114;</font>"
    'sImagea = "<font face='corflaticon' size=5 color='black'  >)</font>"
	sImagex = "<font face='corflaticon' size=5 color='#000000'  >&#0156;</font>"
    sImaged = "<font face='corflaticon' size=5 color='#f0f0f0' style='background-color:#993333;border-style:OUTSET;border-width:2;' >&#114;</font>"


    sImageSW = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0161;</font>"
    sImageSS = "<font face='corflaticon' size=5 color='#000000'  >x</font>"


	'BOTON CERRAR 'X'
    wkproccod = "<a class=nt  href='javascript:void(0);' >"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</a>"

	'BOTON AGREGAR COMENTARIO
    wkproccodA = "<a class=nt target='WK4' href='genCOMENTARIO.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=INS' >"
    wkproccodA = wkproccodA & sImagex
    wkproccodA = wkproccodA & "</a>"

    'FILTRO PROPIOS
	mios = "<a class=nt target='WK4' href='genCOMENTARIO.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=PROPIOS' >"
	mios = mios & "<b><i>VER PROPIOS</i></b>"
    mios = mios & "</a>"

    'FILTRO TODOS
	todos = "<a class=nt target='WK4' href='genCOMENTARIO.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=TODOS' >"
	todos = todos & "<b><i>VER TODOS</i></b>"
    todos = todos & "</a>"



if OP="TODOS" then
''TODOS
 strsql = "SELECT * from OBJOBSCMT where TIPO='CMD' and  codsol='9999' AND codcircuito = '"& idcircuito &"'  and codinx2 = '" & id & "' and "
 strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' order by fecha desc,hora desc"
else
'MIOS
 strsql = "SELECT * from OBJOBSCMT where TIPO='CMD' and  codsol='9999' AND codcircuito = '"& idcircuito &"'  and codinx2 = '" & id & "' and "
 strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and codusr = '" & session("usuario") & "' order by fecha desc,hora desc"
end if

'response.write strsql

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  'CODIGO ANTERIOR
  'sql = "select * from objsdkdet where codobjeto='0' and tipobjeto='" & stipo & "' and tiposeg='9' and grpusr='USR-" & session("usuario") & "'"
  'response.write sql

  'CODIGO NUEVO
  'INC 446: No se estaba teniendo en cuenta el grupo del usuario logueado en la seguridad, se agrega el grupo como autorizante del tipo de esquema
  sql = "select * from objsdkdet where codobjeto='0' and tipobjeto='" & stipo & "' and tiposeg='9'"
  sql = sql & " and (grpusr='USR-" & session("usuario") & "' or grpusr='GRP-" & session("SegGrp") & "')"
  'response.write sql

  set usr = ofv.crearconsultaex(sql,cn,1,parametros)
  sql1= "select autor from objprty where codobjeto=" & ncodigo & " and autor='" & session("usuario") & "'"
  'response.write sql1
  set aut = ofv.crearconsultaex(sql1,cn,1,parametros)

    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","uti style='background-color:#000077;color:#ffffff;' ","","100%","","","","<b><i><h3>AGREGAR/CONSULTAR COMENTARIOS</h3></i></b>","si")
if (not usr.eof) or (not aut.eof) then
    if OP="PROPIOS" or OP="" then
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Filtrar comentarios' style='background-color:#87CEFA;border-style:OUTSET;border-width:1;'","center","2%","","","",todos,"si")
    else
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Filtrar comentarios' style='background-color:#87CEFA;border-style:OUTSET;border-width:1;'","center","2%","","","",mios,"si")
    end if
end if
    if not rs.eof then
    if trim(rs(13) & "") <> "" then
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Agregar comentarios' style='background-color:#ADD8E6;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
    else
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Agregar comentarios' style='background-color:#ADD8E6;border-style:OUTSET;border-width:1;' ","center","2%","","","",sImageSS,"si")
    end if
    else
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Agregar comentarios' style='background-color:#D8D8D8;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
    end if
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Cerrar.' onclick='GTCOBSCL()' style='background-color:#FF0000;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccod,"si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = shtml & sFormhed

     fila=""





  color = "#FFFFFF"
  iform=0

  do while  not rs.eof

rechazado=""

if instr(rs(1) & "" ,"<RCH>") > 0 then rechazado = "(Rechazado)"

xxxx     = getInfoUser(rs(9))
xusrlog  = session("SGU_nombre")


    Celda = ofv.GenCelda("","ut style='color:#000000;' align=center","","","","","","<b><i>" & (iform+1) & "</i></b>","si")
	Celda = Celda & ofv.GenCelda("","ut style='color:#000000;' align=left","","","","","","<i><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & xusrlog & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Fecha: " & GTFechadorsh(rs(10)) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & " Hora: " & GETGLBHMSW(CLNG(rs(11))) & "</i></b>","si")
    Fila = fila & ofv.genrow("","a style='background-color:" & color & ";'","","","20","","",Celda,"si")


      stylob  = ""
      sform="modtxt" & iform

      'Texto

    sInput = "<font class=at><textarea style='background-color:" & color & ";border-color:#6E6E6E;color:#000000;' class=dt name=" & rs(13).name & "  id=" & rs(13).name & " rows=4 cols=100 "
	
'sHTML = sHTML & "<html><script languaje='javascript'>alert('" & chr(34) & session("Usuario") & chr(34) & "')</SCRIPT></html>"


    if  rs("actpend")& "" = "" and rs(13)& "" = "" then
    	if rs("CODUSR")=session("Usuario") then
       boton = "<input type=button class=BT value=Guardar onclick=""if (this.value.length > 2300) "
			if session ("APLMDCLI") = "BGAL" then
				boton = boton & " {alert('El texto es demasiado extenso'); return false;} else {document." & sform & ".submit() ;alert('Gracias por dejar comentarios! El sector CLARIDAD responder&aacute a la brevedad sus inquietudes.');}"">"
				else
					if session ("APLMDCLI") <> "BGAL" then
					boton = boton & " {alert('El texto es demasiado extenso'); return false;} else {document." & sform & ".submit() ;alert('Gracias por dejar comentarios! Ser&aacute;n respondidos a la brevedad.');}"">"
		else
	   boton = boton & " {alert('El texto es demasiado extenso'); return false;} else {document." & sform & ".submit() ;alert('Mensaje guardado con exito');}"">"
	   end if
	   end if
	   CeldaX = ofv.GenCelda("","ut colspan=2","center","","","","",boton,"si")
	   else
	   CeldaX = ofv.GenCelda("","ut colspan=2","center","","","","","","si")
	   end if
    elseif rs("actpend")& "" = "" then
    if rs("CODUSR")=session("Usuario") then
       boton = "<input type=button class=BT value=Actualizar onclick=""if (this.value.length > 2300) "
       boton = boton & " {alert('El texto es demasiado extenso'); return false;} else {document." & sform & ".submit() ;alert('Mensaje actualizado con exito!');}"">"
	   CeldaX = ofv.GenCelda("","ut colspan=2","center","","","","",boton,"si")
	   else
	   CeldaX = ofv.GenCelda("","ut colspan=2","center","","","","","&nbsp;","si")
	   end if
    else
    	xxxx1     = getInfoUser(rs(15))
		xusrlog1  = session("SGU_nombre")
       CeldaX = ofv.GenCelda("","ut style='color:#000000;' colspan=2","right","","","","","<i>Confirmado: " & xusrlog1 & "&nbsp;&nbsp;&nbsp;&nbsp;Fecha: " & GTFechadorsh(rs(16)) & "&nbsp;&nbsp;&nbsp;&nbsp;Hora: " & GETGLBHMSW(CLNG(rs(17))) &"</i>","si")
    end if

    if rs("actpend")& "" = "" and rs("CODUSR")=session("Usuario") then
    sInput = sInput & " >" &  trim(rs(13)&"")  &"</textarea></font>"
    else
    sInput = sInput & " readonly>" &  trim(rs(13)&"")  &"</textarea></font>"
    end if

    sinput3=""
  if rs("actpend") & "" = "" then
    sinput1 ="<input style='display: none' type=hidden name=fecha value=" & vFECHAHOY & ">"
    sinput2 ="<input style='display: none' type=hidden name=hora  value=" & vHORAHOY & ">"
 end if


    'Eliminar
      sHref = ""

      if rs("actpend")& "" <> "" then
      sImageSH = "<font face='corflaticon' size=6 color='#04B404'>a</font>"
      else
      sImageSH = "<font face='corflaticon' size=5 color='#B40404'>=</font>"
      end if

      if rs("actpend")& "" = "" and rs("CODUSR")=session("Usuario") then

        wkproccodA = "<a class=nt target='WK4' href='genCOMENTARIO.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=DEL" & "&objid=" & rs(0) & "' ><b>"
        wkproccodA = wkproccodA & sImaged
        wkproccodA = wkproccodA & "</b></a>"
	    Celda =   ofv.GenCelda("","ut style='border-style:solid;border-color:#888888;border-width:1;' title='Eliminar COMENTARIO.' ","center","2%","","","",wkproccodA,"si")

      else
	     Celda =   ofv.GenCelda("","ut","center","2%","","","",sImageSH,"si")
     end if



    if  rs("actpend")& "" = "" then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0)& "&Tabla=OBJOBSCMT&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"WK4","post","",sInput & sinput3 & sinput1 & sinput2 ,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",sInput & sinput3 & sinput1 & sinput2 ,"si")
      end if


      celda =celda &  ofv.GenCelda("","","","","","","",Formulario,"si")
      Fila  = fila & ofv.GenRow("","a style='background-color:" & color & ";'","","","","","",celda,"si")
	  Fila = Fila & ofv.genrow("","a style='background-color:" & color & ";'","","","","","",CeldaX,"si")
      celdas = ofv.GenCelda("","a style='background-color:#000077;' colspan=2 align=center","","","","","","<hr>","si")
      Fila  = fila & ofv.GenRow("","a style='background-color:" & color & ";'","","","","","",celdas,"si")

	if color = "#FFFFFF" then
	color = "#E6E6E6"
	else
	color = "#FFFFFF"
	end if



rs.movenext
iform=iform+1
loop


shtml =shtml & ofv.GenTabla("","aaa",""," ","","0","","0","0","","",fila,"si")
sHTML = sHTML & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", Agregar, Retornar, "")


  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml


Function GTFechadorsh(GTWfecha)

      IF CSTR(GTWfecha & "") = "0" OR CSTR(GTWfecha & "") = "" THEN

         GTFechadorSH  = "-----"

      ELSE


         GTwfec = cdate(GTWfecha + clng(cdate("01-01-1900")))

         GTWFDIAB = CSTR(Day(GTwfec))
         GTWFMESB = CSTR(Month(GTwfec))

         IF LEN(GTWFDIAB) = 1 THEN GTWFDIAB = "0" & GTWFDIAB
         IF LEN(GTWFMESB) = 1 THEN GTWFMESB = "0" & GTWFMESB

         GTwFechash = GTWFDIAB & "/" & GTWFMESB & "/" & Year(GTwfec)
         GTFechadorSH  = ucase(GTwFechash)

      END IF

End Function

function GETGLBHMSW(XITIME)

 if not isnumeric(XITIME & "") then

 GETGLBHMSW = "000005"

 else

XITIME=clng("0" & replace(cstr(XITIME & ""),"-",""))
GLBbb = XITIME MOD 3600
GLBbC = (GLBBB) MOD 60
GLBbD = (GLBBC) MOD 1

GLBSH = (XITIME - GLBBB) / 3600
GLBSM = (GLBBB - GLBBC) / 60
GLBSS = GLBBC / 1

GLBSH = cstr(GLBSH)
GLBSM = cstr(GLBSM)
GLBSS = cstr(GLBSS)

if len(GLBSH) = 1 then GLBSH = "0" & GLBSH
if len(GLBSM) = 1 then GLBSM = "0" & GLBSM
if len(GLBSS) = 1 then GLBSS = "0" & GLBSS

 XOSH = GLBSH & ":"
 XOSM = GLBSM & ":"
 XOSS = GLBSS & ""


 GETGLBHMSW = XOSH & XOSM & XOSS

 end if

END FUNCTION
%>