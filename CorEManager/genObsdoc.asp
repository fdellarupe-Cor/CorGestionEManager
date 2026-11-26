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
  if session("Nomodificar") = "S" then bModificar = False

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



  sform = "modtxt"

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

  vuelta = "../genObsdoc.asp?id=" & ID



  eststt = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  estst1 = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;} "
  estst2 = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#88bbee;} "
  estst3 = " ;font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#77aadd;} "
  estst4 = " ;font-family:tahoma,arial;color:#000000;font-size:8pt;background-color:#ffffff;} "
  estst5 = " ;font-family:tahoma,arial;color:#eeeeee;font-size:12pt;background-color:#5588bb;} "
  esttxt = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#ffffff;border-style:outset;border-width:2;} "

  'f7eedd

 ''response.write ofv.FormHeader(session("FormN"))


  set cn = ofv.conectar(ofv.strconn2)





 ' response.write strsql

 'busco el circuito

   set cnx = ofv.conectar(ofv.strconn5)

   strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
   set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)
   if not rss.eof then idcircuito=rss("segckt")

   call ofv.cerrarconsulta(rss)
   call ofv.cerrarconn(cnx)



'''cstr(session("codSol")&"")


'   CASE "2"   'CONSIDERACIONES

'   CASE "3"   'SECTORES

'   CASE "5"   'PROCEDIMIENTOS

'   CASE "7"   'CONTROLES

'   CASE "8"   'ESQUEMAS

'   CASE "9"   'TABLAS

'   CASE "10" 'documentos asociados

'   CASE "11" 'CONTENIDO




   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
   xxxx      = GETFECH()
   vHORAHOY  = XLHORDAT



  if OP = "INS" then

    strsql= "insert into OBJOBSCMT (codsol,codobjeto,tipobjeto,codusr,codinx1,codinx2,codinx3,codcircuito,fecha,hora,TIPO) values('" & session("codSol") & "','" & nCodigo & "','" & sTipo & "','" & session("usuario") & "','" & ID2 & "','" & id & "','" & id3 & "','" & idcircuito & "','" & vFECHAHOY & "','" & vHORAHOY & "','OBS')"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if

  if OP = "DEL" then

    strsql= "delete from OBJOBSCMT where objid = 0" & sobjid
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if


  'response.write "<center>"

    Fila = ""

    wkproccod = ""
    wkproccodA = ""
    wkproccodAD = ""

    sImage = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0114;</font>"
    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#078;</font>"

    sImaged = "<font face='corflaticon' size=2 color='#f0f0f0' style='background-color:#993333;border-style:OUTSET;border-width:2;' >&#114;</font>"

    sImageSH = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#053;</font>"
    sImageSW = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#056;</font>"
    sImageSS = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#099;</font>"


    wkproccod = "<a class=nt  href='javascript:void(0);' >"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</a>"

    wkproccodAD = "<a class=nt target='WK4' href='genObsdoc.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=INS' >"
    wkproccodAD = wkproccodAD & sImagea
    wkproccodAD = wkproccodAD & "</a>"

    wkproccodSH = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSH = wkproccodSH & sImageSH
    wkproccodSH = wkproccodSH & "</a>"


    wkproccodSw = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSw = wkproccodSw & sImageSw
    wkproccodSw = wkproccodSw & "</a>"

    wkproccodSS = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSS = wkproccodSS & sImageSS
    wkproccodSS = wkproccodSS & "</a>"





     fila=""







  strsql = "          SELECT * from OBJOBSCMT where TIPO='OBS' AND codcircuito = '"& idcircuito &"'  and codinx2 = " & id & " and "
  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and "
  strsql = strsql & " codusr = '" & session("usuario") & "' and codsol= 0"  & session("codSol")
  strsql = strsql & " order by objid"
'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



   if  rs.eof then


   ' Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
   ' Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")
    'INFO="<big>No hay observaciones para el elemento.</big>"

    'Celda = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","",INFO,"si")
    'Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")
    shtml= shtml & "<script language=JavaScript>var et=window.top.getIFD('Principal').getElementById('mxbtrobs'); if (et  != null)  {et.style.display="&chr(34)&"none"&chr(34)&";}</script>"
    SESSION("WKFOBSBTN")="S2"

   else
    shtml= shtml & "<script language=JavaScript>var et=window.top.getIFD('Principal').getElementById('mxbtrobs'); if (et  != null)  {et.style.display="&chr(34)&""&chr(34)&";}</script>"
    SESSION("WKFOBSBTN")="S1"

   end if

  iform=0


  obsvacio = false

  do while  not rs.eof

      if trim(rs(13) & "" ) ="" and obsvacio=false then obsvacio = true end if

      'response.write session("obsvacio")
    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")


      stylob  = ""
      sform="modtxt" & iform




      'Texto

    sInput = "<font class=at><textarea class=dt name=" & rs(13).name & "  id=" & rs(13).name & " rows=4 cols=70 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else { "
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(13)&"")  &"</textarea></font>"


    sinput3=""

    sinput1 ="<input type=hidden name=fecha value=" & vFECHAHOY &">"
    sinput2 ="<input type=hidden name=hora  value="&vHORAHOY&">"

    'Eliminar
      sHref = ""
      if bModificar  then



        wkproccodA = "<a class=nt target='WK4' href='genObsdoc.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=DEL" & "&objid=" & rs(0) & "' ><b>"
        wkproccodA = wkproccodA & sImaged
        wkproccodA = wkproccodA & "</b></a>"

	    Celda = ofv.GenCelda("","ut style='border-style:solid;border-color:#888888;border-width:1;' title='Eliminar observacion.' ","center","2%","","","",wkproccodA,"si")

      else
	     Celda = ofv.GenCelda("","ut","center","2%","","","","","si")
     end if


  '  Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput & sinput3 & sinput1 & sinput2 ,"si")






    if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0)& "&Tabla=OBJOBSCMT&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"WK4","post","",sInput,"si")
        'Formulario = ofv.GenForm(sForm,"",sAccion,"WK4","post","",celda,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",sInput,"si")
        'Formulario = ofv.GenForm("","","","","post","",celda,"si")
      end if

      Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",Formulario,"si")

      Fila  = fila & ofv.GenRow("","aaa bgcolor='f7eedd' ","","","","","",celda,"si")


     '' Fila  = fila & ofv.GenRow("","aaa bgcolor='f7eedd' ","","","","","",Formulario,"si")






rs.movenext
iform=iform+1
loop
  

    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","AGREGAR/CONSULTAR OBSERVACIONES","si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ventana normal.'  onclick='GTCOBSSS();' style='background-color:#6699cc;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSS,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar alto.'  onclick='GTCOBSHM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSH,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar ancho.'  onclick='GTCOBSWM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSw,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Agregar observaciones.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodAD,"si")

    if obsvacio = false then 
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Cerrar.' onclick=""GTCOBSCLEX('S');"" style='background-color:#bb0000;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccod,"si")

    else 
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Cerrar.' onclick=""GTCOBSCLEX('N');"" style='background-color:#bb0000;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccod,"si")
    end if 
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = sFORMhed & shtml 

 shtml =shtml & ofv.GenTabla("","","","80%  border ","","0","","0","0","","",fila,"si")



  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)





  sHTML = sHTML & "</BODY>" 

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml


%>