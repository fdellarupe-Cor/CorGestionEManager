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

  vuelta = "../genCMDdoc.asp?id=" & ID



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

    strsql= "insert into OBJOBSCMT (codsol,codobjeto,tipobjeto,codusr,codinx1,codinx2,codinx3,codcircuito,fecha,hora,TIPO) values('" & session("codSol") & "','" & nCodigo & "','" & sTipo & "','" & session("usuario") & "','" & ID2 & "','" & id & "','" & id3 & "','" & idcircuito & "','" & vFECHAHOY & "','" & vHORAHOY & "','CMD')"
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

    sImage = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0120;</font>"
    sImagea = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0121;</font>"

    sImaged = "<font face='corflaticon' size=2 color='#f0f0f0' style='background-color:#993333;border-style:OUTSET;border-width:2;' >&#0214;</font>"

    sImageSH = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0244;</font>"
    sImageSW = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0243;</font>"
    sImageSS = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0113;</font>"


    wkproccod = "<a class=nt  href='javascript:void(0);' >"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</a>"

    wkproccodA = "<a class=nt target='WK4' href='genCMDdoc.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=INS' >"
    wkproccodA = wkproccodA & sImagea
    wkproccodA = wkproccodA & "</a>"

    wkproccodSH = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSH = wkproccodSH & sImageSH
    wkproccodSH = wkproccodSH & "</a>"


    wkproccodSw = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSw = wkproccodSw & sImageSw
    wkproccodSw = wkproccodSw & "</a>"

    wkproccodSS = "<a class=nt  href='javascript:void(0);' >"
    wkproccodSS = wkproccodSS & sImageSS
    wkproccodSS = wkproccodSS & "</a>"



    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","AGREGAR/CONSULTAR COMENTARIOS","si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ventana normal.'  onclick='GTCOBSSS();' style='background-color:#6699cc;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSS,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar alto.'  onclick='GTCOBSHM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSH,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar ancho.'  onclick='GTCOBSWM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSw,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Agregar observaciones.' style='background-color:#0000bb;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccoda,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Cerrar.' onclick='GTCOBSCL()' style='background-color:#bb0000;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccod,"si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = shtml & sFORMhed


     fila=""







  strsql = "SELECT * from OBJOBSCMT where TIPO='CMD' and  codsol='" &session("codSol") &"' AND codcircuito = '"& idcircuito &"'  and codinx2 = " & id & " and "
  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and "
  strsql = strsql & " codusr = '" & session("usuario") & "' "
  strsql = strsql & " order by objid"
'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



   if  rs.eof then


   ' Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
   ' Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")
    'INFO="<big>No hay observaciones para el elemento.</big>"

    'Celda = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","",INFO,"si")
    'Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")





   end if

  iform=0

  do while  not rs.eof

    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")


      stylob  = ""
      sform="modtxt" & iform




      'Texto

    sInput = "<font class=at><textarea class=dt name=" & rs(13).name & "  id=" & rs(13).name & " rows=4 cols=70 "
    if bModificar  then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
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
      if bModificar  and rs("actpend")& "" = "" then



        wkproccodA = "<a class=nt target='WK4' href='genCMDdoc.asp?id=" & ID  & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OP=DEL" & "&objid=" & rs(0) & "' ><b>"
        wkproccodA = wkproccodA & sImaged
        wkproccodA = wkproccodA & "</b></a>"

	    Celda = ofv.GenCelda("","ut style='border-style:solid;border-color:#888888;border-width:1;' title='Eliminar observacion.' ","center","2%","","","",wkproccodA,"si")

      else
	     Celda = ofv.GenCelda("","ut","center","2%","","","","","si")
     end if


    Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput & sinput3 & sinput1 & sinput2 ,"si")





    if bModificar and rs("actpend")& "" = "" then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0)& "&Tabla=OBJOBSCMT&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"WK4","post","",celda,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",celda,"si")
      end if



      Fila  = fila & ofv.GenRow("","aaa bgcolor='f7eedd' ","","","","","",Formulario,"si")






rs.movenext
iform=iform+1
loop


 shtml =shtml & ofv.GenTabla("","","","80%  border ","","0","","0","0","","",fila,"si")



  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml



%>