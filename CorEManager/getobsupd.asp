<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  sHTML = sHTML & "<HTML>"

 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN
    sHTML = sHTML & "<BODY onload='window.top.LDFOBJRSK();' >"
 ELSE
    sHTML = sHTML & "<BODY onload='window.top.ldfobj();' >"
 END IF



  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  IDTXT   = request.querystring("IDTXT")
  vuelve  = request.querystring("vuelve")
  sOPcion = request.querystring("OPcion")
  sObjid  = request.querystring("objid")
  wasotxt = request.querystring("wasotxt")
  ID      = request.querystring("ID")
  ID2     = request.querystring("ID2") & ""
  ID3     = request.querystring("ID3") & ""

  if wasotxt<> "" then
   session("wasotxt")=wasotxt
  else
   wasotxt=session("wasotxt")
  end if



   vuelve=replace(vuelve,"_","&")


    if vuelve <> "" then
     session("vuelve")=vuelve
    else
     vuelve=session("vuelve")
    end if

   'response.write request.querystring







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

  vuelta = "../genOBS.asp?id=" & ID  '& "&" &volver

 ' response.write request.querystring

  eststt = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  estst1 = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;} "
  estst2 = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#88bbee;} "
  estst3 = " ;font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#77aadd;} "
  estst4 = " ;font-family:tahoma,arial;color:#000000;font-size:8pt;background-color:#ffffff;} "
  estst5 = " ;font-family:tahoma,arial;color:#eeeeee;font-size:12pt;background-color:#5588bb;} "
  esttxt = " ;font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#ffffff;border-style:outset;border-width:2;} "

  'f7eedd

''  response.write ofv.FormHeader(session("FormN"))


  set cn = ofv.conectar(ofv.strconn2)


  if sopcion = "SI" then

  ''   response.write "entro aca"

    strsql = "update OBJOBSCMT set actpend = '1' where objid = 0" & sobjid
'' response.write strsql
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)




  end if


  if sopcion = "NO" then

    ''   response.write "entro aca"

      strsql = "update OBJOBSCMT set actpend = '2' where objid = 0" & sobjid
  '' response.write strsql
      call ofv.crearconsultaEx(StrSql,cn,1,parametros)




  end if



' if session("ModMain")= "CL" then

' strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
' else
  strsql = "SELECT  TOP 1 * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
'end if


 'busco el circuito

   set cnx = ofv.conectar(ofv.strconn5)

   set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)

  'if not rss.eof then idcircuito=rss("segckt")

  ' call ofv.cerrarconsulta(rss)
  'call ofv.cerrarconn(cnx)

  do while   not rss.eof

  idcircuito=rss("segckt")
   IF wasotxt = "S" THEN

    stylob = "UT valign=middle align=center style='border-style:solid;border-width:1;border-color:#990000;' "

   end if



     sImageOK = "<font face='corflaticon' color='#FFFFFF' size=3 >&#097;</font>"


 '  shtml = "<div id='mt2'  style='{height:100px;overflow:auto;font-family:tahoma;font-size:8pt;color:#bbbbbb;border-style:solid;border-color:#993333;border-width:1;' >"
 '  shtml = shtml & "<br>&nbsp;" & wkdatoobs & "</div><br>"

    Fila = ""

    wkproccod = ""

    sImage = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0114;</font>"

    sImageSH = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0244;</font>"
    sImageSW = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0243;</font>"
    sImageSS = "<font face='corflaticon' size=2 color='#f0f0f0'  >&#0113;</font>"

    wkproccod = "<a class=nt  href='javascript:void(0);' >"
    wkproccod = wkproccod & sImage
    wkproccod = wkproccod & "</a>"

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
    CeldaD = CeldaD & ofv.GenCelda("","utI  ","","95%","","","","OBSERVACIONES","si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ventana normal.'  onclick='GTCOBSSS();' style='background-color:#6699cc;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSS,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar alto.'  onclick='GTCOBSHM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSH,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Ampliar ancho.'  onclick='GTCOBSWM();' style='background-color:#f7aa22;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccodSw,"si")
    CeldaD = CeldaD & ofv.GenCelda("","ut title='Cerrar.' onclick='GTCOBSCL()' style='background-color:#bb0000;border-style:OUTSET;border-width:1;' ","center","2%","","","",wkproccod,"si")
    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")
    sFORMhed = ofv.GenTabla("","TWIND","","100%","","0","","0","1","","",FilaD,"si")


   shtml = shtml & sFORMhed






   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
   xxxx      = GETFECH()
   vHORAHOY  = XLHORDAT


  strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codinx2 = " & id & " and "
  'strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and "
  'strsql = strsql & " actpend is null "

  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"' and "
  strsql = strsql & "  and codcircuito='" & idcircuito & "'  and codsol= 0"  & session("codSol")
  'strsql = strsql & " actpend is null "

  strsql = strsql & " order by codcircuito, codusr, objid"
''  response.write strsql & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then


  iform=0

  do while  not rs.eof

    Celda = ofv.GenCelda("","ut colspan=4 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

        xxxx = getInfoUser(RS(9))
	if session("SGU_nombre")<> "" then    sInput =  session("SGU_nombre")

   celda = ""

   sInput = sImageOK

   sAccion = "getObsupd.asp?wasotxt=N&Codobjeto=" & nCodigo
   sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & id  & "&ID2=" & id2  & "&opcion=SI&objid=" & rs(0)

   sAccion2 = "getObsupd.asp?wasotxt=N&Codobjeto=" & nCodigo
   sAccion2 = sAccion2 & "&Tipobjeto=" & sTipo & "&ID=" & id  & "&ID2=" & id2  & "&opcion=NO&objid=" & rs(0)

   sHref   = sAccion
   sHref2   = sAccion2

 IF WKAPLIC = "OR" OR WKAPLIC = "PL" THEN
   Celdaobs  = ofv.GenLink("","ut",sHref,"fesquemas","","",sInput,"si")
   Celdaobs2  = ofv.GenLink("","ut",sHref2,"fesquemas","","","" & sInput,"si")
 ELSE
   Celdaobs  = ofv.GenLink("","ut",sHref,"WK4","","","" & sInput,"si")
   Celdaobs2  = ofv.GenLink("","ut",sHref2,"WK4","","","" & sInput,"si")
 END IF

 select case rs("actpenD")

 case 1 ' si
     borderSi="border-style:outset ;border-color: #000000;border-width:2px;"
     borderNo="border-style:none;"

 case 2 'no
     borderNO="border-style:outset ;border-color: #000000;border-width:2px"
     borderSI="border-style:none;"

 case else
     borderSi="border-style:none;"
     borderNo="border-style:none;"
 end select


   if session("ModMain") <> "CL" THEN
    Celda =  celda &  ofv.GenCelda("","ut ROWSPAN=2 BGCOLOR='#00AA00' TITLE='Confirmar observacion.' style='cursor:pointer;"&borderSi&"' ","CENTER","2%","","","",Celdaobs,"si")
    Celda =  celda &  ofv.GenCelda("","ut ROWSPAN=2 BGCOLOR='#ff0000' TITLE='Cancela observacion.' style='cursor:pointer;"&borderNo&"' ","CENTER","2%","","","",Celdaobs2,"si")
   END IF
 '  end if
    'Celda =  celda &  ofv.GenCelda("","ut  BGCOLOR='#00AA00' TITLE='Confirmar observacion.' style='cursor:pointer;' ","CENTER","2%","","","",Celdaobs,"si")
   sInput = RS(9) & "&nbsp;-&nbsp;" & session("SGU_nombre")

   Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","60%","","","",sInput,"si")

   sInput = GTFechadorsh(rs(10)) & "&nbsp;-&nbsp;" & GETGLBHMSW(rs(11))

   Celda =  celda &  ofv.GenCelda("","uti style='border-style:solid;border-color:#888888;border-width:1;' ","","40%","","","",sInput,"si")



   Fila  = fila & ofv.GenRow("","aaa bgcolor='#F7EEDD' ","","","","","",Celda,"si")


   'Texto


    sInput = trim(rs(13)&"")

    if sinput = "" then sinput = "-----"



   if wasotxt <> "N" then

    if rs(12) then
        sinput3= ofv.combosino(sForm,"1","actpend")
     else
       sinput3= ofv.combosino(sForm,"0","actpend")
    end if

    sinput3 = REPLACE(sinput3,"all.actpend",sform & ".all.actpend")
  else
    sinput3=""
  end if




   celda = ""
   Celda =  celda &  ofv.GenCelda("","uti COLSPAN=2 style='border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput & sinput3 ,"si")

   Fila  = fila & ofv.GenRow("","aaa bgcolor='#F5F5F5'  ","","","35","","",Celda,"si")




rs.movenext
iform=iform+1
loop


else

    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

    Celda = ofv.GenCelda("","uti colspan=2 style='border-style:solid;border-color:#888888;border-width:1;'  ","center","","","","","<big>No hay observaciones para el elemento.</big>","si")
    Fila = fila & ofv.genrow("","","","","50","","",Celda,"si")


end if

    Celda = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    Fila = fila & ofv.genrow("","","","","20","","",Celda,"si")

 shtml =shtml & ofv.GenTabla("","aaa  ","","98% ","","0","","1","1","","",fila,"si")



  Call ofv.cerrarconsulta(rs)

  rss.movenext
loop
call ofv.cerrarconsulta(rss)
call ofv.cerrarconn(cnx)

CAll ofv.cerrarconn(cn)


  sHTML = sHTML & "</BODY>"

  sHTML = sHTML & "</HTML>"
end if

 response.write shtml



%>