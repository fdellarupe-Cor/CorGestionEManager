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
  if session("Nomodificar") = "S" then bModificar = False


  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  IDTXT = request.querystring("IDTXT")
  vuelve = request.querystring("vuelve")

  vuelve=replace(vuelve,"_","&")


  if vuelve<> "" then
   session("vuelve")=vuelve
  else
   vuelve=session("vuelve")
   end if

   nCodigo = request.querystring("Codobjeto")
   sTipo   = request.querystring("Tipobjeto")
   sEst    = request.querystring("Tipo")
   sPart   = request.querystring("Part")
   ID      = request.querystring("ID")
   ID2      = request.querystring("ID2")

  response.write request.querystring
  sform="modtxt"

  wposi = 0
  wpossal = 0
  if id2 <> "" then
   session("id2")=id2
  else
  id2=session("id2")
  end if

  vuelta = "../edestrPRTYOBS.asp?id=" & ID  '& "&" &volver



  eststt = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  estst1 = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;} "
  estst2 = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#88bbee;} "
  estst3 = " STYLE={font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#77aadd;} "
  estst4 = " STYLE={font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#6699cc;} "
  estst5 = " STYLE={font-family:tahoma,arial;color:#eeeeee;font-size:12pt;background-color:#5588bb;} "
  esttxt = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#ffffff;border-style:outset;border-width:2;} "

'f7eedd

  response.write ofv.FormHeader(session("FormN"))


  set cn = ofv.conectar(ofv.strconn2)

  response.write "<center>"

  response.write "<TABLE     width=60%  marginwidth=0  cellspacing=0  cellpadding=5 >"



    'null
     sicK = "<font face='corflaticon' size=4 color=#993333 >&#0157;</font>&nbsp;&nbsp;"
     Celda =  ofv.GenCelda("","ttT valign=top  ","","","","","",sicK & "<b>Modificacion de Texto de contenido</b>","si")

     Fila = ofv.GenRow("","","","","","","",celda,"si")
     response.write Fila


    strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"
   ' response.write strsql

   'busco el circuito

     set cnx = ofv.conectar(ofv.strconn5)

     set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)

     if not rss.eof then idcircuito=rss("segckt")

     call ofv.cerrarconsulta(rss)
   call ofv.cerrarconn(cnx)




    strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codcircuito='"& idcircuito &"'  and codinx2=" & id & " and codobjeto = " & nCodigo & " and tipobjeto='"& sTipo &"' and codusr='" & session("usuario") & "'"
   ' response.write strsql


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if  not rs.eof then

      wasotxt = "N"
      stylob  = ""


      ' timez=clng(timeserial(hour(now),minute(now),second(now)))

       IF wasotxt = "S" THEN

          stylob = "UT valign=middle align=center style='border-style:solid;border-width:1;border-color:#990000;' "

       end if


      'Texto

    sInput = "<font class=at><textarea class=dt name=" & rs(13).name & "  id=" & rs(13).name & " rows=20 cols=120 "
    'if bModificar and wasotxt = "N" then
    '   sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
    '   sInput = sInput & " extenso'); return false;} else {"
    '   sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    'else
       sInput = sInput & " readonly "
    'end if
    sInput = sInput & " >" &  trim(rs(13)&"")  &"</textarea></font>"




      Celda =  ofv.GenCelda("",stylob,"","","","","",sInput & sinput1 & sinput2 & sinput3& sinput4,"si")




      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0)& "&Tabla=OBJOBSCMT&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",celda,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",celda,"si")
      end if



      Fila  = ofv.GenRow("","","","","","","",Formulario,"si")

      wstextoLNK = rs(9) & ""

       if  isnumeric(wstextoLNK) and wstextoLNK <> "0" then
           sImagenP = "<font face='corflaticon' size=2 color=#005500 >&#0252;</font>"
       else
           sImagenP = "<font face='corflaticon' size=2 color=#f7aa22 >&#0252;</font>"
       end if

       sHref = "javascript:void(0);"

      sID = "edestrdetLnk.asp?"
      sID = sID & "IDTXT=" & IDTXT & "&Tabla=OBJOBSCMT&volver=" & vuelta
      sID = sID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHref = sID

       sinputk = "<a class='ft1o'  href='" & sHref
       sinputk = sinputk & "' title='Actualizar link a otro documento' >" & sImagenP & "</a>"
       slink   = "&nbsp;" & sinputk

      Celda = ofv.GenCelda("","ft","left","5%","","","",sLink,"si")
 '     fila  = fila & ofv.GenRow("","aaa   ","","","","","",celda,"si")



            response.write Fila




end if


response.write ofv.GenFTabla()

  Agregar = ""
  Buscar = ""

  'Boton Volver
 ' sAccion   = right(volver,len(volver)-3) 'replace(request.querystring("volver"),"../","") '"edestrdet.asp?id="& id2
 ' sRetornar = ofv.BotonVolver(sAccion)
 ' if  volver <> "" then
 '      sAccion = sAccion & "&codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

   sAccion = right(session("vuelve") ,len(session("vuelve") )-3)& "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       sRetornar = ofv.BotonVolver(sAccion)
 ' end if
 'response.write "<br>" & saccion
  Response.write ofv.botoneraESSX("0",rs,cantregmover,cantregmostrar,"no","no","","","","",Agregar,Buscar,sRetornar)

  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

response.write  "</body></html>"
end if



%>