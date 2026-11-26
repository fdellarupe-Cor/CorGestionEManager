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


  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  IDTXT   = request.querystring("IDTXT")
  vuelve  = request.querystring("vuelve")
  OP      = request.querystring("OP")
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

  response.write ofv.FormHeader(session("FormN"))


  set cn = ofv.conectar(ofv.strconn2)

  'response.write "<center>"

  response.write "<TABLE     width=60%  marginwidth=0  cellspacing=0  cellpadding=5 >"



    'null
     sicK = "<font face='corflaticon' size=4 color=#993333 >&#0157;</font>&nbsp;&nbsp;"
     Celda =  ofv.GenCelda("","ttT valign=top  colspan=2 ","","","","","",sicK & "<b>Observacioes al documento</b>","si")

     Fila = ofv.GenRow("","","","","","","",celda,"si")
     response.write Fila & "</table>"

     fila=""




  strsql = "SELECT * from SEGCIRCUITO where objcod = '" & nCodigo  & "' and objtipo='" & sTipo & "' order by id desc"

 ' response.write strsql

 'busco el circuito

   set cnx = ofv.conectar(ofv.strconn5)

   set rss = ofv.crearconsultaEx(StrSql,cnx,1,parametros)

   if not rss.eof then idcircuito=rss("segckt")

   call ofv.cerrarconsulta(rss)
   call ofv.cerrarconn(cnx)

   IF wasotxt = "S" THEN

    stylob = "UT valign=middle align=center style='border-style:solid;border-width:1;border-color:#990000;' "

   end if




  SELECT CASE cstr(session("codSol")&"")

   CASE "2"   'CONSIDERACIONES



   CASE "5"   'PROCEDIMIENTOS

    strsql = "select * from OBJACCION "
    strsql = strsql & "WHERE OBJID = 0" & id

   ' response.write strsql
       set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not xrs.eof then

         'Acciones
         if Xrs(4) <> "" then
            Strsql = "Select descri,MODO From Tipoaccalt Where tipoaccionalt = " & Xrs(4)
            set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rsAcciones.eof then
                sAccionM = rsAcciones(0)
                sModo   = rsAcciones(1)
            end if
            call ofv.cerrarconsulta(rsAcciones)
         end if

         sVerbo   = sAccionM
         sAcciont = sAccionM
         i        = instr(sVerbo," ")

         if i > 0 then
            sVerbo   = left(sAccionM,i-1)
            sAcciont = mid(sAccionM,i+1)
         end if

         sInput = Xrs(10) & ".&nbsp;<i>" & sVerbo & "</i>" & "&nbsp;&nbsp;&nbsp;"
         sInput = sInput & sAcciont

          wkdatoobs = replace(sInput &"",chr(13),"<br>")


       end if



   CASE "7"   'CONTROLES

    Strsql = "SELECT * from Objcontrol2 "
    strsql = strsql & "WHERE OBJID = 0" & id

   ' response.write strsql
       set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not xrs.eof then



         sInput = Xrs(5) & ".&nbsp;" & Xrs(4)
         wkdatoobs = replace(sInput &"",chr(13),"<br>")


       end if



   CASE "8"   'ESQUEMAS


    Strsql = "SELECT OBJESQUEMA2.*, OBJCUENTA.descri, OBJCUENTA.numero, OBJCUENTA.objid "
    Strsql = Strsql & "FROM OBJESQUEMA2, OBJCUENTA where "
    Strsql = Strsql & "OBJESQUEMA2.codcuenta = OBJCUENTA.objid "
    strsql = strsql & " and OBJESQUEMA2.OBJID = 0" & id

   ' response.write strsql
       set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not xrs.eof then

  'Texto Sentido
      sInput = xrs(5)
      if  sInput = "DEBE" THEN
          sInput = "DB.&nbsp;"
      ELSE
          sInput = "&nbsp;&nbsp;&nbsp;HB.&nbsp;"
      END IF

      sInput = sInput & replace(xrs(16)," ","&nbsp;") & "&nbsp;" & replace(xrs(15)," ","&nbsp;")
      sInput = sInput & "<br>" & xrs(6)

         wkdatoobs = replace(sInput &"",chr(13),"<br>")


       end if



   CASE "9"   'TABLAS

    strsql = "SELECT * From Objtabcon Where OBJID = 0" & id
   ' response.write strsql
       set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not xrs.eof then

          sInput = ""

          for ix = 0 to 10

              sConten = xrs(ix + 6)

              sInput = sInput & sConten & "&nbsp;-&nbsp;"

          next


         wkdatoobs = replace(sInput &"",chr(13),"<br>")


       end if

   CASE "10" 'documentos asociados

      Strsql = "SELECT descri, archivo, objid, directorio "
      Strsql = Strsql & " from ObjHtm Where OBJID = 0" & id



   ' response.write strsql
       set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not xrs.eof then
          sInput = Xrs(0) & "&nbsp;-&nbsp;" & Xrs(1) & "&nbsp;-&nbsp;" & Xrs(3)
          wkdatoobs = replace(sInput &"",chr(13),"<br>")

      end if


   CASE "11" 'CONTENIDO

        Strsql = "SELECT * FROM OBJESTdet WHERE OBJID = 0" & id


   ' response.write strsql
       set Xrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not xrs.eof then
          select case xrs("dettipo")
                 case "A"
                   col="dettitulo"
                 case "G"
                   col="dettexto"
                 case else
                   col="dettitulo"
          end select

      wkdatoobs = replace(Xrs(col) &"",chr(13),"<br>")



      end if





   CASE ELSE
      wkdatoobs = cstr(session("codSol")&"")
  END SELECT



   shtml = "<div id='mt2'  style='{height:100px;overflow:auto;font-family:tahoma;font-size:8pt;color:#bbbbbb;border-style:solid;border-color:#993333;border-width:1;' >"
   shtml = shtml & "<br>&nbsp;" & wkdatoobs & "</div><br>"










   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
   xxxx      = GETFECH()
   vHORAHOY  = XLHORDAT


  strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codcircuito = '"& idcircuito &"'  and codinx2 = " & id & " and "
  strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"' and "
  strsql = strsql & " codusr = '" & session("usuario") & "' "
  strsql = strsql & " order by objid"
' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



   if  rs.eof then
    strsql= "insert into OBJOBSCMT (codsol,codobjeto,tipobjeto,codusr,codinx1,codinx2,codinx3,codcircuito,fecha,hora,TIPO) values('" & session("codSol") & "','" & nCodigo & "','" & sTipo & "','" & session("usuario") & "','" & ID2 & "','" & id & "','" & id3 & "','" & idcircuito & "','" & vFECHAHOY & "','" & vHORAHOY & "','OBS')"
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codcircuito='"& idcircuito &"'  and codinx2=" & id & " and codobjeto = " & nCodigo & " and tipobjeto='"& sTipo &"' and codusr='" & session("usuario") & "'"
    strsql=strsql & " order by objid"

    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   end if

  iform=0

  do while  not rs.eof


      stylob  = ""
      sform="modtxt" & iform

      ' timez=clng(timeserial(hour(now),minute(now),second(now)))




      'Texto

    sInput = "<font class=at><textarea class=dt name=" & rs(13).name & "  id=" & rs(13).name & " rows=8 cols=120 "
    if bModificar and wasotxt = "N" then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" &  trim(rs(13)&"")  &"</textarea></font>"




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


    sinput1 ="<input type=hidden name=fecha value=" & vFECHAHOY &">"
    sinput2 ="<input type=hidden name=hora value="&vHORAHOY&">"

    'Eliminar
      sHref = ""
      if bModificar and wasotxt="N" then
	sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=OBJOBSCMT&volver=" & vuelta
	sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
	Celda = ofv.BotonEliminar(rs(0),sHref)
      else
	Celda = ofv.GenCelda("","","","","","","","","si")
     end if


    Celda =  celda &  ofv.GenCelda("",stylob,"","","","","",sInput & sinput3 & sinput1 & sinput2 ,"si")





    if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0)& "&Tabla=OBJOBSCMT&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",celda,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",celda,"si")
      end if



      Fila  = fila & ofv.GenRow("","","","","","","",Formulario,"si")

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







rs.movenext
iform=iform+1
loop


 shtml =shtml & ofv.GenTabla("","","","80%   ","","0","","0","0","","",fila,"si")

  Agregar = ""
  Buscar = ""

   if bModificar then

     IF OP="" THEN
      sAccion = "genObs.asp?" & REQUEST.QUERYSTRING & "&OP=INS"
     ELSE
      sAccion = "genObs.asp?" & REQUEST.QUERYSTRING
     END IF
      Agregar = ofv.BotonAgregar(sAccion)
   end if
  ' response.write sAccion

   sAccion = right(session("vuelve") ,len(session("vuelve") )-3)& "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
   sRetornar = ofv.BotonVolver(sAccion)


  shtml =shtml &  ofv.botoneraESSX("0",rs,cantregmover,cantregmostrar,"no","no","","","","","",Buscar,sRetornar)

  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

response.write shtml &  "</body></html>"
end if



%>