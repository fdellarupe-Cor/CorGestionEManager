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
''  if session("Nomodificar") = "S" then bModificar = False

'''response.write bModificar

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  IDTXT = request.querystring("IDTXT")

  IDF = request.querystring("F")
  IDC = request.querystring("C")

  vuelve = request.querystring("vuelve")
  vuelve = replace(vuelve,"***","_")
  volver = "_Vuelve=" & replace(vuelve,"_","***")
  addqry = replace(vuelve,"_","&")

  sform="modtxt"

  wposi = 0
  wpossal = 0

  vuelta = "../edportadadet.asp?F=" & IDF & "_C=" & IDC

''response.write bModificar

  eststt = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;border-style:outset;border-width:2;} "
  estst1 = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#99ccff;} "
  estst2 = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#88bbee;} "
  estst3 = " STYLE={font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#77aadd;} "
  estst4 = " STYLE={font-family:tahoma,arial;color:#eeeeee;font-size:8pt;background-color:#6699cc;} "
  estst5 = " STYLE={font-family:tahoma,arial;color:#eeeeee;font-size:12pt;background-color:#5588bb;} "
  esttxt = " STYLE={font-family:tahoma,arial;color:#000000;font-size:10pt;background-color:#ffffff;border-style:outset;border-width:2;} "

'f7eedd

  redim nivel(12)
  nivel(0)  = "120"
  nivel(1)  = "110"
  nivel(2)  = "100"
  nivel(3)  = "90"
  nivel(4)  = "80"
  nivel(5)  = "70"
  nivel(6)  = "60"
  nivel(7)  = "50"
  nivel(8)  = "40"
  nivel(9)  = "30"
  nivel(10) = "20"
  nivel(11) = "10"
  nivel(12) = "5"

  redim nivel2(2)
  nivel2(0)  = "IZQUIERDA"
  nivel2(1)  = "CENTRADO"
  nivel2(2)  = "DERECHA"




  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

response.write sHTML

  sHTML = ""

  set cn = ofv.conectar(ofv.strconn2)


  strsql = "SELECT * from Objportada where  OBJportada.codobjeto=" & nCodigo & " and  OBJportada.detcon = " & IDF & " and OBJportada.detsecuen = " & IDC
 'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if  rs.eof then

             strsql = "insert into Objportada  ("
             strsql = strsql & "codobjeto,tipobjeto,"
             strsql = strsql & "detcon,"
             strsql = strsql & "detsecuen,"
             strsql = strsql & "dettipo,"
             strsql = strsql & "detdepen,"
             strsql = strsql & "dettitulo,"
             strsql = strsql & "dettexto,"
             strsql = strsql & "detimagen,"
             strsql = strsql & "dettabla,"
             strsql = strsql & "detfont,"
             strsql = strsql & "detfsize,"
             strsql = strsql & "detfcolor,"
             strsql = strsql & "detbgcolor,"
             strsql = strsql & "detborde,"
             strsql = strsql & "detftipo,"
             strsql = strsql & "detsang,"
             strsql = strsql & "detindent1,"
             strsql = strsql & "detindent2,"
             strsql = strsql & "detvin,"
             strsql = strsql & "detvintipo,"
             strsql = strsql & "detvinindent,"
             strsql = strsql & "detnumer,"
             strsql = strsql & "detnumertipo,"
             strsql = strsql & "detnumerclase,"
             strsql = strsql & "detvinfont,"
             strsql = strsql & "detalign,"
             strsql = strsql & "dettextalign,"
             strsql = strsql & "modi,"
             strsql = strsql & "tipmodi,"
             strsql = strsql & "detdefault,"
             strsql = strsql & "rellink,"
             strsql = strsql & "reflink,"
             strsql = strsql & "relant,"
             strsql = strsql & "modautor,"
             strsql = strsql & "modfecha,"
             strsql = strsql & "modhora"

             strsql = strsql & ") values ("
             strsql = strsql & nCodigo & ",'" & sTIPo & "',"
             strsql = strsql & "'" & IDF & "',"
             strsql = strsql & "'" & IDC & "',"
             strsql = strsql & "'N',"
             strsql = strsql & "'0',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'SI',"
             strsql = strsql & "'"& nivel2(1) &"',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'1',"
             strsql = strsql & "'',"
             strsql = strsql & "'0',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'',"
             strsql = strsql & "'" & SESSION("USUARIO") & "',"
             strsql = strsql & "'" & WKFECHA & "',"
             strsql = strsql & "'" & WKHORA & "'"

             strsql = strsql & ")"
          '   response.write strsql
             'call ofv.crearconsultaEx(StrSql,cn,1,volver)
             xxx=SqlExLog("INSERT",strsql,CN,"Objportada","","","")


  strsql = "SELECT * from Objportada where OBJportada.codobjeto=" & nCodigo & " and OBJportada.detcon = " & IDF & " and OBJportada.detsecuen = " & IDC
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  end if


  response.write "<center>"

  response.write "<TABLE class='aaa' width=95%  marginwidth=0  cellspacing=0  cellpadding=5 >"



    'null
     sicK = "<i class='las la-file-alt' style='font-size:12pt;' ></i>&nbsp;&nbsp;"
     Celda =  ofv.GenCelda("","ttT valign=top  ","","","","","",sicK & "<b>Modificacion de Texto de portada</b>","si")
     Fila = ofv.GenRow("","","","","","","",celda,"si")
  ''   response.write Fila


  if  not rs.eof then


      scomK = "S" & chr(9) & "N"
      scomD = "SI" & chr(9) & "NO"

         if  isnull(rs(5) & "")   then
             sreducd = "N"
         else
             sreducd = rs(5)
         end if

''response.write " "

         sInputr2 = ofv.gencomboFX(sForm & "x",rs(5).name,sreducD & "",scomK,scomD,"")

      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objportada&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm & "x","",sAccion,"","post","",sInputr2,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",sInputr2,"si")
      end if


      sreduc = "<table width=100% ><tr HEIGHT=30 BGCOLOR='#FFEEDD' >"
      sreduc = sreduc & "<td class='uti' width=10%  >ACTIVO:&nbsp;</td>"
      sreduc = sreduc & "<td class='uti' width=90%  >" & Formulario & "</td>"
      sreduc = sreduc & "</tr></table>"

     Celda =  ofv.GenCelda("","ut","","95%","","","",sreduc,"si")

     Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
   ''  response.write Fila


  if (rs(5) & "") <> "N" then



      wasotxt = "N"
      stylob  = ""




      'Texto

    sInput = "<font class=at><textarea class=dt name=" & rs(8).name & "  id=" & rs(8).name & " rows=10 cols=120 "
    if bModificar and wasotxt = "N" then
       sInput = sInput & " onblur=" & chr(34) & " if (this.value.length > 2300) {alert('El texto es demasiado"
       sInput = sInput & " extenso'); return false;} else {"
       sInput = sInput & " document." & sform & ".submit() ;}" & chr(34)
    else
       sInput = sInput & " readonly "
    end if
    sInput = sInput & " >" & trim(rs(8)&"") & "</textarea></font>"


      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objportada&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",sInput,"si")
      else
        Formulario = ofv.GenForm("","","","","post","",sInput,"si")
      end if

      sInput =  ofv.GenCelda("",stylob,"","","","","",Formulario,"si")








      Fila  = fila & ofv.GenRow("","","","","","","",sInput,"si")









      'IMAGENES


  sform="modimg"

      sicK = "&nbsp;<u><font face='corflaticon' size=5 color=#993333 >N</font></u>"
      if rs(9) <> "" then
         estia=" style={font-size:7pt;font-family:tahoma,arial;color:#eeeeee;background-color:#000080;font-weight:600;border-style:outset;border-width:2;}"

         sImagenP = "&#0232;"
         sinput = "<a class='ft1o'  href='Coedestimg.asp?A=" & rs(16) & "&B=" & rs(15)
         sinput = sinput & "&R=" & rs(12) & "&S=" & server.urlencode(rs(9))
         sinput = sinput & "' target=_blank  title='Ver Imagen' >" & sImagenP & "</a>"
         slink  = sinput



      else
         slink = ""
      end if


     SCOMSEP = ""
     scomK   = ""

     For i=0 to 12
         scomK = scomK & SCOMSEP & nivel(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     if not  isnull(rs(12) & "") then
        sreducD = rs(12)
     else
        sreducD = 100
     end if



      if bModificar then

         sInputr1 = ofv.gencomboFX(sForm & "1",rs(12).name,sreducD & "",scomK,scomD,"")

      else
         if  isnull(rs(12) & "")   then
             sreduc = "100"
         else
             sreduc = rs(12)
         end if

         sInputr1 = ofv.gencomboFX(sForm & "1",rs(12).name,sreducD & "",scomK,scomD,"no")

      end if

      scomK = "SI" & chr(9) & "NO"
      scomD = "SI" & chr(9) & "NO"

         if  isnull(rs(15) & "")   then
             sreducd = "NO"
         else
             sreducd = rs(15)
         end if

      if bModificar then

         sInputr2 = ofv.gencomboFX(sForm & "2",rs(15).name,sreducD & "",scomK,scomD,"")

      else
         if isnull(rs(15))  then
            sbord = "NO"
         else
             sbord = rs(15)
         end if

         sInputr2 = ofv.gencomboFX(sForm & "2",rs(15).name,sreducD & "",scomK,scomD,"no")

      end if


     SCOMSEP = ""
     scomK   = ""

     For i=0 to 2
         scomK = scomK & SCOMSEP & nivel2(i)
         SCOMSEP = chr(9)
     next

     scomD = scomK

     if not  isnull(rs(16) & "") then
        sreducD = rs(16)
     else
        sreducD = "IZQUIERDA"
     end if




      if bModificar then

         sInputr3 = ofv.gencomboFX(sForm & "3",rs(16).name,sreducD & "",scomK,scomD,"")

      else
         if isnull(rs(16) & "")  then
            salng = "IZQUIERDA"
         else
             salng =  rs(16)
         end if

         sInputr3 = ofv.gencomboFX(sForm & "3",rs(16).name,sreducD & "",scomK,scomD,"no")

      end if



      wstexto = rs(9)

      if wstexto <> "" then
         wstexto = chr(34) & wstexto & chr(34)
      else
         wstexto = "<font color='#000000'>--Sin&nbsp;Asociar--</font>"
      end if


      wstexto = "<u>" &  wstexto   & "</u>"

      addqry = "F=" & IDF & "_C=" & IDC
      addqry = addqry & "&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo

      sinput = "<a class=nt href=" & chr(34) & "edPORTADAIMG.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sinput = sinput & "&IDTXT=" & rs(0) & "&vuelve="
      sinput = sinput & replace(addqry,"&","_") & chr(34) & " title='Asignar/Modificar Imagen' >" & wstexto & "</a>"



      sicK = "<font face='corflaticon' size=3 color=#993333 >4</font>&nbsp;"



      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJPORTADA&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sreduc1 = ofv.GenForm(sForm & "1","",sAccion,"","post","",sInputr1,"si")
      else
        sreduc1 = ofv.GenForm("","","","","post","",sInputr1,"si")
      end if



      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJPORTADA&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sreduc2 = ofv.GenForm(sForm & "2","",sAccion,"","post","",sInputr2,"si")
      else
        sreduc2 = ofv.GenForm("","","","","post","",sInputr2,"si")
      end if


      if bModificar then

        vlk = vuelta
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJPORTADA&volver=" & vlk
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        sreduc3 = ofv.GenForm(sForm & "3","",sAccion,"","post","",sInputr3,"si")
      else
        sreduc3 = ofv.GenForm("","","","","post","",sInputr3,"si")
      end if


'''''''''''''''''''
      wstextoLNK = rs(32) & ""

       if  isnumeric(wstextoLNK) and wstextoLNK <> "0" then
           sImagenP = "<font face='corflaticon' size=2 color=#005500 >&#0252;</font>"
       else
           sImagenP = "<font face='corflaticon' size=2 color=#f7aa22 >&#0252;</font>"
       end if


         sID = "EdGENLnk.asp?"
         sID = sID & "ID=" & rs(0) & "&TABLA=OBJPORTADA&Codobjeto=" & nCodigo
         sID = sID & "&Tipobjeto=" & sTipo & "&volver=" & vlk & "_Codobjeto=" & nCodigo
         sID = sID & "_Tipobjeto=" & sTipo

'''  & "&nblk=AL"

      sHref = sID


       sinputk = "<a class='ft1o'  href='" & sHref
       sinputk = sinputk & "' title='Actualizar link a otro documento' >" & sImagenP & "</a>"
       slinkX   = "&nbsp;" & sinputk

    '  Celda = ofv.GenCelda("","ft","left","5%","","","",sLink,"si")
    '  fila  = fila & ofv.GenRow("","aaa   ","","","","","",celda,"si")
''''''''''''''''''


      sreduc = "<table width=95%  STYLE='border-style:solid;border-width:1;border-color:#000000;' >"
      sreduc = sreduc & "<tr BGCOLOR='#DDCCAA' >"
      sreduc = sreduc & "<td class=Ut width=10%  >LINK</td>"
      sreduc = sreduc & "<td class=Ut width=50% COLSPAN=2 >IMAGEN</td>"
      sreduc = sreduc & "<td class=Ut width=10%  >REDUCE</td>"
      sreduc = sreduc & "<td class=Ut width=10%  >RECUADRO</td>"
      sreduc = sreduc & "<td class=Ut width=10%  >ALINEACION</td>"
      sreduc = sreduc & "</tr>"

      sreduc = sreduc & "<tr HEIGHT=50 BGCOLOR='#FFEEDD' >"
      sreduc = sreduc & "<td class=ft width=10% ALIGN=CENTER >" & slinkX & "</td>"
      sreduc = sreduc & "<td width=50% >" & sInput & "&nbsp;</td>"
      sreduc = sreduc & "<td class=ft width=10% ALIGN=CENTER >" & slink & "</td>"
      sreduc = sreduc & "<td width=10% ALIGN=CENTER >" & sreduc1 & "</td>"
      sreduc = sreduc & "<td width=10% ALIGN=CENTER >" & sreduc2 & "</td>"
      sreduc = sreduc & "<td width=10% ALIGN=CENTER >" & sreduc3 & "</td>"
      sreduc = sreduc & "</tr>"
      sreduc = sreduc & "</table>"


      Celda = ofv.GenCelda("","ut","","95%","","","",sreduc,"si")

      fila  = fila & ofv.GenRow("","aaa   ","","","","","",celda,"si")


   ''         response.write Fila


   end if

end if

  response.write Fila
  response.write ofv.GenFTabla()
  response.write "</center>"

  Agregar = ""
  Buscar = ""

  'Boton Volver
  sAccion   = "edPortada.asp?"
  sAccion = sAccion & "Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sRetornar = ofv.BotonVolver(sAccion)





  sHTML = ""
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmover,cantregmostrar,"no","no","","","","",Agregar,Buscar,sRetornar)

response.write sHTML

  Call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

response.write  "</body></html>"


end if



%>