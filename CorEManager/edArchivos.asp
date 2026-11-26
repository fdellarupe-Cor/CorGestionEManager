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
  aplcheck()

  XXXX = GENLOGEXT("EDIDET")

     session("comp31") = "TODOS"
     session("comp32") = "HTM"

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sKIND = request.querystring("SKIND")
  SOLN=request.querystring("SOLN")
  IF SKIND <> "" THEN
     SESSION("SKIND") = SKIND
  ELSE
     SKIND = SESSION("SKIND")
  END IF

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

  sHTML = ofv.FormHeader(session("FormN"))
'RESPONSE.WRITE REQUEST.QUERYSTRING
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------




  vuelta = ofv.vueltaasp2("../")

  'Encabezados
  fila = ofv.Encabezados("Orden","#","Descripcion","Documento","Test")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
 'response.write strsql & " strsql " & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
    nestado = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

    Strsql = "SELECT Objhtm.*, b.codinx2 "

    strsql = strsql & " FROM Objhtm "
    'strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (Objhtm.objid = b.codinx2) "
   ' strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where  codsol = " & session("codSol") & ") b on (Objhtm.objid = b.codinx2) "
    strsql = strsql & " left join (select distinct codinx2 from OBJOBSCMT where TIPO='OBS' AND  codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on (cast (Objhtm.objid as nvarchar) = b.codinx2) "
    Strsql = Strsql & " where  codobjeto=" & nCodigo  & " AND TIPO in (3, " & SKIND & ",1) Order By secuen"
   'response.write strsql & " strsql " & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
  CONT = CONT + 1
    sForm = "edarchivos" & rs(0)

    wsCampo = rs(0)

   'Eliminar
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objhtm&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
 ''   Celda = ofv.BotonEliminar(rs(4),sHref)
      Celda = ofv.BotonEliminar("Eliminacion",sHref)
      Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

xclas = "ut"

      ' estado observacion

      Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")


     'response.write wsCampo & " wsCampo " & "<br>"
     'response.write nCodigo & " nCodigo " & "<br>"
     'response.write sTipo   & " sTipo " & "<br>"

''''' agregado para color del cuadrito de Observaciones version 4.6
       wComentario = ""
       strsql = "SELECT * from OBJOBSCMT where TIPO='OBS' AND codinx2 = " & wsCampo & " and "
       strsql = strsql & " codobjeto = " & nCodigo & " and tipobjeto = '"& sTipo &"'"     ' and "
       strsql = strsql & " order by codcircuito, codusr, objid"
	  'response.write strsql
	   set rss = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rss.eof then
          wComentario = rss(13)
       end if
       call ofv.cerrarconsulta(rss)
''''''

      if (Nestado=5  or Nestado=6) and (rs(15) & "") <> "" and bModificar then

      	       sAccion = "getObsupd.asp?wasotxt=N&vuelve=" & vuelta  & "&Codobjeto=" & nCodigo
      	       sAccion = sAccion & "&Tipobjeto=" & sTipo & "&ID=" & rs(15)  & "&ID2=" & rs(15)

      	       xsffun  = " onclick=""document.getElementById('DX4').style.display='';"" "
      	       sHref   = sAccion

      	 ''''''agregado para color del cuadrito de Observaciones version 4.6
               if wComentario = "" then
      	          sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;color:#993333;' >" & "OBS" & "</div>"
       	       else
       	          sInput  = "<div style='border-style:solid;border-color:#008000;border-width:1;color:#00AA00;' >" & "OBS" & "</div>"
               end if
         ''''''
            Celdaobs  = ofv.GenLink("",xclas & "  ",sHref,"WK4","","",sInput,"si")
      	       Celdaobs  = ofv.GenCelda("",xclas & xsffun & "  title='Ver observaciones para el elemento.' ","","","1","","",Celdaobs,"si")

      end if


      Celda = Celda & Celdaobs

       if bModificar then
       sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objhtm&volver=" & vuelta
       sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
       end if

    'Secuencia
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm & "r1","5","dtn","SI")
      sInput = ofv.GenForm(sForm & "r1","",sAccion,"","post","",sInput,"si")
    else
      'sInput = ofv.GenerarInput("",rs(3),"readonly","","5","dtn","SI")
        sInput = rs(3)
    end if
    Celda = Celda & ofv.GenCelda("","vt","","15%","","","",sInput,"si")

    'Descripcion
    if bModificar then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm & "r2","200","dt maxlength=70","")

      sInput = REPLACE(sInput,"{ document." & sform  & "r2" & ".submit(); return true;}","{ if (this.value.length > 70) {alert('El texto es demasiado extenso'); return false;} else { document." & sform  & "r2" & ".submit(); return true;}}")
      sInput = ofv.GenForm(sForm & "r2","",sAccion,"","post","",sInput,"si")
    else
      sInput = rs(4)
      'sInput = ofv.GenerarInput("",rs(4),"readonly","","60","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","vt","","50%","","","",sInput,"si")

       sdirec = rs(8) & ""

  ' response.write rs(8)
    if  session("APLMDCLI") ="BC" AND session("codSol")=13 THEN
     sdirec = "ARCHIVOSLEGALES"
    ELSE
    IF SESSION("BASE")="K"  THEN
      sdirec="ARCCOMPONENTES"
    else
      sdirec = RS("DIRECTORIO") '"ARCHIVOS"
    end if
   ' if sdirec = "" then sdirec = "ARCHIVOS"
    END IF



  'Archivo
     if bModificar then
       sImagen = ofv.GenImage("","aaa valign=bottom ","","","Iconos/ICorEsq1.jpg","0","")
       sHref = "EdArchivosProp.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Docum=" & rs(0) & "&masdir=" & sdirec
       sLink = ofv.GenLink("","",sHref,"","","",sImagen,"si")
       sinput = sLink

      ' sinput = "<a class=nt href='" & sHref
      ' sinput = sinput  & "'  target='_blank' >" & sImagen & "</a>"


    else
       sinput = "&nbsp;"
    end if
    Celda = Celda & ofv.gencelda("","vt","center","15%","","","",sInput,"si")





''    if bModificar then
''      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objhtm&volver=" & vuelta
''      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
''      Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")
''    else
''      Formulario = ofv.GenForm("","","","","post","",Celda,"si")
''    end if




    if rs(6) <> "" then

	       sImagenP = "<b><font face=CorFlaticon size=3 color=#F7AA22>&#0232;</font></b>"

      if  rs("tipo") <> 3 then
      if session("EMPMUN") = "S" then
        Sadddir="Emp" & SESSION("D" & nCodigo & "-0" ) &"/"
      else
         Sadddir=""
      end if
       'sImagenP = "<b><font face=WINGdings size=3 color=#F7AA22>&#0232;</font></b>"

       sinput = "<a class=nt href=CoArchivos3.asp?Archivo=" & Sadddir & sdirec & "/" & replace(rs(6)," ","%20")
       sinput = sinput  & " title='Propiedades de " & rs(4) & "' target='_blank' >" & sImagenP & "</a>"
      else
      ' sImagenP = "<b><font face=WINGdings size=3 color=#F7AA22>&#0232;</font></b>"
       sImagenP = "<b><font face=CorFlaticon size=3 color=#F7AA22>&#0232;</font></b>"
	    sHref =  rs("archivo")
	    sLink = ofv.GenLink("","",sHref,"_blank","","",sImagenp,"si")
       sinput = sLink


     end if

     '  sinput = sinput  & "' title='Propiedades de " & rs(4) & "' target='_blank' >" & sImagenP & "</a>"

    else
       sinput = "&nbsp;"
    end if

     Celda = Celda & ofv.gencelda("","vt","center","","15%","","",sInput,"si")

   'Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")
    Fila = fila & ofv.GenRow("","","","","","","",celda,"si")



    rs.movenext
  loop

'  sHTML = sHTML & "<br><br><br><br>"
 ' sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")
  if CONT = 0 then
  			sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes documentos agregados.</div>"
  			       	else
  			       	    sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")
  			       	    sHTML = sHTML & tabla
			end if

    Agregar = ""

  'Botonera
  if bModificar then
    sAccion = "AgregarArchivos.asp?Tabla=Objhtm&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

         'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)


      end if


   end if





'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
 ' sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
    sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCOBS()

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>