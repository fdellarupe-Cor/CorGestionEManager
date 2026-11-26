<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")



  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  sHTML  = ofv.FormHeader(session("FormN"))
  vuelta = ofv.vueltaasp2("../")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div name=mt2 id=mt2 style=""sdisplay:none;height:350px;overflow:auto;" & styl & """ >"

'-------------------------------------------------

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

%>
<!-- #INCLUDE FILE="IncFile/CorVinculadorPL.inc" -->
<%


  IF WKSIGUE = "S" THEN

  'Encabezados

  if wkaplic = "PL" then

     Formulario = ofv.Encabezados("Descripcion","Secuencia")

  else

     Formulario = ofv.GenTabla("","","","90%","","0","","0","0","","",ofv.Encabezados("Descripcion","Secuencia","Proceso"),"si")

  end if

  'Formulario= ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario,"si")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False


  Strsql = "Select * From Objeven Where codobjeto=" & nCodigo & " Order By secuen"
 ' response.write Strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    sForm = "xxedEventos" & rs(0)
  sAccion=""
      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objeven&volver=" & vuelta
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

      '  Fila =  ofv.GenRow("","","","","","","",fila,"si")  ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
     ' else

      '  Fila = ofv.GenForm("","","","","post","",Fila,"si")
    end if
    'Eliminar

    Formulario= formulario & ofv.GenForm(sForm,"",sAccion,"","post","","","no")
    Formulario= formulario & ofv.GenTabla("","","","90%","","0","","0","0","","","","no")
    if bModificar and (rs(13) & "") <> "PL" then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objeven&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
 ''   Celda = ofv.BotonEliminar(rs(2),sHref)
      Celda = ofv.BotonEliminar("Eliminacion",sHref)
      'Celda = replace(Celda,"<TD ","<TD title='Eliminacion' ")
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Descripcion
    if bModificar and (rs(13) & "") <> "PL" then
      sInput =ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"50","dt","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","50","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","vt","","20%","","","",ofv.GenForm(sForm ,"",sAccion,"","post","","","no")& sInput,"si")

    'Secuencia
    if bModificar and (rs(13) & "") <> "PL" then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"5","dtn","si")
    else
      sInput = ofv.GenerarInput("",rs(4),"readonly","","5","dtn","si")
    end if
    Celda = Celda & ofv.GenCelda("","vt","","5%","","","",sInput,"si")

   if wkaplic = "PL" then

   else

    'Proceso
    if  bModificar then
        Strsql = "Select * From CSPROCESOS Order By DSCSECUEN"
         IF LEN(rs(11).value &"" )=0 THEN
          VALOR=""
        ELSE
         VALOR=rs(11).value
        END IF
        sInput =ofv.GenerarCombo(sForm,rs(11).name,VALOR,Strsql,cn,0,10,"")
    else
        if  cstr(rs(11) & "") <> "" then
            Strsql = "Select DSCPROC From CSPROCESOS Where CODPROCESO = '" & rs(11) & "' "
            set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
            call ofv.cerrarconsulta(rsInstrumento)
        end if
        sInput = ofv.GenerarInput("",sInstrumento,"readonly","","90","dt","")
    end if
   Celda = Celda & ofv.GenCelda("","vt","","50%","","","",sInput  ,"si")& ofv.GenCelda("","vt","","50%","","","",""  ,"si")

   end if


   ' Fila = Celda
    sAccion=""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objeven&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    '  Fila =  ofv.GenRow("","","","","","","",fila,"si")  ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
   ' else

    '  Fila = ofv.GenForm("","","","","post","",Fila,"si")
    end if
     Fila =ofv.GenRow("","","","","","","",celda ,"si") & "</table></form>"
     Formulario = Formulario & Fila '& ofv.GenForm(sForm ,"",sAccion,"","post","",Fila,"si")

   ' Formulario = Formulario & ofv.GenRow("","","","","","","",fila,"si")

rs.movenext
  loop

' & "<br><br><br><br>"
'sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario,"si")

  END IF
'sHTML = sHTML & Formulario

  'Botonera
  if bModificar then
    sAccion = "AgregarEvento.asp?Tabla=Objeven&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
    Agregar = ofv.BotonAgregar(sAccion)

  end if

  'cambios
   Listar = ""


   if bVersion = "SI" then
      if SESSION("SHWMOD") <> "N" then

      elseif SESSION("VIEWMOD") <> "N" then

        ' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
         Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

         sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
         sHTML = sHTML & ConsCambios(sasp)

         Celda = Listar


      end if


   end if




'--------------------------------------------------------------------------------

      sHTML = sHTML & formulario & "</div>"
 Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1 id=rrr","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

 sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

    sHTML = sHTML & ofv.botoneraESSX("2",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

 ' sHTML = sHTML & " <script>document.getElementById('mt2').innerHTML=document.getElementById('rrr').innerHTML;alert(document.getElementById('mt2').innerHTML);</script></body></html>"
end if

response.write sHTML  & "</body>"
'Whtml =   "<script> document.getElementById('mt1').style.display='none';" &chr(13)
'Whtml =   "<script>alert(document.getElementById('mt1'));"
''' WhtmlWhtml =   "<script>" = Whtml & " alert(document.getElementById('TXHI3M').innerHTML);"

' Whtml = Whtml & " document.getElementById('tmtshwxz').innerHTML='<div NAME=divtmtshwxz id=divtmtshwxz style={height:600px;overflow:auto;}>' + document.getElementById('mt2').innerHTML  + '</div>';" &chr(13)
'Whtml = Whtml & "document.getElementById('tmtshwxz').innerHTML=document.getElementById('mt2').innerHTML;"  &chr(13)
' Whtml = Whtml & "document.getElementById('tmtshwxy').innerHTML= '<div id=xdestino style={width:200px;height:600px;overflow:auto;}>' + document.getElementById('TXHI3M').innerHTML + '</div>';alert('1');</script>" &chr(13)
'Whtml = Whtml &   "alert('1');</script> " &chr(13)
'response.write  Whtml & "</html>"


%>