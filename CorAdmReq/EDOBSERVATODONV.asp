<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0


  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  set cn = ofv.conectar(ofv.strconn6)
  set cn3 = ofv.conectar(ofv.strconn2)

  solapa = request.querystring("solapa")
  Nsolapa = request.querystring("Nsolapa")

  SMODAL = request.querystring("MODAL") & ""

  IF SMODAL = "TODOS" THEN

     SWHERE = ""


  ELSE

     SWHERE = " AND USUARIO = '" & SESSION("USUARIO") & "' "

  END IF 

 sopcion = request.querystring("opcion")

 IF SOPCION = "SI" THEN
    StrSql = "INSERT INTO REQINFO (CODIGO,TIPO,DESCRI,USUARIO,FECHA) VALUES ('" & session("NCliente") & "','1','DESCRIPCION','" & SESSION("USUARIO") & "','0') "
   ''' response.write strsql & "<br>" 

    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 END IF 


    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"


  sHTML = ofv.FormHeader(session("FormN"))


  sPath = Request.ServerVariables("PATH_INFO")
  Xruta = Request.ServerVariables("SERVER_NAME") & left(sPath,instrRev(sPath,"/",-1))


  Xfer   =  Xruta
  rutdir = mid(Xfer,1,instrrev(Xfer,"/",-1)) & "legajos/" & session("NCliente")
  nomdir = "Documentos-archivos-imagenes"



  rutdir = "legajos\" & session("NCliente")
  rutdir2 = "legajos/" & session("NCliente") & "/"




  StrSql = "Select * From REQUERIMIENTO where codigo = '" & session("NCliente") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     xCLIENTE = RS(2) 
  end if


    sInput = "REQUERIMIENTO: " & xCLIENTE
    celdas = ofv.GenCelda("","VtI style='background-color:#AAAAAA;color:#ffffff;'","","","","","",sInput,"si")


    filas = ofv.GenRow("","","","","25","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#990000;' ","","98%","","0","","1","1","","",filas,"si") & "<BR>"



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","VtI style='background-color:#AAAAAA;color:#ffffff;' colspan=3","","","","","",sImagen & "&nbsp;OBSERVACIONES","si")
  filas  = ofv.GenRow("","","","","25","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
'  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





  formulario0 = filas

  filas = ""



  StrSql = "Select * From REQINFO where CODIGO = '" & session("NCliente") & "' " & SWHERE & " ORDER BY USUARIO, ID DESC  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 


      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","90","","",celdas,"si")


      celdas = ofv.GenCelda("","ut COLSPAN=2","center","90%","","","","<p><h5>No existen observaciones ingresadas .</h5></p>","si")

      sinput = "&nbsp;"


      celdas = celdas & ofv.GenCelda("","VT style='border-style:solid;border-width:1;border-color:#dddddd;' ","center","10%","","","",sinput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
     filas = filas & ofv.GenRow("","","","","90","","",celdas,"si")


      formulario0 = formulario0 & filas

 

  else



  set cn2 = ofv.conectar(ofv.strconn0)

    DO WHILE NOT RS.EOF


  IF (SMODAL & "") <> "TODOS" THEN

       vNuevoValor = clng(DATE()) - clng(cdate("01-01-1900"))

      sInput = ofv.GenerarInput(rs(5).name,vNuevoValor,"HIDDEN","Usuarios","1","dt","")

      nCol = 80

      if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

      if not  isnull(rs(3)) then 
         wdescri = rs(3)
      end if 




      t3b = "<font class=at><textarea class=dt STYLE='text-align:justify;background-color:transparent;padding-top:2px;FONT-SIZE:7pt;'' "
      t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=5 "
   '   if bModificar then
      '   t3b = t3b & " onblur='document.Usuarios.submit();' "
         t3b = t3b & " onkeypress=" & chr(34)
         t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
         t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
   '   else
   '      t3b = t3b & " readonly >"
    '  end if
      xdatos =  wdescri
      t3b = t3b & xdatos & "</textarea></font>" & sInput

  ELSE

      t3b = rs(3) & ""

  END IF






  celdas = ofv.GenCelda("","UTi  rowspan=2 STYLE='text-align:justify;' ","","60%","","","",t3b,"si")




    if clng(rs(5)) = 0 then
      dFecha = "---"
    else
    ''  dFecha = cdate(rs(5) + clng(cdate("01-01-1900")))
      dFecha = GTFechadorsh(rs(5))
    end if

  sInputx = "------"
       strsql = "Select * From USUARIOS where ID = '" & rs(6) & "' "
       set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
       if not rsniv.eof then sInputx = rsniv(1)
       call ofv.cerrarconsulta(rsniv)

  SINPUT = dFecha & "&nbsp;/&nbsp;" & sInputx

  celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","","10%","","","",dFecha,"si")
  celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","","30%","","","",sInputx,"si")



  filas = ofv.GenRow("","","","","10","","",celdas,"si")



  IF (RS(7) & "") <> "" THEN

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  TARGET='_blank' href=" & chr(34) & rutdir2 & rs(7) & chr(34) &  " "
     sinput = sinput  & "  TITLE='CONSULTAR'  ><b>CONSULTAR</b></a>" 

     sinputk1 = sinput

  ELSE

     sinputk1 = "&nbsp;---&nbsp;"

  END IF

  IF (SMODAL & "") <> "TODOS" THEN

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href='GPuploadn.asp?direct=" &  rutdir & "&nomdir=" & nomdir & "&GPITEM=" & rs(0) & "' "
     sinput = sinput  & "  TITLE='ASOCIAR'  ><b><FONT COLOR='#990000' >ASOCIAR</FONT></b></a>" 

     sinputk2 = sinput

  ELSE

     sinputk2 = "&nbsp;---&nbsp;"

  END IF


  celdas = ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","CENTER","","","","",sinputk1,"si")
  celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:" & session("CMFMENU") & ";' ","CENTER","","","","",sinputk2,"si")

  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  IF (SMODAL & "") <> "TODOS" THEN

     celdas = ofv.GenCelda("","AA COLSPAN=3 bgcolor='#f0f0f0'  ","RIGHT","","","","","<input type='submit' value='ACTUALIZAR' class=bt >","si")
     filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  end if

     celdas = ofv.GenCelda("","AA COLSPAN=3 bgcolor='#f0f0f0' style='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:" & session("CMFMENU") & ";'  ","CENTER","","","","","&nbsp;","si")
     filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



  '''formulario0 = formulario0 &  Filas

  '  if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQINFO"
      formulario0 = formulario0 &  ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
  '  else
   '   formulario0 = formulario0 &  ofv.GenForm("Usuarios","","","","post","",filas,"si")
  '  end if




  RS.MOVENEXT
  LOOP


  end if

    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;' ","","100%","","0","","2","2","","",formulario0,"si")
    sHTML = sHTML & "</center>"

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  'Botonera
  Agregar = ""
  if (SMODAL & "") <> "TODOS" then
     saccion = "edobservaTODOnv.asp?OPCION=SI"
     Agregar = ofv.BotonAgregar(sAccion)
  end if

  sAccion = "EDREQUERNV.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","",Agregar,Buscar,celda,VOLVER)



    sHTML = sHTML & "</body></html>"
    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)

end if

Response.Write sHTML

%>