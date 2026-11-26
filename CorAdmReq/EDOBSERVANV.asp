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
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






 sopcion = request.querystring("opcion")

 IF SOPCION = "SI" THEN
    StrSql = "INSERT INTO REQINFO (CODIGO,TIPO,DESCRI,USUARIO,FECHA) VALUES ('" & session("NCliente") & "','1','DESCRIPCION','" & SESSION("USUARIO") & "','0') "
   ''' response.write strsql & "<br>" 

    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 END IF 




  StrSql = "Select * From REQINFO where CODIGO = '" & session("NCliente") & "' AND USUARIO = '" & SESSION("USUARIO") & "' ORDER BY ID DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 


      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","90","","",celdas,"si")


      celdas = ofv.GenCelda("","ut COLSPAN=2","center","90%","","","","<p><h5>No existen observaciones ingresadas por " & SESSION("segUSUARIO") & ".</h5></p>","si")

      sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
      sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
      sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
      sinput = sinput  & "none" & chr(34) & ";' "
      sinput = sinput  & "  href=EDOBSERVANV.asp?OPCION=SI  "
      sinput = sinput  & "  TITLE='CREAR'  >&nbsp;&nbsp;<b>" & sImagenI3 & "&nbsp;CREAR</b>&nbsp;&nbsp;</a>" 


      celdas = celdas & ofv.GenCelda("","VT style='border-style:solid;border-width:1;border-color:#dddddd;' ","center","10%","","","",sinput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
     filas = filas & ofv.GenRow("","","","","90","","",celdas,"si")


      formulario0 = filas

 

  else


      nCol = 100

  if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

  if not  isnull(rs(3)) then 
     wdescri = rs(3)
  end if 




    t3b = "<font class=at><textarea class=dt STYLE='text-align:justify;background-color:transparent;padding-top:2px;'' "
    t3b = t3b & "name=" & rs(3).name & " cols=" & nCol & " rows=20 "
    if bModificar then
      t3b = t3b & " onblur=Usuarios.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = ofv.GenCelda("","abc colspan=3 ","center","100%","","","",t3b,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","aaa colspan=3 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & session("CMFMENU") & ";' ","","100%","","","","&nbsp;","si")
    FilaS =  FilaS & ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then

       if nFecini < clng(cdate("01-01-1900")) then
          nFecini = clng(cdate("01-01-1900"))
       else
          nFecini = 0
       end if
       vNuevoValor = clng(DATE()) - nFecini
       sInput = ofv.GenerarInput(rs(5).name,vNuevoValor,"HIDDEN","Usuarios","1","dt","")

    end if

    celdas = ofv.GenCelda("","ut","","","","","","FECHA&nbsp;ACTUALIZACION:&nbsp;" & sInput ,"si")
    celdas = celdas & ofv.GenCelda("","","","","","","","","si")

    if clng(rs(5)) = 0 then
      dFecha = "---"
    else
      dFecha = cdate(rs(5) + clng(cdate("01-01-1900")))
    end if



      sInput = dFecha




    celdas = celdas & ofv.GenCelda("","vt","","80%","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=REQINFO"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if


  end if

    sHTML = sHTML & "<center>"
    sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:2;border-color:#990000;' ","","100%","","0","","0","0","","",formulario0,"si")
    sHTML = sHTML & "</center>"

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sAccion = "edobservatodonv.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","",Agregar,Buscar,celda,VOLVER)



    sHTML = sHTML & "</body></html>"
    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)

end if

Response.Write sHTML

%>