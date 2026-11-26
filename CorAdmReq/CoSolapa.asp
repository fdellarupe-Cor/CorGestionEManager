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

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

  set cn = ofv.conectar(ofv.strconn6)
  set cn3 = ofv.conectar(ofv.strconn2)

  solapa = request.querystring("solapa")
  Nsolapa = request.querystring("Nsolapa")

  IF SOLAPA <> "" THEN 
     SESSION("SOLAPA") = SOLAPA
     SESSION("NSOLAPA") = NSOLAPA
  ELSE
     SOLAPA = SESSION("SOLAPA")
     NSOLAPA = SESSION("NSOLAPA")
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





    sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;" & NSOLAPA & " ","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






 sopcion = request.querystring("opcion")

 IF SOPCION = "SI" THEN
    StrSql = "INSERT INTO REQINFO (CODIGO,TIPO,DESCRI,SECTOR,FECHA) VALUES ('" & session("NCliente") & "','" & SOLAPA & "','DESCRIPCION','0','0') "
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 END IF 




  StrSql = "Select * From REQINFO where CODIGO = '" & session("NCliente") & "' AND TIPO = " & solapa
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 
    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)

      celdas = ofv.GenCelda("","ut COLSPAN=2","","90%","","","","<p>NO HAY INFORMACION " & NSOLAPA & ".</p>","si")

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=COSOLAPA.asp?OPCION=SI  "
     sinput = sinput  & "  TITLE='CREAR'  >&nbsp;&nbsp;<b>" & sImagenI3 & "&nbsp;CREAR</b>&nbsp;&nbsp;</a>" 


      celdas = celdas & ofv.GenCelda("","VT style='border-style:solid;border-width:1;border-color:#dddddd;' ","center","10%","","","",sinput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    sHTML = sHTML & ofv.GenTabla(""," ","","800","","0","","-1","-1","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"
    sHTML = sHTML & "</body></html>"

  else



    'Encabezado DESCRIPCION
    Celda = ofv.GenCelda("","aaa colspan=3","","100%","","","","&nbsp;","si")
    FilaS =  FilaS & ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","ut colspan=3  ","","100%","","","","DESCRIPCION " & NSOLAPA,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

      nCol = 120

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







    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla(""," ","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"

    sHTML = sHTML & "</body></html>"
    call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)
  end if
end if

Response.Write sHTML

%>