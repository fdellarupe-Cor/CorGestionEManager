<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
 
ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



 sopcion = request.querystring("opcion")
  XTIPO = REQUEST.QUERYSTRING("TIPO")

  IF XTIPO <> "" THEN
     SESSION("TIPO") = XTIPO
  ELSE
     XTIPO = SESSION("TIPO")
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



  StrSql = "Select * From TIPOS WHERE ID = 0" & XTIPO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then





    sHTML = sHTML & "<center>"


      CeldaS = ofv.GenCelda("","ttt colspan=3","center","100%","","","","INFORMACION A COMPLETAR PARA " & RS(1),"si")
      filas = ofv.GenRow("","","","","25","","",celdas,"si")



    'Encabezado DESCRIPCION

      nCol = 80

  if len(wdescri) = 0 then wdescri = "SIN INFORMACION" 

  if not  isnull(rs(2)) then 
     wdescri = rs(2)
  end if 




    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
    t3b = t3b & "name=" & rs(2).name & " cols=" & nCol & " rows=15 "
    if bModificar then
      t3b = t3b & " onblur=Usuarios.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = ofv.GenCelda("","abc colspan=3 VALIGN=MIDDLE ","center","100%","","","",t3b,"si")
    FilaS =  filaS & ofv.GenRow("","","","","250","","",Celda,"si")



    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=TIPOS"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if



  cnbot = "<input type=submit class=bt value='VOLVER' >"
  Cons = ofv.gencelda("UNO","","center","","1%","","",cnbot,"si")
  sAccion = "CORREQTIPOPROP.asp"
  Cons = ofv.genform("","",sAccion,"","post","",Cons,"si")



    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")



    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"

   sAccion = "CORREQTIPOPROP.asp"
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)


  end if



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML




%>