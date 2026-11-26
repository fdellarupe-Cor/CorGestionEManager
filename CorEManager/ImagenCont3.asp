<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("ID")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From compestr Where codcomp = " & ID & " order by objid "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  strsql = "Select * From compdet Where codcomp = " & ID & " order by codestr,datoformul,objid"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


    Fila = ""
    sestr = "no"
    do while not rs.eof
    sestr = "si"
        'Estructura
           Celda = ofv.GenCelda("","ttt colspan=2","","","","","","Estructura: " & rs(3),"si")
           Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
           Celda = ofv.GenCelda("","ttr","","","","","","Descripcion","si")
           Celda = Celda & ofv.GenCelda("","ttr","","","","","","Tipo","si")
           Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

              sdeta = "no"
              sind = 0
      'Detalle
    do while not rs2.eof
           if cstr(rs(0)) = cstr(rs2(3)) THEN
              sdeta = "si"
              sind = sind + 1
           Celda = ofv.GenCelda("","uti","","70%","","","",cstr(sind) & " - " & rs2(4),"si")
           Celda = Celda & ofv.GenCelda("","uti","","30%","","","",rs2(5),"si")
           Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
              rs2.movenext
           elseIF cstr(rs(0)) > cstr(rs2(3)) THEN
              rs2.movenext
           else
              exit do
           end if
    loop

           if sdeta = "no" THEN
           Celda = ofv.GenCelda("","dt","","70%","","","","Sin Datos","si")
           Celda = Celda & ofv.GenCelda("","dt","","30%","","","","&nbsp;","si")
           Fila = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
           end if

      rs.movenext
    loop

           if sestr = "no" THEN
              sestr = "<center><font color=#336699><h4><p align=center>Sin Estructura disponible"
              sestr = sestr & "</p></h4></font></center>"
           Celda = ofv.GenCelda("","xx colspan=2","","70%","","","",sestr,"si")
           Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
           end if

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  sHTML = sHTML & ofv.GenTabla("","aaa","","95% ","","0","0","0","0","","",Fila,"si")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  'sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)



  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>