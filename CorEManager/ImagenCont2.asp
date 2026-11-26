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

  strsql = "Select * From Compimag Where codcomp = " & ID
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  i = 0
  redim preserve aImagen(i)
  redim preserve aDescri(i)
  redim preserve aSecuen(i)
  do until rs.eof
    redim preserve aImagen(i)
    redim preserve aDescri(i)
    redim preserve aSecuen(i)
    aImagen(i) = rs(3)
    aDescri(i) = rs(4)
    aSecuen(i) = rs(5)
    rs.movenext
    i = i + 1
  loop


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------






  sHTML = sHTML & "<center><font color=" & Session("CMFPNITFT") & ">"

  for i = lbound(aImagen) to ubound(aImagen)
    sHTML = sHTML & "<h4><p align=center>"
    if aImagen(i) = "" then
      sInput = "No poseee imagen asociada"
      sImagen = ""
    else
      sInput = "Descripcion: " & aDescri(i) & " Cuerpo: " & aSecuen(i)
if session("EMPMUN") = "S" then
      sImagen = ofv.GenImage("","","","","Emp" & SESSION("C" & ID & "-0" ) & "/Componentes/" & aImagen(i),"1","")
      sImagen = "<a class=nt href='" & "Emp" & SESSION("C" & ID & "-0" ) & "/Componentes/" & aImagen(i) & "' target='_blank' title='Abrir imagen' >" & SImagen & "</a>"
else
	  sImagen = ofv.GenImage("","","","","Componentes/" & aImagen(i),"1","")
      sImagen = "<a class=nt href='" & "Componentes/" & aImagen(i) & "' target='_blank' title='Abrir imagen' >" & SImagen & "</a>"
end if
    end if
    sHTML = sHTML & sInput & "</p></h4>" & sImagen
  next
  sHTML = sHTML & "</font></center>"


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  sHTML = sHTML &  "</body></html>"
   call ofv.cerrarconsulta(rs)
    call ofv.cerrarconn(cn)


end if

Response.Write sHTML

%>