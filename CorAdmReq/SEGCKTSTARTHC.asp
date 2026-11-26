<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0

 ''' RESPONSE.WRITE REQUEST.QUERYSTRING
  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


 on error resume next
 CLI = request.querystring("CLI")

 IF CLI <> "" THEN
    SESSION("SEGCLIENTE") = CLI
 END IF

  set cn2 = ofv.conectar(ofv.strconn0)

  StrSql = "Select * From USUARIOS WHERE ID = '" & Session("Usuario") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs.eof THEN




       sInput = "USUARIO:"
       celdas = ofv.GenCelda("","TT","","20%","","","",sInput,"si")

       sInput = Session("SEGUsuario")
       celdas = celdas & ofv.GenCelda("","CC","","80%","","","",sInput,"si")

       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "GRUPO:"
       celdas = ofv.GenCelda("","TT","","20%","","","",sInput,"si")

       sInput = SESSION("SEGGRP")
       celdas = celdas & ofv.GenCelda("","CC","","80%","","","",sInput,"si")

       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "SECTOR:"
       celdas = ofv.GenCelda("","TT","","20%","","","",sInput,"si")

       sInput = "SIN ASIGNAR"

  set cn2 = ofv.conectar(ofv.strconn2)
           ON ERROR RESUME NEXT
           strsql = "Select * From ECO where ECO = " & rs(21)
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)
  call ofv.cerrarconn(cn2)

       celdas = celdas & ofv.GenCelda("","CC","","80%","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       sInput = "CIRCUITO:"
       celdas = ofv.GenCelda("","TT","","20%","","","",sInput,"si")

       sInput = session("WKFDSC")
       celdas = celdas & ofv.GenCelda("","CC","","80%","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"
  END IF 

    sInput = "CLIENTE:"
    celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

    sInput = "SIN ASIGNAR"

  IF SESSION("SEGCLIENTE") <> "NO" THEN

  StrSql = "Select NOMBRE From CLIENTE WHERE CODIGO = '" & SESSION("SEGCLIENTE") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sInput = SESSION("SEGCLIENTE") & " - <B>" & RS(0) & "</B>"
  END IF

  END IF
    celdas = celdas & ofv.GenCelda("","cc","","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")
    sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"

    IF SESSION("SEGCLIENTE") <> "NO" THEN

       sinput = "<a class=tt TARGET=_parent href=SEGCKTSTARTINI.asp"
       sinput = sinput  & " ><b>INICIAR</b></a>" 
       celdas = ofv.GenCelda("","tt","center","","50%","","",sInput,"si")

    ELSE
       sinput = "&nbsp;"
       celdas = ofv.GenCelda("","","center","","50%","","",sInput,"si")

    END IF

       sinput = "<a class=tt TARGET=_parent href=SEGCIRCUITO.asp"
       sinput = sinput  & " ><b>CANCELAR</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","50%","","",sInput,"si")




    filas =  ofv.GenRow("","","","","","","",celdas,"si")

    sHTML = sHTML  & ofv.GenTabla("","","","20%","","0","","0","0","","",filas,"si")



  




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)



  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML





%>