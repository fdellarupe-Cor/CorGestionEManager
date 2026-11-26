<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
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
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

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

 sTIPO = request.querystring("TIPO")
 sPRY = request.querystring("PRY")





    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=2>&#0246;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
 
    sImagenEC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0192;</FONT></B>"
    sImagenNI = "<B><FONT FACE=WINGDINGS COLOR=#336699 SIZE=4>6</FONT></B>"
    sImagents = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>&#0216;</FONT></B>"



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"


    IF STIPO = "OP" THEN

      celdas = ofv.GenCelda("","TTT colspan=2","","","","","","LISTA DE PROYECTOS","si")
      filas  = ofv.GenRow("","","","","20","","",celdas,"si")



      celdas = ofv.GenCelda("","TTR ","","","","","","DESCRIPCION","si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  StrSql = "Select * From PROYECTO where estado <> 13 order by titulo  "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  if rs1.eof then

      SINPUT = "NO HAY PROYECTOS PARA SELECCIONAR"
      celdas = ofv.GenCelda("","vt colspan=2 ","CENTER","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
     do while not rs1.eof

     simagene = sImagenEC
     if rs1(11) = 1 then
        simagene = sImagenNI
     end if

     sinput = "<a CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtasklstbus.asp?TIPO=TP&PRY=" & RS1(1)
     sinput = sinput  & " >" & simagene & "<b>&nbsp;&nbsp;" &  RS1(2) & "</b></a>" 
    
      celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






    rs1.movenext
    loop
 
  end if



  ELSE

  StrSql = "Select * From PROYECTO WHERE CODIGO = '" & sPRY & "' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS1.EOF THEN
     XVPLAN  = RS1(17)
     XPRYNOM = RS1(2)
     if session("NCliente") = sPRY then XVPLAN  = session("VPLAN")
  END IF
  

      celdas = ofv.GenCelda("","TTT colspan=2","","","","","","LISTA DE TAREAS DEL PROYECTO " & XPRYNOM & "","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","TTR ","","","","","","DESCRIPCION","si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select * From PROYETPTSK where ESTADO < 2 AND codigo = '" & sPRY & "' AND VPLAN = " & XVPLAN & " and tarea <> '" & SESSION("TAREA") & "' ORDER BY ORDEN,TSKORD  "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "NO HAY TAREAS DE DEPENDENCIA PARA SELECCIONAR"
      celdas = ofv.GenCelda("","vt colspan=2 ","CENTER","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
     do while not rs1.eof

     simagene = sImagents

     sinput = "<a CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYtasklst.asp?ACCION=SI&TSK=" & RS1(20)
     sinput = sinput  & " >" & simagene & "<b>&nbsp;" &  RS1(0) & "." & RS1(1) & ":&nbsp;&nbsp;" &  RS1(8) & "</b></a>" 
    
      celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")







    rs1.movenext
    loop
 
  end if


  END IF



 
    sHTML = sHTML & ofv.GenTabla("","aaA   ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  Agregar = ""
  Buscar  = ""

  sAccion = "PROYtasklst.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML





%>