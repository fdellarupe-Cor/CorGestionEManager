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



 sopcion = request.querystring("opcion")

 if sopcion <> "" then
    session("OPC") = sopcion
 else
    sopcion = session("OPC")
 end if 

 sacc = request.querystring("accion")

 if sacc = "SI" THEN

    stsk = request.querystring("tsk")
    sxxx = GENDEPEN(stsk)
 END IF




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
 




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"


  StrSql = "Select * From PROYtareas where codigo = '" & session("NCliente") & "' and TAREA = '" & SESSION("TAREA") & "'  AND VPLAN = " & SESSION("VPLAN")
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then

     XTAREA = rs1(4) & "." & RS1(5)

  END IF






      celdas = ofv.GenCelda("","TTT colspan=4","","","","","","DEPENDENCIA DE TAREAS","si")
      filas  = ofv.GenRow("","","","","20","","",celdas,"si")


      sInput = XTAREA
      celdas = ofv.GenCelda("","VT COLSPAN=4 ","","","","","","TAREA DEPENDIENTE:&nbsp;" & sInput ,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","TTR COLSPAN=4  ","CENTER","","","","","LISTA&nbsp;DE&nbsp;TAREAS","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt ","CENTER","1%","","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","vt ","CENTER","30%","","","","TAREA","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","30%","","","","PROYECTO","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","30%","","","","ESTADO","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  StrSql = "Select * From PROYtareasLST where codigo = '" & session("NCliente") & "' and TAREA = '" & SESSION("TAREA") & "' order by orden "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "NO HAY TAREAS DE DEPENDENCIA INFORMADAS"
      celdas = ofv.GenCelda("","vt colspan=4 ","CENTER","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
     do while not rs1.eof

  StrSql = "Select * From PROYECTO WHERE CODIGO = '" & RS1(4) & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS2.EOF THEN
     XVPLAN  = RS2(17)
     XPRYNOM = RS2(2)
  END IF

  StrSql = "Select * From PROYTAREAS WHERE TAREA = '" & RS1(5) & "' AND CODIGO = '" & RS1(4) & "' AND VPLAN = " & session("vplan")
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS2.EOF THEN





     XTSKNOM = RS2(4) & "." & RS2(5)
  END IF
      SINPUT = XTSKNOM
      IF RS1(1) <> RS1(4) THEN
         SINPUT = SIMAGEN6OT & "&nbsp;" & SINPUT
      ELSE
         SINPUT = SIMAGEN6N & "&nbsp;" & SINPUT
      END IF

    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarRegistro.asp?ID=" & rs1(0)
      sHref = sHref & "&Tabla=PROYtareasLST&volver=" & vuelta & "&Eliminar=NSNC"
      CeldaS = ofv.BotonEliminar(" Dependencia: " & XTSKNOM,sHref)
    else
      CeldaS = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if
    
      celdas = celdas & ofv.GenCelda("","vt ","CENTER","","","","",SINPUT,"si")

      SINPUT = XPRYNOM
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","",SINPUT,"si")
        if rs1(6) = 0 then sInput = "ACTIVO"
        if rs1(6) = 1 then sInput = "FINALIZADO"
        if rs1(6) = 2 then sInput = "CANCELADO"
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






    rs1.movenext
    loop
 
  end if



 
    sHTML = sHTML & ofv.GenTabla("","aaA   ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center>"


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "PROYtasklstbus.asp?TIPO=" & SESSION("OPC") & "&PRY=" & session("NCliente")
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar  = ""
  sAccion = "PROYtaskprop.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)



  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENDEPEN(XTSK)

  StrSql = "Select MAX(orden) From PROYtareasLST where codigo = '" & session("NCliente") & "' and TAREA = '" & SESSION("TAREA") & "' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then
     xord = 10
  else
     xord = rs1(0) + 10
  end if 

  StrSql = "Select codigo From PROYtareas where  TAREA = '" & xtsk & "' "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then
     xtskpry = "NO"
  else
     xtskpry = rs1(0)
  end if 



        STRSQL = "INSERT INTO PROYTAREASLST (CODIGO,TAREA,ORDEN,DCODIGO,DTAREA,ESTADO) VALUES ("  
        STRSQL = STRSQL & "'" & session("NCliente") & "','" & session("TAREA") & "','" & XORD & "','" & XTSKPRY & "','" & XTSK & "','0')" 
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql





END FUNCTION



%>