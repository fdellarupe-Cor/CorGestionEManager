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

 sopcion = request.querystring("opcion")
 VPLAN = request.querystring("VPLAN")

   StrSql = "Select * From PROYPLAN where codigo = '" & session("NCliente") & "' AND VPLAN = " & VPLAN 
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs1.eof then
     VPLAND = RS1(2) & "." & RS1(3)
  END IF
  call ofv.cerrarconsulta(rs1)
    YXFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 

 IF SOPCION = "HT" THEN
     SESSION("VPLAN") = VPLAN
     SESSION("VPLAND") = VPLAND




    response.redirect "PROYTASK.asp"

 END IF 

 IF SOPCION = "HE" THEN
     SESSION("VPLAN") = VPLAN
     SESSION("VPLAND") = VPLAND




    response.redirect "PROYEST.asp"

 END IF 


 IF SOPCION = "APX" THEN
     SESSION("VPLAN") = VPLAN
     SESSION("VPLAND") = VPLAND

    IF VPLAN = "9999" THEN
       XXXX = OWKF.PRYLNCH(session("NCliente"),"ACTIVACION","K")
    ELSE
       XXXX = OWKF.PRYLNCH(session("NCliente"),"ACTIVACION","P")
    END IF

    StrSql = "Select VPLAN From PROYECTO where codigo = '" & session("NCliente") & "' "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS.EOF THEN

     SESSION("VPLAN") = RS(0)

    StrSql = "Select tarea From PROYETPTSK where codigo = '" & session("NCliente") & "' and VPLAN = " & RS(0) & " ORDER BY ORDEN desc,TSKORD DESC "
    set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT rs5.EOF THEN
       OWKF.YFMAXTSK = rs5(0)

       SXFSTARTD = OWKF.GETFENDV(rs5(0), session("NCliente"))

       XIIX = OWKF.TIX

       FEPLAN = OWKF.ZXMFMAX


       StrSql = "update PROYECTO set FECHAENT = 0" & FEPLAN & " where codigo = '" & session("NCliente") & "' "
   '    RESPONSE.WRITE STRSQL & "<BR>" 
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
       


    end if


    call ofv.cerrarconsulta(rs5)

    END IF

    response.redirect "PROYRPLLST.asp"

 END IF 

 IF SOPCION = "CPX" THEN
     SESSION("VPLAN") = VPLAN
     SESSION("VPLAND") = VPLAND

    IF VPLAN = "9999" THEN
       XXXX = OWKF.PRYPLANRVRS(session("NCliente"),"CANCELACION","0",VPLAN,"NO",YXFSTART,"KC")
    ELSE
       XXXX = OWKF.PRYPLANRVRS(session("NCliente"),"CANCELACION","0",VPLAN,"NO",YXFSTART,"C")
    END IF




    response.redirect "PROYRPLLST.asp"

 END IF 

 IF SOPCION = "CP" THEN
     SESSION("VPLAN") = VPLAN
     SESSION("VPLAND") = VPLAND

     VDISP = "CANCELACION"
     VENAB = "S"




 END IF 

 IF SOPCION = "AP" THEN
     SESSION("VPLAN") = VPLAN
     SESSION("VPLAND") = VPLAND

     VDISP = "ACTIVACION"
     VENAB = "S"

     StrSql = "Select * From PROYETPTSK where USRRESP = 'NO' AND codigo = '" & session("NCliente") & "' AND VPLAN = " & VPLAN 
     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rs1.eof then
        VENAB = "N"
     END IF
     call ofv.cerrarconsulta(rs1)


 END IF 


    sHTML = sHTML & "<center>"

    sImagen  = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
    sImagenA = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>n</FONT>"
    sImagenN = "<FONT FACE=WINGDINGS COLOR=#88bbdd SIZE=2>n</FONT>"


  IF  sopcion = "AP" OR sopcion = "CP" THEN

      celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>" & VDISP & " DE LA NUEVA PLANIFICACION</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt  ","","","","","","FECHA:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
      sinput = ofv.fechador
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      XMSG = "&nbsp;"
      celdas = ofv.GenCelda("","vt  COLSPAN=3 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","CENTER","","","","",XMSG,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")





         IF VENAB = "S" THEN
             XMSG = "<B>CONFIRMA LA " & VDISP & " DE LA NUEVA PLANIFICACION PARA EL PROYECTO</B><BR>" & VPLAND
         ELSE
             XMSG = "<B>NO ES POSIBLE LA " & VDISP & " DE LA NUEVA PLANIFICACION PARA EL PROYECTO<BR>PORQUE CONTIENE TAREAS SIN ASIGNAR</B><BR>" & VPLAND
         END IF

             celdas = ofv.GenCelda("","vt  COLSPAN=3","CENTER","","","","",XMSG,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModificar then
      sAccion = "PROYRPLDSP.asp?OPCION=" & SOPCION & "X&VPLAN=" & VPLAN
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


   IF VENAB = "S" THEN
     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
     sinput = sinput  & " ><b>CONFIRMA</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")
   ELSE
     celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
   END IF

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRPLLST.asp"
     sinput = sinput  & " ><b>CANCELA</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")








    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#dddddd,direction=125);' ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"




    end if



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML







%>