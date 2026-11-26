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


 IF SOPCION = "AT" THEN
    XUSR = session("usuario")
    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900"))
    VPDESCRI = REQUEST.FORM("VPDESCRI") 
    VPOBSERV = REQUEST.FORM("VPOBSERV") 


    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       XVPLAN = RS2(17)
    end if
    call ofv.cerrarconsulta(rs2)

    XXXX = OWKF.PRYPLAN(session("NCliente"),VPDESCRI,XVPLAN,"NEXT",XUSR,XFSTART,"P")

    StrSql = "UPDATE PROYPLAN SET OBSERVACIONES = '" & VPOBSERV & "' where codigo = '" & session("NCliente") & "' AND VPLAN = " & XXXX
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


    response.redirect "PROYRPLLST.asp"

 END IF 

 IF SOPCION = "CPX" THEN
    XUSR = session("usuario")
    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900"))




    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       XVPLAN = RS2(17)
    end if
    call ofv.cerrarconsulta(rs2)

    StrSql = "Select * From PROYPLAN where codigo = '" & session("NCliente") & "' AND VPLAN = " & XVPLAN
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       XVPLAND = RS2(3)
    end if
    call ofv.cerrarconsulta(rs2)

    SESSION("VPLAN") = XVPLAN
    SESSION("VPLAND") = XVPLAND

    VPDESCRI = XVPLAND & " (COPIA)"

    XXXX = OWKF.PRYPLAN(session("NCliente"),VPDESCRI,XVPLAN,"9999",XUSR,XFSTART,"K")


    response.redirect "PROYRPLLST.asp"

 END IF 

 IF SOPCION <> "VT" AND SOPCION <> "CP" THEN sopcion = "VP"


    sHTML = sHTML & "<center>"

    sImagen  = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
    sImagenA = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>n</FONT>"
    sImagenN = "<FONT FACE=WINGDINGS COLOR=#88bbdd SIZE=2>n</FONT>"



  IF  sopcion = "VP" OR sopcion = "VT"  OR sopcion = "CP"  THEN

      celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;GENERACION DE NUEVA PLANIFICACION","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt  ","","","","","","FECHA:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
      sinput = ofv.fechador
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","aaa bgcolor='#f0f0f0' ","","","10","","",celdas,"si")

      XMSG = "&nbsp;"
      celdas = ofv.GenCelda("","vt  COLSPAN=3 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","CENTER","","","","",XMSG,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      if sopcion = "VP" OR sopcion = "CP"  then


      if sopcion = "VP"  then
             XMSG = "CONFIRMA LA GENERACION DE UNA NUEVA PLANIFICACION<BR>PARA EL PROYECTO ??"
      ELSE
             XMSG = "CONFIRMA LA GENERACION DE UNA COPIA DE LA PLANIFICACION<BR>VIGENTE ACTUALMENTE PARA EL PROYECTO ??"
      END IF




             celdas = ofv.GenCelda("","ut valign='middle' COLSPAN=3 style='text-align:justify;border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","","","","",XMSG,"si")
             filas = filas & ofv.GenRow("","","","","90","","",celdas,"si")




    if bModificar then

      if sopcion = "VP"  then
         sAccion = "PROYRPLGEN.asp?OPCION=VT"
      ELSE
         sAccion = "PROYRPLGEN.asp?OPCION=CPX"
      END IF

      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if


    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


      FilaS = ofv.GenTabla("","","","100%","","0","","-1","3","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
     sinput = sinput  & " ><b>CONFIRMA</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRPLLST.asp"
     sinput = sinput  & " ><b>CANCELA</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")


   else




     StrSql = "Select MAX(VPLAN) FROM PROYPLAN where codigo = '" & session("NCliente") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rs.eof then
        YVPLANN = RS(0) + 1
     ELSE
        YVPLANN = 2
     END IF
     call ofv.cerrarconsulta(RS)

        SINPUTT2 = YVPLANN



          celdas = ofv.GenCelda("","vt  ","","","","","","NUEVO PLAN: " & YVPLANN ,"si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
          sInput = ofv.GenerarInput("VPDESCRI","NUEVA PLANIFICACION","agregar","Usuarios","80","dtb","")
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


          celdas = ofv.GenCelda("","vt COLSPAN=3 style='border-BOTTOM-style:ridge;border-BOTTOM-width:2;border-BOTTOM-color:#f7aa22;'","CENTER","","","","","<BR>OBSERVACIONES","si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      nCol = 80

      wdescri = "Observaciones" 


    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
    t3b = t3b & "name=VPOBSERV cols=" & nCol & " rows=4 "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 195) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    sInput = t3b

          celdas = ofv.GenCelda("","vt COLSPAN=3","CENTER","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



    if bModificar then
      sAccion = "PROYRPLGEN.asp?OPCION=AT"
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


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
     sinput = sinput  & " ><b>ingresar</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRPLLST.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")



      end if



    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;' ","","70%","","0","","-1","3","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"




    end if



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML







%>