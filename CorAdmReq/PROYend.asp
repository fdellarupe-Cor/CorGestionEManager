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

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if


 sopcion = request.querystring("opcion")

 if sopcion <> "SI" THEN

     sopcion = "VP"

 END IF



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"





    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then
          XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
          XFINIT  = rs(4) 
          XFSTS   = rs(11) 

 

      celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>FINALIZACION DEL PROYECTO</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","VT  ","","","","","","PROYECTO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","VT","","","","","",RS(2),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","VT  ","","","","","","FECHA DE INICIO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","VT","","","","","",CDATE(CLNG(XFINIT) + clng(cdate("01-01-1900"))),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")







      celdas = ofv.GenCelda("","VT COLSPAN=3 VALIGN=BOTTOM ","center","","30","","","MENSAJE DE FINALIZACION","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  IF  sopcion = "VP" THEN


         wdescri = "FINALIZACION DEL PROYECTO CON FECHA: " & OFV.FECHADOR()



    t3b = "<font class=at><textarea class=dt "
    t3b = t3b & "name=" & rs(10).name & " cols=90 rows=10 "
    if bModificar then
 '''     t3b = t3b & " onblur=Usuarios.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    CeldaS = ofv.GenCelda("","CC colspan=3","center","100%","","","",t3b,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "PROYEND.ASP?OPCION=SI"
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
     sinput = sinput  & " ><b>&nbsp;FINALIZAR&nbsp;proyecto&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

 

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYPROP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")
     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")




   else

    wdescri = request.form("DESCRIPCION") 


    STRSQL = "UPDATE PROYECTO SET FECHAFIN = " & XFSTART & ",ESTADO = 13 WHERE CODIGO = '" & session("NCliente") & "'  "
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

    strsql = "UPDATE REQUERIMIENTO SET FECHAFIN = " & XFSTART & ",ESTADO = 6 WHERE codigo = '" & session("NCliente") & "' "
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)


    t3b = "<font class=at><textarea class=dt "
    t3b = t3b & " cols=90 rows=10 "
    t3b = t3b & " readonly >"
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    CeldaS = ofv.GenCelda("","CC colspan=3","center","100%","","","",t3b,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


 
      celdas = ofv.GenCelda("","VT  ","","","","","","FECHA DE FINALIZACION","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","VT","","","","","",CDATE(clng(XFSTART) + clng(cdate("01-01-1900"))),"si")
      filas = filas & ofv.GenRow("","AAA VALIGN=BOTTOM ","","","30","","",celdas,"si")



    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


      FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")



     celdas = ofv.GenCelda(""," ","","10%","","","","&nbsp;","si")

 

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYPROP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")
     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")



    END IF


    end if


    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#DDDDDD,direction=125);' ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





 
  IF session("ModMain2") = "CL" THEN
     sinput = "<a  class=tt href=Clientes.asp"
     sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ELSE
     sinput = "&nbsp;"
     celdas =  ofv.GenCelda("","","","1","","","",sInput,"si")
  END IF


  celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
  filas = ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","aa ","","98%","","0","","0","0","","",filas,"si")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML




%>