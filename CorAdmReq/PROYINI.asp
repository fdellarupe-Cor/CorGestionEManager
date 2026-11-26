<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(" ")



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



  IF  sopcion = "VP" THEN

    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then
          XFSTART = CLNG(DATE())
          XFSTS = rs(11) 
          XFFCH = RS(18)
          IF XFFCH > 0 THEN XXX = OWKF.ADMFERIADOS("C",0)
 

      celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;<B>INICIO DEL PROYECTO</B>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vT  ","","","","","","PROYECTO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","ut","","","","","",RS(2),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vT  ","","","","","","FECHA DE INICIO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","ut","","","","","",CDATE(XFSTART),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      XPRYINI = "NO" 
       StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' and VPLAN = 0 ORDER BY ORDEN,TSKORD,FECHASPLAN DESC "
  ''     response.write "1 " & strsql & "<br>" 
       set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       IF NOT rs5.EOF THEN

''''''' Controla que la tarea inicial inicie en modo secuencia '' estado = no iniciada (0) y modo inicio = secuencia (1) '''               

          IF rs5(4) = 0 AND rs5(6) = 1 THEN
             XPRYINI = "SI" 
          END IF

       END IF  
       call ofv.cerrarconsulta(rs5)



   IF XPRYINI = "SI" THEN

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt COLSPAN=3 style='border-style:solid;border-width:1;border-color:#888888;background-color:#f9f9f9;' ","center","","","","","MENSAJE DE INICIO","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      if len(wdescri) = 0 then wdescri = "Sin Descripcion" 

      if not  isnull(rs(10)) then 
         wdescri = rs(10)
      end if 


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
    CeldaS = ofv.GenCelda("","ut colspan=3 style='border-style:solid;border-width:1;border-color:#888888;'  ","center","100%","","","",t3b,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "PROYINI.ASP?OPCION=SI"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")


  ELSE

    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    XMSG = "<BIG>EL PROYECTO NO ESTA EN CONDICIONES DE INICIAR PORQUE<BR>SU TAREA INICIAL DEPENDE DE OTRO PROYECTO</BIG>"

    Celda = ofv.GenCelda("","VT colspan=3 aaa style='border-style:solid;border-width:1;border-color:#888888;'  ","center","100%","","","",XMSG,"si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")



  END IF



    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


      FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

   IF XPRYINI = "SI" THEN

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
     sinput = sinput  & " ><b>&nbsp;INICIAR&nbsp;proyecto&nbsp;</b></a>" 


     sInput  = ofv.GenerarInput("boton","INICIAR","button","","10","bt","Usuarios.submit();")
     sAccion = "none"
     btcel   = ofv.GenCelda("","UT TITLE='Iniciar proyecto.' ","","","","","",sInput,"si")
     celdas  = ofv.GenForm("init","",sAccion,"","post","",btcel,"si")




   ELSE
     celdas = ofv.GenCelda("","","","10%","","","","&nbsp;","si")
   END IF

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYPROP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 


     sInput  = ofv.GenerarInput("boton","VOLVER","submit","","10","bt","")
     sAccion = "PROYPROP.asp"
     btcel   = ofv.GenCelda("","UT TITLE='FINALIZAR EL PROYECTO.' ","","","","","",sInput,"si")
     celdas  = celdas & ofv.GenForm("init2","",sAccion,"","post","",btcel,"si")





     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;' ","","","","0","","4","0","","",fila,"si")

    end if


   else

    wdescri = request.form("DESCRIPCION") 

    XXXX = OWKF.PRYLNCH(session("NCliente"),wdescri,"S")


   ''' FIN LOOP ACTUALIZACION 

   response.redirect "PROYPROP.asp"


  '   sinput = "<a  class=tt href=proyest.asp"
   '  sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
   '  fila = sinput



    END IF





    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#888888;' ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"






  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML




%>