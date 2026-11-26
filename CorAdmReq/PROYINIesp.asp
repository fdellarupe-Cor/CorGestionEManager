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



  IF  sopcion = "VP" THEN

    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then
          XFSTART = CLNG(cDATE("06-03-2005"))
          XFSTS = rs(11) 
          XFFCH = RS(18)
          IF XFFCH > 0 THEN XXX = OWKF.ADMFERIADOS("C",0)
 

      celdas = ofv.GenCelda("","cc colspan=3","","","","","",sImagen & "&nbsp;<BIG><B>INICIO DEL PROYECTO</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","TT  ","center","","","","","PROYECTO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","CC","","","","","",RS(2),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","TT  ","center","","","","","FECHA DE INICIO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
 
      celdas = celdas & ofv.GenCelda("","CC","","","","","",CDATE(XFSTART),"si")
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

      celdas = ofv.GenCelda("","TT COLSPAN=3 ","center","","","","","MENSAJE DE INICIO","si")
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
    CeldaS = ofv.GenCelda("","CC colspan=3","center","100%","","","",t3b,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "PROYINIesp.ASP?OPCION=SI"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    filas = ofv.GenRow("","","","","","","",formulario0,"si")


  ELSE

    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    XMSG = "<BIG>EL PROYECTO NO ESTA EN CONDICIONES DE INICIAR PORQUE<BR>SU TAREA INICIAL DEPENDE DE OTRO PROYECTO</BIG>"

    Celda = ofv.GenCelda("","VT colspan=3","center","100%","","","",XMSG,"si")
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
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

   ELSE
     celdas = ofv.GenCelda("","","","10%","","","","&nbsp;","si")
   END IF

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYPROP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")
     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")

    end if


   else

          XFSTART = CLNG(cDATE("06-03-2005"))
    wdescri = request.form("DESCRIPCION") 

    XXXX = PRYLNCH(session("NCliente"),wdescri,"S",XFSTART)


   ''' FIN LOOP ACTUALIZACION 

   response.redirect "proyest.asp"


  '   sinput = "<a  class=tt href=proyest.asp"
   '  sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
   '  fila = sinput



    END IF





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


Function PRYLNCH(YPRY, YDESCRI, YLNCH,yfecstart)





 Set CN3 = OFV.conectar(OFV.strconn6)

     YXFSTART = yfecstart - CLng(CDate("01-01-1900"))
     YXFCURR = YXFSTART



        ''' estado = 12

   If YLNCH <> "P" And YLNCH <> "K" Then

     xxxx = OWKF.PRYPLANRVRS(YPRY, YDESCRI, "0", "1", "NO", YXFSTART, "R")

   End If




      YNPLAN = Session("VPLAN")


     StrSql = "Select * From PROYecto where codigo = '" & YPRY & "'  "
     Set RSXP = OFV.crearconsultaEx(StrSql, CN3, 1, "")
     If Not RSXP.EOF Then
           YXASIG = RSXP(14)
           YXFSTS = RSXP(11)
           YXFFCH = RSXP(18)
           YVPLAN = RSXP(17)
           If CLng(YXFFCH) > 0 Then XXX = OWKF.ADMFERIADOS("C", 0)
     End If

     If YPRY <> Session("NCliente") And YLNCH <> "K" Then
        YNPLAN = YVPLAN
     End If


      If YLNCH = "P" Or YLNCH = "K" Then

''''        xxxx = OWKF.PRYPLANRVRS(YPRY, YDESCRI, "0", "1", "NO", YXFSTART, "R")

         StrSql = "Select TAREA,TSKSPL,ID From PROYtareas where codigo = '" & YPRY & "' and VPLAN = " & YNPLAN & " AND TSKSPL <> 'NO' order by ID "
         Set RSXP1 = OFV.crearconsultaEx(StrSql, CN3, 1, "")
         If Not RSXP1.EOF Then
            Do While Not RSXP1.EOF
               XNVTAREA = RSXP1(0)
               XNVTSKSPL = RSXP1(1)
               StrSql = "UPDATE PROYtareas SET TSKSPL = 'NO' where ID = " & RSXP1(2)
               Call OFV.crearconsultaEx(StrSql, CN3, 1, "")

               StrSql = "UPDATE PROYtareasLST SET DTAREA = '" & XNVTSKSPL & "' where DTAREA = '" & XNVTAREA & "' AND DCODIGO = '" & YPRY & "' "
               Call OFV.crearconsultaEx(StrSql, CN3, 1, "")

            RSXP1.MOVENEXT
            Loop


         End If
         Call OFV.cerrarconsulta(RSXP1)

      End If


        ''' LOOP DE ACTUALIZACION DE TAREAS

  YWETP = 0
  YWTSK = 0
           YXFREAL = YXFCURR

   StrSql = "Select * From PROYetapa where codigo = '" & YPRY & "' AND VPLAN = " & YNPLAN & " order by orden "
   Set RSXP = OFV.crearconsultaEx(StrSql, CN3, 1, "")
   If Not RSXP.EOF Then
      Do While Not RSXP.EOF

      YWETP = YWETP + 1

      YXFINIETP = YXFCURR

   StrSql = "Select * From PROYtareas where codigo = '" & YPRY & "' and etapa = '" & RSXP(2) & "' AND VPLAN = " & YNPLAN & " order by orden "
   Set RSXP1 = OFV.crearconsultaEx(StrSql, CN3, 1, "")
   If Not RSXP1.EOF Then
      Do While Not RSXP1.EOF




     XIIX = OWKF.GETTSKCLC(RSXP1, YXFSTART, YXFREAL, YXFFCH)




     YXFREAL = ATABTSKW(XIIX, 1)
     YTSKFS = ATABTSKW(XIIX, 2)
     YTSKFE = ATABTSKW(XIIX, 3)
     YTSKFSR = ATABTSKW(XIIX, 4)
     YTSKFER = ATABTSKW(XIIX, 5)
     YTSKST = ATABTSKW(XIIX, 6)
     YFSPLAN = ATABTSKW(XIIX, 7)
     YFEPLAN = ATABTSKW(XIIX, 8)
     YFSREAL = ATABTSKW(XIIX, 9)
     YFEREAL = ATABTSKW(XIIX, 10)
     YXFFDU = ATABTSKW(XIIX, 11)
     YXFFTI = ATABTSKW(XIIX, 12)
     YPADV = ATABTSKW(XIIX, 13)
     YXAGRD = ATABTSKW(XIIX, 14)
     YXAGRR = ATABTSKW(XIIX, 15)
     YXMODA = ATABTSKW(XIIX, 16)
     YXTDUR = ATABTSKW(XIIX, 17)
     YXUDUR = ATABTSKW(XIIX, 18)
     YXTDUP = ATABTSKW(XIIX, 19)
     YXUDUP = ATABTSKW(XIIX, 20)








     YXFINI = YTSKFS
     YXFCURR = YXFREAL
   If YLNCH = "P" Or YLNCH = "K" Then
      YXESTADO = RSXP1(18)
   Else
      YXESTADO = "0"
   End If

     YWTSK = YWTSK + 1
     If YWETP = 1 And YWTSK = 1 Then
        YXFPRYSTRT = YXFINI
        YXFPRYFTSK = RSXP1(3)
     End If

     StrSql = "UPDATE PROYTAREAS SET FECHASPLAN = " & YXFINI & ",FECHAEPLAN = " & YXFREAL & ",ESTADO = " & YXESTADO & " WHERE ID = " & RSXP1(0)
  ''   Response.WRITE StrSql & "<BR>"
         Call OFV.crearconsultaEx(StrSql, CN3, 1, "")


    


     RSXP1.MOVENEXT
     Loop
     Call OFV.cerrarconsulta(RSXP1)

   End If

    YXFENDETP = YXFCURR

   

     StrSql = "UPDATE PROYETAPA SET FECHASPLAN = " & YXFINIETP & ",FECHAEPLAN = " & YXFREAL & " WHERE ID = " & RSXP(0)
     Call OFV.crearconsultaEx(StrSql, CN3, 1, "")
     RSXP.MOVENEXT
     Loop
      Call OFV.cerrarconsulta(RSXP)

   End If

   If YLNCH = "P" Or YLNCH = "K" Then

     xxxx = OWKF.PRYPLANRVRS(YPRY, "ACTIVACION", YVPLAN, YNPLAN, Session("USUARIO"), YXFSTART, YLNCH)
       
   Else

     YNVPLAN = OWKF.PRYPLAN(YPRY, "PLAN INICIAL", "0", "1", YXASIG, YXFPRYSTRT, "S")

     If YLNCH = "S" Then
      '''''  dar inicio a la primera tarea del pry


        YXESTADO = "1,FECHASREAL = " & YXFPRYSTRT

        StrSql = "UPDATE PROYTAREAS SET ESTADO = " & YXESTADO & " WHERE TAREA = '" & YXFPRYFTSK & "' AND VPLAN = 1 "
        Call OFV.crearconsultaEx(StrSql, CN3, 1, "")

     End If

     StrSql = "UPDATE PROYECTO SET FECHAENT = " & YXFCURR & ",ESTADO = 12 WHERE CODIGO = '" & YPRY & "'  "
     Call OFV.crearconsultaEx(StrSql, CN3, 1, "")

   End If

         ''' FIN LOOP ACTUALIZACION
   
    Call OFV.cerrarconn(CN3)



End Function

%>