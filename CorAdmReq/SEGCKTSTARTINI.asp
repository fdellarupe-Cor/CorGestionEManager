<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

'  sHTML = sHTML  & "<script language=javascript> "
'  sHTML = sHTML  & "function AbrirCons(Datos) { "
'  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
'  sHTML = sHTML  & " } "
'  sHTML = sHTML  & "</script>"

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


''' INICIA EL CIRCUITO 
  WXX = GENCKTINI()



''' GENERA EL PASO INICIAL

  StrSql = "Select * From WKFCKTSEG WHERE INI=1 AND WKFCOD = '" & SESSION("WKF") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     SEGCKTPSC = GTCKTPSC()
     XMODEL  = RS(3)
     XRTIPO  = RS(6)
     XRCOD   = RS(7)
     XPASCOD = RS(12)
     XPDSC   = RS(4) & ":" & RS(5) 
     XPDSC1   = RS(4) & ":" & RS(5) 

    SEGCKTPSCSIG = GTCKTPSC() & "A"
    WXX = GENCKTPSC(SEGCKTPSC,SEGCKTPSCSIG,XMODEL,"SI",XRTIPO,XRCOD,SESSION("USUARIO"),XPDSC)


    StrSql = "Select * From WKFCKTSEG WHERE WKFPASCOD = '" & XPASCOD & "' "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
       XMODEL  = RS2(3)
       XRTIPO  = RS2(6)
       XRCOD   = RS2(7)
       XPASCOD = RS2(3)
       XPDSC   = RS2(4) & ":" & RS2(5) 
       XRASG   = "NO"
       IF XRTIPO = "USR" THEN
          XRASG   = XRCOD
       END IF

''' GENERAR EL PRIMER PASO OK
       WXX = GENCKTPSC(SEGCKTPSCSIG,"NO",XMODEL,"NO",XRTIPO,XRCOD,XRASG,XPDSC)


    END IF
    call ofv.cerrarconsulta(rs2)

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

       sinput = "PASO :" & XPDSC1 
       celdas = ofv.GenCelda("","TT","","","50%","","",sInput,"si")



       sinput = "INICIADO Y CONFORMADO " & NOW()
       celdas = celdas & ofv.GenCelda("","CC","","","50%","","",sInput,"si")




    filas =  ofv.GenRow("","","","","","","",celdas,"si")

       sinput = "PASO :" & XPDSC 
       celdas = ofv.GenCelda("","TT","","","50%","","",sInput,"si")



       sinput = "INICIADO " & NOW()
       celdas = celdas & ofv.GenCelda("","CC","","","50%","","",sInput,"si")




    filas = FILAS &  ofv.GenRow("","","","","","","",celdas,"si")

 
       sinput = "&nbsp;"
       celdas = ofv.GenCelda("","","center","","50%","","",sInput,"si")



       sinput = "<a class=tt  href=SEGCIRCUITO.asp"
       sinput = sinput  & " ><b>VOLVER A INICIO DE CIRCUITOS</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","50%","","",sInput,"si")




    filas = FILAS & ofv.GenRow("","","","","","","",celdas,"si")

    sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")




  ELSE
   SHTML = SHTML & "<P>ERROR EN EL INICIO DEL CIRCUITO</P>"

  END IF





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENCKTINI()

    SEGCKT = GTCKT()

    SESSION("SEGCKT") = SEGCKT

    strsql = "insert into SEGCIRCUITO (MODEL,COMIENZO,FIN,OBJTIPO,OBJCOD,SEGCKT,RESPONSABLE,CKTDSC) values "
    STRSQL = STRSQL & " ('" & SESSION("WKF") & "','" & NOW() & "','NO','CLI','" & SESSION("SEGCLIENTE") & "','" & SEGCKT & "','" & SESSION("USUARIO") & "','" & session("WKFDSC") & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION

FUNCTION GENCKTPSC(XPSC,XSIG,XMODEL,XFIN,XRTIPO,XRCOD,XRASG,XPDSC)

    SEGCKTPSC = XPSC

    SESSION("SEGCKTPSC") = SEGCKTPSC

    XCOM = NOW()
    IF XFIN = "SI" THEN
       XFIN = NOW()
    ELSE
       XFIN = "NO"
    END IF

    strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC) values "
    STRSQL = STRSQL & " ('" & SESSION("SEGCKT") & "','" & SESSION("SEGCKTPSC") & "','" & XMODEL & "','" & NOW() & "','" & XFIN & "','" & XSIG & "','" & XRTIPO & "','" & XRCOD & "','" & XRASG & "','" & XPDSC & "')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

END FUNCTION






%>