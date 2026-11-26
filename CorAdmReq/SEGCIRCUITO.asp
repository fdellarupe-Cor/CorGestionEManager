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
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 OBJTIPO = request.querystring("OBJTIPO")
 if OBJTIPO = "" then
    OBJTIPO = session("OBJTIPO")
 else
    session("OBJTIPO") = OBJTIPO
 end if 

  SESSION("SEGCLIENTE") = "NO"



' SEGWHERE = " WHERE INI = 1 AND INICIO = 'S' AND ( (OBJTIPO = 'GEN' OR OBJTIPO = '" & session("OBJTIPO") & "' ) AND ( (CLASE = 'M')   "

 SEGWHERE = " WHERE INI = 1 AND INICIO = 'S' AND ( (OBJTIPO = 'GEN'  )   "

  StrSql = "Select * From USUARIOS WHERE ID = '" & Session("Usuario") & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs.eof THEN
     SEGWHERE = SEGWHERE & "  AND (  ( "  
     SEGWHERE = SEGWHERE & " ((RTIPO = 'USR') AND (RCODIGO = '" & Session("Usuario") & "')) "  
     SEGWHERE = SEGWHERE & " OR ((RTIPO = 'GRP') AND (RCODIGO = '" & SESSION("SEGGRP") & "')) "  
     IF RS(21) > 0 THEN
        SEGWHERE = SEGWHERE & " OR ((RTIPO = 'SEC') AND (RCODIGO = '" & RS(21) & "')) "  
     END IF
     SEGWHERE = SEGWHERE & " ) ) "  

  END IF 
     SEGWHERE = SEGWHERE & "  ) "  



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

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si") & "<BR>"



    celdas = ""



  'Encabezados
  filas =  ofv.GenCelda("","tt colspan=4","center","","","","","CIRCUITOS DISPONIBLES","si")
  filas =  filas & ofv.Encabezados("CODIGO","DESCRIPCION","&nbsp;")


  StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
 ''' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

'' RESPONSE.WRITE STRSQL

  if not rs.eof then
  posi = 0
  Do While (Not rs.EOF)

    sform = "WKF" & rs(0)


       sc = "CC"



       celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")



       sInput = rs(0)
       celdas = celdas & ofv.GenCelda("",sc,"","15%","","","",sInput,"si")


       sInput = rs(1)
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")


       sinput = "<a class=tt  href=SEGCKTSTART.asp?WKF=" & RS(0)
       sinput = sinput  & " ><b>iniciar</b></a>" 
       celdas = celdas & ofv.GenCelda("","tt","center","","10%","","",sInput,"si")



    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



    posi = posi + 1
    rs.Movenext
  loop
  END IF
  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",filas,"si")


  'Botonera
  Agregar = ""


  Buscar = ""
  


  
 ''' sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML





%>