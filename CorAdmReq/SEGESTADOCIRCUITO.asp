<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn1 = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  SESSION("SEGCLIENTE") = "NO"

 OBJTIPO = request.querystring("OBJTIPO")

 if OBJTIPO = "" then
    OBJTIPO = session("OBJTIPO")
 else
    session("OBJTIPO") = OBJTIPO
 end if 


  sopcion = request.querystring("opcion")

  if sopcion <> "" then
     session("sopc") = sopcion
  else
     sopcion = session("sopc")
  end if 


    celdas = ""



  'Encabezados
  if sopcion = "C" then
     filas =  ofv.GenCelda("","ttR COLSPAN=3 ","center","","","","","CIRCUITOS FINALIZADOS","si")
  else
     filas =  ofv.GenCelda("","ttR COLSPAN=3 ","center","","","","","CIRCUITOS ACTIVOS","si")
  end if
  filas = ofv.GenRow("","","","","22","","",filas,"si")

  if sopcion = "C" then
     StrSql = "Select * From SEGCIRCUITO WHERE FIN <> 'NO' AND OBJTIPO = '" & session("OBJTIPO") & "' ORDER BY FIN DESC, RESPONSABLE "
  ELSE
     StrSql = "Select * From SEGCIRCUITO WHERE FIN ='NO' AND OBJTIPO = '" & session("OBJTIPO") & "' ORDER BY COMIENZO DESC, RESPONSABLE "
  END IF
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     rs.move(CantRegMoverx)
  posi = 0
    filasx = filas
    filas = GENLISTA(RS)
    filasx = filasx & filas
  END IF


  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","-1","1","","",filasx,"si")


  'Botonera
  Agregar = ""


  Buscar = ""
  

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GENLISTA(RSX)

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"

    sImagenW = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=4>&#0154;</FONT></B>"

   xTrColor = "#FFFFFF"

  posi = 0
  prow = "N" 
  Do While (Not rsX.EOF) and (posi < CantRegMostrarx)

    sform = "REQ" & rsX(0)

    if prow = "S" then
       prow = "N"
       sr   = " bgcolor=#eeeeee "
    else 
       prow = "S"
       sr   = ""
    end if

       sr   = ""
       sc = "Ut"

    xlink = "<A class=NT href='&nbsp;'  "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('editarPRY.asp?codigo=" & rsx(5) & "&tipo=REQ'); return false;" & chr(34)
    xlink = xlink & " >" & sImagenW & "</A>"

    sInput = xlink

      celdas = ofv.GenCelda("","Ut rowspan=2 ","","1%","","","",sInput,"si")

       sInput2 = rsx(5)

       if session("OBJTIPO") = "CLI" then

          StrSql = "Select NOMBRE From CLIENTE WHERE CODIGO = '" & rsx(5) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rs2.eof then
             sInput2 = "" & RS2(0) & ""
          END IF

       else

          StrSql = "Select * From REQUERIMIENTO where codigo = '" & rsx(5) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn1,1,parametros)
          if NOT rs2.eof then
             sInput2 = "" & RS2(2) & ""
          end if

       end if




    xlink = "<A class=NT  href='SEGCKTESTADO.asp?CKT=" & RSx(6)
    xlink = xlink & "' "
    xlink = xlink & " >" & sInput2 & "</A>"

      celdas = celdas & ofv.GenCelda("",sc & " rowspan=2 style='border-right-style:solid;border-right-width:1;border-right-color:#888888;' ","","30%","","","",xlink,"si")


      sInput = "" & sImagenC & "&nbsp;" & rsx(8)
      celdas = celdas & ofv.GenCelda("","Ut","","3%","","","",sInput,"si")

    filas = ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")

  StrSql = "Select * From USUARIOS WHERE ID = '" & RSx(7) & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs2.eof THEN

       sInput = "" & sImagenI & " " & rs2(1)
  END IF
       celdas = ofv.GenCelda("",sc & " style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#888888;' ","","50%","","","",sInput,"si")


    filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","","","",celdas,"si")



                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F9F9F9" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


     formu = formu & ofv.GenForm("","","","","post","",filas,"si")


    posi = posi + 1
    rsX.Movenext
  loop

  IF SOPCION = "C" THEN

       cnbot = "<input type=submit class=bt value='VER CIRCUITOS ACTIVOS' >"
       Cons = ofv.gencelda("","aaa colspan=2","","","1%","","",cnbot,"si")
       sAccion = "SEGESTADOCIRCUITO.asp?OPCION=A"
       Cons = ofv.genform("cons","",sAccion,"","post","",Cons,"si")
  ELSE

       cnbot = "<input type=submit class=bt value='VER CIRCUITOS FINALIZADOS' >"
       Cons = ofv.gencelda("","aaa colspan=2","","","1%","","",cnbot,"si")
       sAccion = "SEGESTADOCIRCUITO.asp?OPCION=C"
       Cons = ofv.genform("cons","",sAccion,"","post","",Cons,"si")
  END IF




       celdas = Cons 

      sInput = "<b>" & sImagenC & "&nbsp;Circuito</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","25%","","","",sInput,"si")
      sInput = "<b>" & sImagenI & "&nbsp;Usuario Originante</b>"
      celdas = celdas & ofv.GenCelda("","vt" ,"","25%","","","",sInput,"si")
      celdas = celdas & ofv.GenCelda("","vt  " ,"","","","","","&nbsp;","si")

    formu = formu & ofv.GenRow("","aaa height=30","","","","","",celdas,"si")

  GENLISTA = FORMU

END FUNCTION



%>