<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = request.querystring("CantRegMover")


  wperfil = request.querystring("perfil")


  IF wperfil <> "" THEN 
     SESSION("wperfil") = wperfil
  ELSE
     wperfil = SESSION("wperfil")
  END IF


  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_perfil=" & Wperfil

  sHTML = ofv.FormHeader(session("FormN"))

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "resizable=yes,menubar=no,toolbar=no,height=350,width=350" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"


  'Encabezados
  Formulario = ""

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>L</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=4>&#064;</FONT></B>"



  dim stabla(12)


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  sopcion = request.querystring("opcion")

  if sopcion = "SI" THEN
     sECTOR = request.querystring("SECTOR")
     GENSECTORDEST(sECTOR)
  END IF




  posi = 0

  Strsql = "Select * From TIPOREQ where ID = " & wperfil
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     fila = ""


      Celda = ofv.GenCelda("","ttt","","","","","","PERFIL: ","si")
      sInput = rs(1) & "&nbsp;-&nbsp;" & rs(2)

      Celda = celda & ofv.GenCelda("","ttt colspan=2","","90%","","","",sInput,"si")
      Fila = ofv.genrow("","","","","","","",Celda,"si")



      Formulario = Fila

      sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si")



      fila = ""

      Celda = ofv.GenCelda("","ttr","","","","","","SOLAPA","si")
      Celda = celda & ofv.GenCelda("","ttr","","","","","","TITULO","si")
      Celda = celda & ofv.GenCelda("","ttr","","","","","","ACTIVA","si")
      Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")
      Formulario = Fila

dim atab(15)

    for xx = 1 to 15
           atab(xx) = 99
    next

    for xx = 1 to 9

           atab(XX) = xx + 2



    next

    for yy = 1 to 15

        if clng(atab(yy)) = 99 then exit for
           xx = atab(yy)
           fila = ""
           sForm = "form" & posi

           if rs(xx)  then
              clase = "vT style='background-color:#f9f9f9;border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#005500;' "
           else
              clase = "vt"
           end if

           sInput = YY
           Celda = ofv.GenCelda("",clase,"right","1%","","","",sInput,"si")

           sInput = ofv.GenerarInput(rs(xx + 15).name,rs(xx + 15),"text",sForm,"60","dt","")
           Celda = celda & ofv.GenCelda("",clase,"","40%","","","",sInput,"si")



              sInput = ofv.combosino(sForm,rs(xx),rs(xx).name)

           Celda = celda & ofv.GenCelda("",clase,"","5%","","","",sInput,"si")





           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

           sAccion = ""
           if bModificar then
              sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
              sAccion = sAccion & "&tabla=TIPOREQ&volver=" & vuelta
           end if
           Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

           posi = posi + 1



    next

    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")




  celdas = ofv.GenCelda("","ttt colspan=2 ","","","","","","GERENCIA/SECTOR INTERVINIENTE","si")

     sinput = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=REQsectoresBUS.asp?VASP=RP"
     sinput = sinput  & " ><b>" & sImagenI3& "&nbsp;afectar&nbsp;sector&nbsp;</b></a>" 

  celdas = celdas & ofv.GenCelda("","ttt","","10%","","","",sInput,"si")
  filas  = ofv.GenRow("","","","","30","","",celdas,"si")



  celdas = ofv.GenCelda("","ut","","30%","","","","GERENCIA: ","si")



    if RS(33) = 0 then
      celdas = celdas & ofv.GenCelda("","uti COLSPAN=2","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(33)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = celdas & ofv.GenCelda("","uti COLSPAN=2","","","","","",sInput,"si")


    end if


  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","ut","","30%","","","","SECTOR: ","si")

    if RS(34) = 0 then
      celdas = CELDAS & ofv.GenCelda("","uti COLSPAN=2 ","","","","","","SIN AFECTAR","si")

    else

     StrSql = "Select * From ECO where ECO = " & RS(34)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sInput = rsX2(1)
     ELSE
        sInput = "NO DETERMINADO"
     END IF
     celdas = CELDAS & ofv.GenCelda("","uti COLSPAN=2 ","","","","","",sInput,"si")


    end if

    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","ut","","30%","","","","NIVEL MINIMO REQUERIDO: ","si")

    if bModificar then
      strsql = "Select * From NivelConfiden Order By NCcod"
      sInput = ofv.GenerarCombo("Usuarios",rs(35).name,rs(35),strsql,cn2,0,1,"")
    else
      strsql = "Select * From NivelConfiden where NCcod = " & rs(35)
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

     sinput2 = "&nbsp;<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";'  href='javascript:void(0);' "
     sinput2 = sinput2  & " onclick=" & chr(34) & "AbrirCons('correqperfilver.asp'); return false;" & chr(34) & " ><b>" & sImagenI1& "&nbsp;VERIFICAR&nbsp;</b></a>" 


    celdas = CELDAS & ofv.GenCelda("","uti ","","","","","",sInput,"si")
    celdas = CELDAS & ofv.GenCelda("","ut","","","","","",sInput2,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    Celda = ofv.GenCelda("","abc colspan=2","center","100%","","","","&nbsp;","si")
  '  FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    celdas = ofv.GenCelda("","ttt colspan=3 ","","","","","","AUTORIZACION/APROBACION/CIRCUITOS","si")
    filas  = filaS & ofv.GenRow("","","","","30","","",celdas,"si")

 SEGWHERE = " WHERE INI = 1 AND INICIO = 'A' AND ( (OBJTIPO = 'GEN'  )  OR (OBJTIPO = 'REQ')) "

  celdas = ofv.GenCelda("","ut  ","","30%","","","","CIRCUITO AUTORIZACION: ","si")

    if bModificar then
        StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
        sInput = ofv.GenerarCombo("Usuarios",rs(37).name,rs(37),strsql,cn,0,1,"sb0")
    else
        StrSql = "Select * From WKFCKTMASTER WHERE WKFCOD = '" & rs(37) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = CELDAS & ofv.GenCelda("","vti COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","ut  ","","30%","","","","CIRCUITO APROBACION: ","si")

    if bModificar then
        StrSql = "Select * From WKFCKTMASTER " & SEGWHERE & " ORDER BY DESCRIPCION "
        sInput = ofv.GenerarCombo("Usuarios",rs(38).name,rs(38),strsql,cn,0,1,"sb0")
    else
        StrSql = "Select * From WKFCKTMASTER WHERE WKFCOD = '" & rs(38) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = CELDAS & ofv.GenCelda("","vti COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","ut  ","","30%","","","","NIVEL MINIMO REQUERIDO: ","si")

    if bModificar then
      strsql = "Select * From NivelConfiden Order By NCcod"
      sInput = ofv.GenerarCombo("Usuarios",rs(36).name,rs(36),strsql,cn2,0,1,"")
    else
      strsql = "Select * From NivelConfiden where NCcod = " & rs(36)
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = ofv.GenerarInput("",sniv,"readonly","","40","dt","")
    end if

    celdas = CELDAS & ofv.GenCelda("","vti COLSPAN=2 ","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


           sAccion = ""
           if bModificar then
              sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
              sAccion = sAccion & "&tabla=TIPOREQ&volver=" & vuelta
           end if
           Formulario = ofv.genform("Usuarios","",sAccion,"","post","",FilaS,"si")



    FilaS = ofv.GenTabla("","aa ","","100%","","0","","0","0","","",Formulario,"si")

    celdas = ofv.GenCelda("","vt","","","","","",FilaS,"si")
    filas = ofv.GenRow("","","","","10","","",celdas,"si")


    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",FilaS,"si")



  end if



  sAccion = "CORREQPERFIL.asp"
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","",Celda)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

FUNCTION GENSECTORDEST(XSECTOR)





     XSECTORIG = XSECTOR
     XSECTORIGB = XSECTORIG
     XSECTGEREN = 0

        DO
          'Sector Principal
          StrSql = "SELECT ECO_1.* FROM ECO, ECOREL, ECO  ECO_1 where ECO.eco = ECOREL.ECODEPEND"
          StrSql = StrSql & " and ECOREL.ECOMASTER = ECO_1.eco and ECO.eco = " & XSECTORIGB
          set rsX3 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
          If rsX3.EOF Then
             EXIT DO
          Else
              XSECTGEREN = rsX3(0)
              XSECTORIGB = rsX3(0)
              StrSql = "SELECT * FROM CSECOSNJ where ECO = " & XSECTORIGB
              set rsX4 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
              If rsX4.EOF Then
                 EXIT DO
              Else
                 XJERAR = RSX4(3)
                 XJERAR = INSTR(XJERAR,"GEREN")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
                 XJERAR = INSTR(XJERAR,"DIRECT")
                 IF XJERAR > 0 THEN
                    EXIT DO
                 END IF
              End If
          End If
        
        LOOP 

        STRSQL = "UPDATE TIPOREQ SET DESTSECTOR = '" & XSECTORIG & "' ,DESTGERENCIA = '" & XSECTGEREN & "' WHERE ID = " & SESSION("wperfil")
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'        response.write strsql











END FUNCTION

%>