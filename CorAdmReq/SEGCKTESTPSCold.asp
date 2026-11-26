<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0




  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn4 = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


 PSC = request.querystring("PSC")

  sHTML = ofv.MENUHeader(""," ONLOAD=" & chr(34) & "document.title='COR OnTime - Consulta de Circuitos';" & chr(34) )



    celdas = ""
    filas  = ""




  'Encabezados

  celdas = ofv.GenCelda("","TTR ","CENTER","","","","","REQUERIMIENTO","si")
  celdas = celdas & ofv.GenCelda("","TTR ","CENTER","","","","","CIRCUITO","si")
  celdas = celdas & ofv.GenCelda("","TTR ","CENTER","","","","","INICIANTE","si")

  filas  = filas & ofv.GenRow("","","","","22","","",celdas,"si")



  StrSql = "Select * From SEGPASO WHERE SEGCKTPSC ='" & PSC & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

  StrSql = "Select * From SEGCIRCUITO WHERE SEGCKT = '" & RS(1) & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsX.eof then
  posi = 0

     WKFOBJTIPO  = RSX(4)
     WKFOBJCODE  = RSX(5)

     sc = "Ut"
     celdas = ""


       sInput = rsX(5)

       IF WKFOBJTIPO = "CLI" THEN
          StrSql = "Select NOMBRE From CLIENTE WHERE CODIGO = '" & rsX(5) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       ELSE
          StrSql = "Select TITULO From REQUERIMIENTO where codigo = '" & rsX(5) & "'  "
          set rs2 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
       END IF



       if not rs2.eof then
          sInput = "<B>" & RS2(0) & "</B>"
       END IF

       SINPUT = REPLACE(SINPUT," ","&nbsp;")   

       celdas = celdas & ofv.GenCelda("",sc,"","15%","","","",sInput,"si")



       sInput = rsX(8)
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")

 set cn2 = ofv.conectar(ofv.strconn0)

  StrSql = "Select * From USUARIOS WHERE ID = '" & RSX(7) & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs2.eof THEN

       sInput = rs2(1)
  END IF
       celdas = celdas & ofv.GenCelda("",sc,"","50%","","","",sInput,"si")





    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



    posi = posi + 1

  END IF
  sHTML = sHTML  & ofv.GenTabla("","  ","","98%","","0","","0","0","","",filas,"si")

'''' PASOS

    sImagenO  = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenI  = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=3>&#0128;</FONT></B>"
    sImagenE  = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC  = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=3>R</FONT></B>"
    sImagenP  = "<B><FONT FACE=WebDINGS COLOR=#336699 SIZE=3>2</FONT></B>"
    sImagenRU = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    sImagenRG = "<B><FONT FACE=WEBDINGS COLOR=#00aa00 SIZE=3>&#0148;</FONT></B>"
    sImagenRS = "<B><FONT FACE=WEBDINGS COLOR=#0000aa SIZE=3>G</FONT></B>"
    sImagenFI = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>5</FONT></B>"
    sImagenFF = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>6</FONT></B>"
    sImagenFA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>5</FONT></B>"
    sImagenFS = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>6</FONT></B>"

    sImagenA  = "<FONT FACE=WINGDINGS COLOR=#005500 SIZE=3>&#0252;</FONT>"
    sImagenR  = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=4>&#0251;</FONT>"
    sImagenK  = "<FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=3>&#0184;</FONT>"
    sImagenIN = "<FONT FACE=WINGDINGS COLOR=#336699 SIZE=3>&#0236;</FONT>"


  'Encabezados

  filas  =  ""

  PASOANT = " "
  IPASO   = 0


  posi = 0


    sform = "WKF" & rs(0)


       sc = "Ut"

       celdas = ofv.GenCelda("","UT ","","30%","","","","PASO","si")




       sInput = "<b>" & sImagenp & "</b> " & rs(10)
       celdas = CELDAS & ofv.GenCelda("","VT BGCOLOR=#EFEFEF  ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")



       celdas = ofv.GenCelda("","VT ","","30%","","","","DESTINO","si")


       sInput = GETRESPO(RS(7),RS(8))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","VT ","","30%","","","","RESPONSABLE","si")

       sInput = "SIN ASIGNAR"

       IF RS(9) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

          StrSql = "Select * From USUARIOS WHERE ID = '" & RS(9) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not rs2.eof THEN
              sInput = "<B>" & rs2(1) & "</B>"
          END IF
       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","VT ","","30%","","","","INICIO","si")


       sInput = "<b>" & sImagenFI & "</b> " & rs(4)
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","VT ","","30%","","","","FINALIZACION","si")

       sInput = "<b>" & sImagenFF & "</b> "

       IF RS(5) = "NO" THEN
          sInput = sInput & "<B><FONT COLOR=#F7AA66 SIZE=1>PENDIENTE</FONT></B>"
       ELSE
          sInput = sInput & rs(5)
       END IF

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","VT ","","30%","","","","ESTADO","si")


       sInput = sImagenIN

       IF RS(5) = "NO" THEN
          sInput = sImagenK & "&nbsp;EN PROCESO"
       ELSE

	  Strsql = "SELECT * FROM SEGPASO WHERE SEGCKTPSC = '" & rs(6) & "' "
	  set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxT.eof then

             PSCMODEL = rsxT(3)
             PSCMOT   = "A"

          END IF

	  Strsql = "SELECT RCHPSC FROM WKFPASOS WHERE WKFPASCOD = '" & rs(3) & "' "
	  set rsxu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxU.eof then

             IF RSXU(0) = PSCMODEL THEN PSCMOT   = "R"

          END IF 

          SELECT CASE PSCMOT
                 CASE "A"
                      sInput = sImagenA & "&nbsp;APROBADO"
                 CASE "R"
                      sInput = sImagenR & "&nbsp;RECHAZADO"
          END SELECT 


       END IF

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

      Strsql = "SELECT * FROM SEGPASO WHERE OKRCH = '" & rs(2) & "' "
      set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxT.eof then

         celdas = ofv.GenCelda("","aaa colspan=2 STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","","","","","","&nbsp;","si")
         filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

           PSCMODEL = rsxT(3)
           PSCMOT   = "A"


	  Strsql = "SELECT RCHPSC FROM WKFPASOS WHERE WKFPASCOD = '" & rsxT(3) & "' "
	  set rsxu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxU.eof then

             IF RSXU(0) = RS(3) THEN PSCMOT   = "R"

          END IF 

          SELECT CASE PSCMOT
                 CASE "A"
                      sInput2 = sImagenA & "&nbsp;POR&nbsp;APROBACION"
                 CASE "R"
                      sInput2 = sImagenR & "&nbsp;POR&nbsp;RECHAZO"
          END SELECT 



       celdas = ""


       sInput = "<b>" & sImagenp & "</b> " & rsxT(10)

       xlink = "<A CLASS=CC href='SEGCKTESTPSC.asp?PSC=" & rsxT(2) & "'"
       xlink = xlink & " ><U><b>" & sImagenFA & "</b> " & " VIENE DE</U> " & SINPUT & "</A>"

       SINPUT = xlink


       celdas = CELDAS & ofv.GenCelda("","VT COLSPAN=2  ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


       celdas = ofv.GenCelda("","VT ","","30%","","","","ORIGEN","si")


       sInput = GETRESPO(rsxT(7),rsxT(8))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","VT ","","30%","","","","RESPONSABLE","si")

       sInput = "SIN ASIGNAR"

       IF rsxT(9) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

          StrSql = "Select * From USUARIOS WHERE ID = '" & rsxT(9) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not rs2.eof THEN
              sInput = "<B>" & rs2(1) & "</B>"
          END IF
       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","VT ","","30%","","","","ACCION","si")


       sInput = SINPUT2
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    END IF

      Strsql = "SELECT * FROM SEGPASO WHERE SEGCKTPSC = '" & rs(6) & "' "
      set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxT.eof then

         celdas = ofv.GenCelda("","aaa colspan=2 STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","","","","","","&nbsp;","si")
         filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

           PSCMODEL = rsxT(3)
           PSCMOT   = "A"


	  Strsql = "SELECT RCHPSC FROM WKFPASOS WHERE WKFPASCOD = '" & rs(3) & "' "
	  set rsxu = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	  if NOT rsxU.eof then

             IF RSXU(0) = PSCMODEL THEN PSCMOT   = "R"

          END IF 

          SELECT CASE PSCMOT
                 CASE "A"
                      sInput2 = sImagenA & "&nbsp;POR&nbsp;APROBACION"
                 CASE "R"
                      sInput2 = sImagenR & "&nbsp;POR&nbsp;RECHAZO"
          END SELECT 



       celdas = ""


       sInput = "<b>" & sImagenp & "</b> " & rsxT(10)

       xlink = "<A CLASS=CC href='SEGCKTESTPSC.asp?PSC=" & rsxT(2) & "'"
       xlink = xlink & " ><U><b>" & sImagenFS & "</b> " & " SIGUIENTE</U> " & SINPUT & "</A>"

       SINPUT = xlink


       celdas = CELDAS & ofv.GenCelda("","VT  COLSPAN=2 ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


       celdas = ofv.GenCelda("","VT ","","30%","","","","DESTINO","si")


       sInput = GETRESPO(rsxT(7),rsxT(8))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","VT ","","30%","","","","RESPONSABLE","si")

       sInput = "SIN ASIGNAR"

       IF rsxT(9) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

          StrSql = "Select * From USUARIOS WHERE ID = '" & rsxT(9) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not rs2.eof THEN
              sInput = "<B>" & rs2(1) & "</B>"
          END IF
       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","VT ","","30%","","","","ACCION","si")


       sInput = SINPUT2
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    END IF



    celdas = ofv.GenCelda("","VT colspan=2 valign=middle STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","CENTER","","","","","OBSERVACIONES","si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

    

		    nCol = 80
		    t3b = "<font class=at><textarea style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#f7aa22;background-color:transparent;color:#000080;' "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=8 "
                    t3b = t3b & " readonly >"
		    t3b = t3b & rs(11) & "</textarea></font>"


	  celdas = ofv.GenCelda("","vt COLSPAN=2   ","center","","","","",t3b,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




  END IF


    celdas = ofv.GenCelda("","aaa colspan=2 style='border-top-style:solid;border-top-width:2;border-top-color:#dddddd;' ","","","","","","&nbsp;","si")
'    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","AAA  style='border-style:solid;border-width:1;border-color:#dddddd;'  ","","100%","","0","","-2","-1","","",filas,"si") & "<BR>"







'''' FIN PASOS



  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML


FUNCTION GETRESPO(RESPTIP,RESPCOD)

  set cn1 = ofv.conectar(ofv.strconn5)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  sInput = "SIN ASIGNAR"



    XRESPO = RESPCOD & ""

    if XRESPO <> "" then

       SELECT CASE RESPTIP
  
         CASE "USR"
           RESPTIP = "<b>" & sImagenRU & "</b> "
           strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

         CASE "GRP"
           RESPTIP = "<b>" & sImagenRG & "</b> "
           strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
           set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

         CASE "SEC"
           RESPTIP = "<b>" & sImagenRS & "</b> "
           strsql = "Select * From ECO where ECO = " & XRESPO
           set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
           if not rsniv.eof then sInput = rsniv(1)
           call ofv.cerrarconsulta(rsniv)

       END SELECT

    END IF




  call ofv.cerrarconn(cn1)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)

  GETRESPO = "<B>" & RESPTIP & " " & sInput & "</B>"

END FUNCTION


%>