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

  sHTML = ofv.MENUHeader(""," ONLOAD=" & chr(34) & "document.title='COR OnTime - Circuito';" & chr(34) )



    celdas = ""
    filas  = ""




  'Encabezados

  celdas = ""
  celdas = celdas & ofv.GenCelda("","ttr colspan=2 ","CENTER","","","","","CIRCUITO","si")
  celdas = celdas & ofv.GenCelda("","ttr  ","CENTER","","","","","INICIANTE","si")

  filas  = filas & ofv.GenRow("","","","","25","","",celdas,"si")


  StrSql = "Select * From SEGPASONV WHERE SEGCKTPSC ='" & PSC & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

  StrSql = "Select * From SEGCIRCUITO WHERE SEGCKT = '" & RS(1) & "' "
  set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rsX.eof then
  posi = 0

     WKFOBJTIPO  = RSX(4)
     WKFOBJCODE  = RSX(5)

     sc = "ut"
     celdas = ""


       sInput = rsX(5)

       IF WKFOBJTIPO = "CLI" THEN
          StrSql = "Select NOMBRE From CLIENTE WHERE CODIGO = '" & rsX(5) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rs2.eof then
	      sInput = "<B>" & RS2(0) & "</B>"
          END IF
       ELSEIF WKFOBJTIPO = "REQ" THEN
         StrSql = "Select TITULO From REQUERIMIENTO where codigo = '" & rsX(5) & "'  "
          set rs2 = ofv.crearconsultaEx(StrSql,cn4,1,parametros)
          if not rs2.eof then
	      sInput = "<B>" & RS2(0) & "</B>"
          END IF
          
       ELSE
	      sInput = " "
         
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





    filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")



    posi = posi + 1

  END IF
 ' sHTML = sHTML  & ofv.GenTabla("","AAA  ","","98%","","0","","-1","-1","","",filas,"si")

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
'  celdas =  ofv.GenCelda("","uT colspan=2 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","center","","","","","&nbsp;","si")
'  filas  =  ofv.GenRow("","","","","","","",celdas,"si")

  PASOANT = " "
  IPASO   = 0

  filas  =  ""
  posi = 0


    sform = "WKF" & rs(0)


       sc = "vt"

       celdas = ofv.GenCelda("","VT BGCOLOR=#EFEFEF ","","30%","","","","PASO","si")




       sInput = "<b>" & sImagenp & "</b> " & rs(18)
       celdas = CELDAS & ofv.GenCelda("","VT BGCOLOR=#EFEFEF  ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")



       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","DESTINO","si")


       sInput = GETRESPO(RS(11),RS(12))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","RESPONSABLE","si")

       sInput = "SIN ASIGNAR"

       IF RS(13) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

          StrSql = "Select * From USUARIOS WHERE ID = '" & RS(13) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not rs2.eof THEN
              sInput = "<B>" & rs2(1) & "</B>"
          END IF
       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","INICIO","si")


       sInput = "<b>" & sImagenFI & "</b> " & GTFechadorsh(rs(20)) & "&nbsp;-&nbsp;" & GETGLBHMSW(rs(21))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","FINALIZACION","si")

       sInput = "<b>" & sImagenFF & "</b> "

       IF RS(24) = "NO" THEN
          sInput = sInput & "<B><FONT COLOR=#F7AA66 SIZE=1>PENDIENTE</FONT></B>"
       ELSE
          sInput = sInput & GTFechadorsh(rs(22)) & "&nbsp;-&nbsp;" & GETGLBHMSW(rs(23))
       END IF

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       IF RS(24) <> "NO" THEN

          celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","TIEMPO","si")

          sInput = rs(22) - RS(20)

          IF SINPUT = 0 THEN 

             sInput = GETGLBHMS(rs(23) - RS(21))

          ELSE

             sInput = sInput & "&nbsp;d."
   
          END IF 



          celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       END IF




       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","ESTADO","si")


       sInput = sImagenIN

       IF RS(24) = "NO" THEN
          sInput = sImagenK & "&nbsp;EN PROCESO"
       ELSE

             PSCMODEL = rs(4)
             PSCMOT   = "A"
             PSCNODO  = rs(17)


 	    StrSql = "Select * From WKFPASOSPRTYNV WHERE  "
 	    StrSql = StrSql & " WKFPASCOD = '" & PSCMODEL & "' AND ITEMCOD = '" & PSCNODO & "'"
  	    set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	    if not rsXS.eof then

      	                     XPASCOD    = RSXS(7)
    	                     XPASACC    = RSXS(9)
      	                     XPASEST    = RSXS(10)
       	                     PSCMOT     = RSXS(4)
       	                     PSCTIPB    = RSXS(5)
        	             PSCTIPD    = RSXS(5)

         	             SOPCDESC = RSXS(4)

            END IF


            IF PSCNODO = "INI" OR PSCNODO = "FIN" OR PSCNODO = "N01" THEN
               sInput = sImagenA & "&nbsp;" & SOPCDESC
            ELSE
               sInput = sImagenR & "&nbsp;" & SOPCDESC
            END IF  	




       END IF

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

      Strsql = "SELECT * FROM SEGPASONV WHERE SEGPASOSIG = '" & rs(2) & "' "
      set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxT.eof then

         celdas = ofv.GenCelda("","aaa colspan=2 STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","","","","","","&nbsp;","si")
         filas = filas & ofv.GenRow("","","","","","","",celdas,"si")



             PSCMODEL = rsxT(4)
             PSCMOT   = "A"
             PSCNODO  = rsxT(17)


 	    StrSql = "Select * From WKFPASOSPRTYNV WHERE  "
 	    StrSql = StrSql & " WKFPASCOD = '" & PSCMODEL & "' AND ITEMCOD = '" & PSCNODO & "'"
  	    set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	    if not rsXS.eof then

      	                     XPASCOD    = RSXS(7)
    	                     XPASACC    = RSXS(9)
      	                     XPASEST    = RSXS(10)
       	                     PSCMOT     = RSXS(4)
       	                     PSCTIPB    = RSXS(5)
        	             PSCTIPD    = RSXS(5)

         	             SOPCDESC = RSXS(4)

            END IF


            IF PSCNODO = "INI" OR PSCNODO = "FIN" OR PSCNODO = "N01" THEN
               sInput2 = sImagenA & "&nbsp;" & SOPCDESC
            ELSE
               sInput2 = sImagenR & "&nbsp;" & SOPCDESC
            END IF  	



       celdas = ""


       sInput = "<b>" & sImagenp & "</b> " & rsxT(18)

       xlink = "<A CLASS=CC href='SEGCKTESTPSCNV.asp?PSC=" & rsxT(2) & "'"
       xlink = xlink & " ><U><b>" & sImagenFA & "</b> " & " VIENE DE</U> " & SINPUT & "</A>"

       SINPUT = xlink


       celdas = CELDAS & ofv.GenCelda("","VT style='background-color:#F0F0F0;' COLSPAN=2  ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","ORIGEN","si")


       sInput = GETRESPO(rsxT(11),rsxT(12))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","RESPONSABLE","si")

       sInput = "SIN ASIGNAR"

       IF rsxT(13) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

          StrSql = "Select * From USUARIOS WHERE ID = '" & rsxT(13) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not rs2.eof THEN
              sInput = "<B>" & rs2(1) & "</B>"
          END IF
       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","ACCION","si")


       sInput = SINPUT2
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    END IF



      Strsql = "SELECT * FROM SEGPASONV WHERE SEGCKTPSC = '" & rs(16) & "' "
      set rsxT = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxT.eof then

         celdas = ofv.GenCelda("","aaa colspan=2 STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","","","","","","&nbsp;","si")
         filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


             PSCMODEL = rsxT(4)
             PSCMOT   = "A"
             PSCNODO  = rsxT(17)


 	    StrSql = "Select * From WKFPASOSPRTYNV WHERE  "
 	    StrSql = StrSql & " WKFPASCOD = '" & PSCMODEL & "' AND ITEMCOD = '" & PSCNODO & "'"
  	    set rsXS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  	    if not rsXS.eof then

      	                     XPASCOD    = RSXS(7)
    	                     XPASACC    = RSXS(9)
      	                     XPASEST    = RSXS(10)
       	                     PSCMOT     = RSXS(4)
       	                     PSCTIPB    = RSXS(5)
        	             PSCTIPD    = RSXS(5)

         	             SOPCDESC = RSXS(4)

            END IF


            IF PSCNODO = "INI" OR PSCNODO = "FIN" OR PSCNODO = "N01" THEN
               sInput2 = sImagenA & "&nbsp;" & SOPCDESC
            ELSE
               sInput2 = sImagenR & "&nbsp;" & SOPCDESC
            END IF  	



       celdas = ""


       sInput = "<b>" & sImagenp & "</b> " & rsxT(18)

       xlink = "<A CLASS=CC href='SEGCKTESTPSCNV.asp?PSC=" & rsxT(2) & "'"
       xlink = xlink & " ><U><b>" & sImagenFS & "</b> " & " SIGUIENTE</U> " & SINPUT & "</A>"

       SINPUT = xlink


       celdas = CELDAS & ofv.GenCelda("","VT style='background-color:#F0F0F0;' COLSPAN=2 ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","DESTINO","si")


       sInput = GETRESPO(rsxT(11),rsxT(12))
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","RESPONSABLE","si")

       sInput = "SIN ASIGNAR"

       IF rsxT(13) <> "NO" THEN

          set cn2 = ofv.conectar(ofv.strconn0)

          StrSql = "Select * From USUARIOS WHERE ID = '" & rsxT(13) & "' "
          set rs2 = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
          if not rs2.eof THEN
              sInput = "<B>" & rs2(1) & "</B>"
          END IF
       END IF

       sInput = "<b>" & sImageni & "</b> " & sInput

       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

       celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;' ","","30%","","","","ACCION","si")


       sInput = SINPUT2
       celdas = CELDAS & ofv.GenCelda("","VT   ","","","","","",sInput,"si")
       filas = filas & ofv.GenRow("","","","","","","",celdas,"si")


    END IF



    celdas = ofv.GenCelda("","UtI style='background-color:#AAAAAA;color:#ffffff;'  colspan=2 valign=middle STYLE='BORDER-TOP-STYLE:ridge;BORDER-TOP-WIDTH:2;border-TOP-color:#6699CC;' ","CENTER","","","","","OBSERVACIONES","si")
    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

    

		    nCol = 80
		    t3b = "<font class=at><textarea style='BORDER-STYLE:solid;BORDER-WIDTH:1;border-color:#f7aa22;background-color:transparent;color:#000080;' "
		    t3b = t3b & "name=OBS cols=" & nCol & " rows=8 "
                    t3b = t3b & " readonly >"
		    t3b = t3b & rs(19) & "</textarea></font>"


	  celdas = ofv.GenCelda("","vt COLSPAN=2   ","center","","","","",t3b,"si")
	  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




  END IF


    celdas = ofv.GenCelda("","aaa colspan=2 style='border-top-style:solid;border-top-width:2;border-top-color:#dddddd;' ","","","","","","&nbsp;","si")
'    filas = filas & ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","aaa  ","","95%","","0","","1","1","","",filas,"si") & "<BR>"







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