<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 '' session("ModMain") = "CL"

  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"

  aplcheck()

  set cn = ofv.conectar(ofv.strconn2)


  sHTML = ofv.FormHeader( session("FormN"))
  sHTML = sHTML & "<center>"

  tt="&nbsp;"

XXXX = NEWEMPRE("NO")

        session("ASP")="SUBCONSULTA.ASP"

        sEMPRE = request.querystring("EMPRE")

         IF sEMPRE = ""  AND SESSION("UN") = "" THEN

           sEMPRE = ATREMP(0, 0)
           SESSION("UN") = sEMPRE & ""
           If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
         elseIF sEMPRE <> "" THEN
          XXXX = NEWEMPRE(sEMPRE)
          SESSION("UN")=sEMPRE
         ELSE
          XXXX = NEWEMPRE(SESSION("UN"))
          sEMPRE=SESSION("UN")
     end if

  'response.write session("UN")
  TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")

shtml = shtml & tablah




     	   XFNAME   = request.querystring("XFNAME")
    	   XFSRH    = request.querystring("XFSRH")
    	   XFSRD    = request.querystring("XFSRD")



    	   IF XFNAME = "" THEN XFNAME = "PV"
    	   IF XFSRD  = "" THEN XFSRD = 0
    	   IF XFSRH  = "" THEN XFSRH = 0

    	 ' PRIMERA VEZ
    	  IF XFNAME="PV"  THEN
    	     XFSRD = clng(DATE()) - clng(cdate("01-01-1900"))
'''''''''''XFSRH = XFSRD + 15
           XFSRH = XFSRD

    	  END IF

    	  IF XFNAME = "FD" THEN
    	  	XFSRD = request.form("DESDE")
    	  	XFSRH = XFSRH

    	  END IF

    	 IF XFNAME = "FH" THEN
    	  	XFSRD = XFSRD
    	  	XFSRH =request.form("HASTA")

  	 END IF


    'Descripcion
    	   'sInput = "<h3>Panel de filtros para la consulta</h3>"
    	   'Celda  = ofv.GenCelda("","ut colspan=2 bgcolor='#f4f7fc' style='border-style:solid;border-width:1;border-color:#dddddd;' ","","","","","",sInput,"si")
    	   'Fila   = ofv.genrow("","","","","30","","",Celda,"si")
    	   Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
    	   Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
    	   'Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

          sImagenE = "<B><i class='las la-calendar'STYLE=font-size:14pt;color:" & Session("CMFPNBOT") &"></i></B>"
         sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,4,2,0,0,0)
         sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
         sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


         sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
         sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
         sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"



         ' CELDAS DE DESDE

  	    if  clng(XFSRD) = 0 then
  	   	   sInput2 = "------------"
  		else
  		   sInput2 = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
  		   swd = datepart("w",sInput2)
  		   sInput2 = Fechador(XFSRD)
  	    END IF




          Celda  = ofv.GenCelda("","tc","","","","","","<font>Desde</font>","si")
          sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",XFSRD,"hidden","","20","dt","")
          Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")
      	  Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


   	' CELDAS DE HASTA
  	    if clng(XFSRH) = 0 then
  	           sInput2 = "------------"
  	        else
  	           sInput2 = cdate(clng(XFSRH) + clng(cdate("01-01-1900")) )
  	           swd = datepart("w",sInput2)
  	           sInput2 = Fechador(XFSRH)
      	   END IF


          Celda  = ofv.GenCelda("","tc","","","","","","<font>Hasta</font>","si")
          sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",XFSRH,"hidden","","20","dt","")
          Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")
          Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


    	' BOTON BUSCAR
    	 sTARGET  = ""
       	 sAccion    = "ConsdocumentoswebScons.asp?Menu=PORREVISION&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=POR REVISION"

    	 Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
    	 sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")

    	 sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
    	 Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
	 formx = fila
       	 formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")


   'IF session("APLMDCLI")="BLP" THEN

	     Celda = ofv.gencelda("","tc","","","","","","Formato Exportacion:&nbsp;&nbsp;","si")
	     SCOMD= "XLS" &  CHR(9) & "NO"
	     scomK=SCOMD
	     sInput = ofv.gencomboFX("Subconsulta","EXPORT","NO",scomK,scomD,"no")
	     Celda = Celda & ofv.gencelda("","uti","","","","","",sInput,"si")
	     Fila = Fila & ofv.genrow("","","","","25","","",Celda,"si")


   'END IF

    sInput = ofv.GenerarInput("","Buscar","submit","","10","bt","")
    Celda  = ofv.GenCelda("","","Center colspan=2","","","","",sInput,"si")
    Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    formx = fila
    formx = formx  & ofv.GenRow("","","","","40","","",Celda,"si")


    sAccion = "COESTDLisSEG.asp?CantRegMostrar=10&CantRegMover=0&tipoC=ACCXESQ"

   ' IF session("APLMDCLI")="BLP" THEN
    Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post Onsubmit="&chr(34) &"if(document.Subconsulta.EXPORT.value=='XLS'){this.target='_NEW'; } else{this.target='_self'};" &chr(34),"",Fila,"si")

   ' Else

   ' Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",Fila,"si")


   ' END IF






    	' ****** CODIGO PARA CALENDARIO

    	 	 sInputDesde=ofv.GenerarInput("Desde",cdate(XFSRD+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
         	 sAccionRetornoD="COESTDSEGUI.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
          	 formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")


    	 	 sInputHasta=ofv.GenerarInput("Hasta",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
          	 sAccionRetornoH="COESTDSEGUI.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
          	 formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")

    	' ****** CODIGO PARA CALENDARIO

    Formulario = Formulario &formRetornoD & formRetornoH





   END IF



   filas2 = "" & filas2 & "" & ofv.GenTabla("","aaa","","80%","","0","","5","3","","",Formulario,"si")

   sHTML = sHTML & filas2



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & "</center>"

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  sHTML = sHTML & "</script>"& GTCCALENDAR("")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)


response.write sHTML

function selecEmpre()

sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

   IF  Session("EMPMUN") = "S" THEN

   Fila    = ""

   SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

   Fila    = ""
   Celda   = ""
   sHref   = ""

   sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
   sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
   Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
   Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
   Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

   Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

   Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





   shtml=shtml & "<script language=javascript>" & chr(13)
   shtml=shtml & "function shwDem(){" & chr(13)
   shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
   shtml=shtml & "}" & chr(13)
   shtml=shtml & "</script>" & chr(13)
 '----------------------------------------------
   sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
   sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
   sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
   TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


   TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
   TABLAX = TablaX & "<tr><td  "
   TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
   TABLAX = TablaX & "<tr><td colspan=3 "
   TABLAX = TablaX & " >&nbsp;</td></tr>"

   FOR IAEMP=1 TO TAEMP

    '''  RESPONSE.WRITE  ATREMP(IAEMP,2)

      IF ATREMP(IAEMP,2) = 1 THEN

       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
       ELSE
          TABLAX = TablaX & "<tr>"
       END IF

       TABLAX = TablaX & "<td width=5 "
       IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
          TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
       ELSE
          TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
       END IF

       sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
       TABLAX = TablaX & "<td width=5 "
       TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

       TABLAX = TablaX & "<td class=vt  >"

       'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
 		sHref = "COESTDSEGUI.asp?EMPRE=" & ATREMP(IAEMP,0)
     '''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
       sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
       TABLAX = TablaX & sLink
       TABLAX = TablaX & "</td>"
       TABLAX = TablaX & "</tr>"
      END IF

   NEXT

   TABLAX = TablaX & OFV.genFtabla()

   TABLA = Tabla & TABLAX

   TABLA = Tabla & "</DIV>"

   sHTML = sHTML & TABLA

  end if

end function

Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>