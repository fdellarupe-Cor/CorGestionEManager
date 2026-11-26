<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 '' session("ModMain") = "CL"
  XXX= APLCHECK()
  response.expires=-1441
  Response.AddHeader "PRAGMA", "NO-CACHE"




  aplcheck()

  set cn = ofv.conectar(ofv.strconn2)




  sHTML = ofv.FormHeader( session("FormN"))
  sHTML = sHTML & "<center>"

'' response.write " Hola "

  tt="&nbsp;"


       	   XFNAME   = request.querystring("XFNAME")
      	   XFSRH    = request.querystring("XFSRH")
      	   XFSRD    = request.querystring("XFSRD")



      	   IF XFNAME = "" THEN XFNAME = "PV"
      	   IF XFSRD  = "" THEN XFSRD = 0
      	   IF XFSRH  = "" THEN XFSRH = 0

      	 ' PRIMERA VEZ
      	  IF XFNAME="PV"  THEN
      	     XFSRD = clng(DATE()) - clng(cdate("01-01-1900"))
'''''''''''''XFSRH = XFSRD + 15
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
            sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Desde3",XFSRD,"hidden","","20","dt","")
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
            sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Hasta3",XFSRH,"hidden","","20","dt","")
            Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")
            Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

      ' estados
	      Celda  = ""
	  	'Celda  = ofv.GenCelda("","ut","","","","","","<font>Estados:</font>","si")

	  	'scomK = "3" & chr(9) & "4" & chr(9) & "TODOS"
	  	'scomD = "VIGENTE" & chr(9) & "NO VIGENTE" & chr(9) & "TODOS"
	  	'sInput = ofv.gencomboFX("Subconsulta","ESELESTADOS","3",scomK,scomD,"no")
	  	'Celda  = celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
	Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
	    IF  session("APLMDCLI")="CFA" THEN
	     scomK ="NOL"
	     scomD = "SIN ACCESOS"
	     sInput = ofv.gencomboFX("Subconsulta","ESEL","NOL",scomK,scomD,"no")

	    ELSE

	     scomK = "MAS" & chr(9) & "MEN" & chr(9) & "NOL"
	     scomD = "MAS ACCESOS" & chr(9) & "MENOS ACCESOS" & chr(9) & "SIN ACCESOS"
	     sInput = ofv.gencomboFX("Subconsulta","ESEL","MAS",scomK,scomD,"no")

	    END IF

            Celda  = ""
            Celda  = ofv.GenCelda("","tc","","","","","","<font>Orden:</font>","si")
	    Celda  = celda& ofv.GenCelda("","uti","","","","","",sInput,"si")
	    Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")



	    IF session("APLMDCLI")="BLP" THEN

	    Celda  = ofv.GenCelda("","tc","","","","","","<font>Formato Exportacion :</font>","si")
	    SCOMD= "XLS" &  CHR(9) & "NO"
	    scomK=SCOMD
	    sInput = ofv.gencomboFX("Subconsulta","EXPORT","NO",scomK,scomD,"no")
	    Celda = Celda & ofv.gencelda("","uti","","","","","",sInput,"si")
	    Fila = Fila & ofv.genrow("","","","","25","","",Celda,"si")
	    END IF





      	' BOTON BUSCAR
      	 sTARGET  = ""
         sAccion    = "ConsdocumentoswebScons.asp?Menu=PORREVISION&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=POR REVISION"

      	 Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
      	 sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")

      	 sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
      	 Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
  	     formx    = fila
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


      sAccion = "COESTMODLis.asp?CantRegMostrar=10&CantRegMover=0&tipoC=ACCXMOD"

     'IF session("APLMDCLI")="BLP" THEN
      Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post Onsubmit="&chr(34) &"if(document.Subconsulta.EXPORT.value=='XLS'){this.target='_NEW'; } else{this.target='_self'};" &chr(34),"",Fila,"si")

     'Else

     'Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",Fila,"si")


     'END IF






      	' ****** CODIGO PARA CALENDARIO

      	 	 sInputDesde=ofv.GenerarInput("Desde",cdate(XFSRD+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
           	 sAccionRetornoD="COESTDXMODULO.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
            	 formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")


      	 	 sInputHasta=ofv.GenerarInput("Hasta",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
            	 sAccionRetornoH="COESTDXMODULO.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
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

Function Fechador(wkfecha)



      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
    '  fec = cdate(wkfecha )
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>