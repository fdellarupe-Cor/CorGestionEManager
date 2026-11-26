<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  WSPSESION = "AND"
  SESSION(WSPSESION) = ""

  sHTML = ofv.FormHeader(session("FormN"))

  WKVISION = ""



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

    sForm = "formlst"
    Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BUSQUEDA DE INTERFAZ" ,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")






  'Encabezados

    Filax  = ""

    Formulario = ""

    Celda = ""

    'LOTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","LOTE","si")

    SINPUT = ofv.GenerarInput("ORLOTE","TODOS","agregar",sForm,"10","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")





    Celda = ""
	'FECHA INICIO
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA INICIO","si")
	SINPUT = ofv.GenerarInput("FINICIO","TODOS","agregar",sForm,"70","dt","")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

	Celda = ""
		'FECHA FIN
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA FIN","si")
	SINPUT = ofv.GenerarInput("FFIN","TODOS","agregar",sForm,"70","dt","")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
     'NUMERO CUENTA 1
	 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","NUMERO CUENTA 1","si")


     Strsql = "Select * From ObjCuenta WHERE ESTADO=3 "

      set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			     SCOMSEP = chr(9)
			     scomK   = "TODOS"
			     scomD   = "TODOS"

			     do while not rsx1.eof
			         scomK = scomK & SCOMSEP & rsx1(2)
			          scomD = scomD & SCOMSEP & "<font color=#ff0000>"& rsx1(2) &"</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &rsx1(1)
			         rsx1.movenext

		         loop


		sInput = ofv.gencomboFX(sForm,"ORCUENTA1","TODOS",scomK,scomD,"no")

	    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



	    Celda = ""
		     'NUMERO CUENTA 2
			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","NUMERO CUENTA 2","si")


		     Strsql = "Select * From ObjCuenta WHERE ESTADO=3 "

		      set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			 SCOMSEP = chr(9)
			 scomK   = "TODOS"
			 scomD   = "TODOS"

			 do while not rsx1.eof
				 scomK = scomK & SCOMSEP & rsx1(2)
				 scomD = scomD & SCOMSEP & "<font color=#ff0000>"& rsx1(2) &"</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &rsx1(1)
				 rsx1.movenext

			 loop


			sInput = ofv.gencomboFX(sForm,"ORCUENTA2","TODOS",scomK,scomD,"no")

			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

			Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

	    Celda = ""
		     'NUMERO CUENTA 3
			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","NUMERO CUENTA 3","si")


		     Strsql = "Select * From ObjCuenta WHERE ESTADO=3 "

		      set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			 SCOMSEP = chr(9)
			 scomK   = "TODOS"
			 scomD   = "TODOS"

			 do while not rsx1.eof
				 scomK = scomK & SCOMSEP & rsx1(2)
				  scomD = scomD & SCOMSEP & "<font color=#ff0000>"& rsx1(2) &"</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" &rsx1(1)
				 rsx1.movenext

			 loop


			sInput = ofv.gencomboFX(sForm,"ORCUENTA3","TODOS",scomK,scomD,"no")

			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



        Celda = ""
  		     'estado
  			 Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","ESTADOS","si")




  			 SCOMSEP = chr(9)
  			 scomK   = "TODOS"
  			 scomD   = "TODOS"

  		     scomK = scomK & SCOMSEP & 0
  			 scomD = scomD & SCOMSEP & "NUEVO"

  	        ' scomK = scomK & SCOMSEP & 1
		    ' scomD = scomD & SCOMSEP & "REGISTRADO"

  	         scomK = scomK & SCOMSEP & 2
		     scomD = scomD & SCOMSEP & "PROCESADO"

  	         'scomK = scomK & SCOMSEP & 3
  	        ' scomD = scomD & SCOMSEP & "REVERTIDO"

  	         scomK = scomK & SCOMSEP & 4

		     scomD = scomD & SCOMSEP & "ELIMINADO"
  	         scomK = scomK & SCOMSEP & 5
		     scomD = scomD & SCOMSEP & "SELECCIONADO"


  			sInput = ofv.gencomboFX(sForm,"ORESTADO","TODOS",scomK,scomD,"no")

  			Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

  	        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")










     Celda = ""


     SINPUT = "<INPUT TYPE='SUBMIT' VALUE='SELECCIONAR' CLASS='BT'>&nbsp;&nbsp;"
	 Celda = Celda & ofv.gencelda("","ut colspan=3 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

      Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")


      sAccion = "ORINTERFASESDETEX.asp"

      Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")

       sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa STYLE='border-style:SOLID;border-width:1;border-color:#888888;' ","","70% ","","0","","0","0","","",Formulario,"si") & "</center>"

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""



  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>