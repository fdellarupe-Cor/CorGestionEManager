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

   Celda = 	 ofv.GenCelda("","uti  ","","90%","","","","GENERACION REGIMEN INFORMATIVO" ,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

   sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")






  'Encabezados

    Filax  = ""

    Formulario = ""

    Celda = ""





    Celda = ""
	'TRIMESTRES
	Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","TRIMESTRE","si")


	 SCOMSEP= CHR(9)
	 scomK=""
	 scomD=""
	 FOR I =1 TO 4
	   scomK   = scomK & I   & SCOMSEP
	   XX= GetFecTrim(I,YEAR(DATE()),fecini,fecfin)
	   scomD   = scomD & CSTR(fecini) & "------" & CSTR(fecfin)   & SCOMSEP
	 NEXT



		sInput = ofv.gencomboFX(sForm,"TRIMESTRE","1",scomK,scomD,"no")

	    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


        Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","FECHA","si")

		SINPUT = ofv.GenerarInput("ORFECHA","TODOS","agregar",sForm,"20","dt","")
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


' fecha Carga
        Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","FECHA CARGA","si")


		SINPUT = "<input type=checkbox NAME=FECREGIS  VALUE=FECREGIS>"
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

' fecha DESCUBRIMENTO
        Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","FECHA DESCUBRIMIENTO","si")


		SINPUT = "<input type=checkbox NAME=FECIDENT VALUE=FECIDENT >"
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


' fecha INICIO
        Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","FECHA INICIO","si")


		SINPUT = "<input type=checkbox NAME=FECORIG VALUE=FECORIG >"
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


' fecha CONTABLE
        Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","FECHA CONTABLE","si")

		SINPUT = "<input type=checkbox NAME=FECCONTA VALUE=FECCONTA>"

		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



' fecha FIN
        Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","FECHA FIN","si")

		SINPUT = "<input type=checkbox NAME=FECFINAL VALUE=FECFINAL >"

		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

        Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




     Celda = ""


     SINPUT = "<INPUT TYPE='SUBMIT' VALUE='SELECCIONAR' CLASS='BT'>&nbsp;&nbsp;"
	 Celda = Celda & ofv.gencelda("","ut colspan=3 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

      Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")


      sAccion = "ORGENRUNOREX.asp"

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