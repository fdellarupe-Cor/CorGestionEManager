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

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE BUSQUEDA DE INDICADORES " ,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")






  'Encabezados

    Filax  = ""

    Formulario = ""

    Celda = ""

     'CODIGO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","CODIGO","si")

    SINPUT = ofv.GenerarInput("ORCODIGO","TODOS","agregar",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""

     'NOMBRE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","NOMBRE","si")


    SINPUT = ofv.GenerarInput("ORTEXTO","TODOS","agregar",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""

     'ESTADO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","ESTADO ","si")
    scomK = "1" & chr(9) & "2" &  chr(9) & "3" & chr(9)  & "VENC" & chr(9) & "TODOS"
    scomD = "NUEVO" & chr(9) & "INICIADO" & chr(9) & "NO VIGENTES" & chr(9) & "VENCIDOS" & chr(9) & "TODOS"


    sInput = ofv.gencomboFX(sForm,"ORESTADO","TODOS",scomK,scomD,"no")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""

     'CLASIFICACION
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CLASIFICACION","si")

    Strsql = "Select * From ORCLASIFICAINDI ORDER BY DESCRI "

   	    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	     SCOMSEP = chr(9)
	     scomK   = "TODOS"
	     scomD   = "TODOS"

	     do while not rsx1.eof
	         scomK = scomK & SCOMSEP & rsx1(0)
	         scomD = scomD & SCOMSEP & rsx1(1)
	         rsx1.movenext

         loop

    sInput = ofv.gencomboFX(sForm,"ORCLASIFICAINDI","TODOS",scomK,scomD,"no")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

      'umbrales
	 Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","UMBRALES","si")

	    Strsql = "Select * From ORUMBRALES WHERE IDINDICADOR=0 ORDER BY DESCRIPCION "

	   	    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		     SCOMSEP = chr(9)
		     scomK   = "TODOS"
		     scomD   = "TODOS"

		     do while not rsx1.eof
		         scomK = scomK & SCOMSEP & rsx1(1)
		         scomD = scomD & SCOMSEP & rsx1(3)
		         rsx1.movenext

	         loop

            scomK = scomK & SCOMSEP & "FUERARANGOS"
            scomD = scomD & SCOMSEP & "FUERA DE RANGOS"
	sInput = ofv.gencomboFX(sForm,"ORUMBRAL","TODOS",scomK,scomD,"no")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")





     'fRecuencia
		 Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FRECUENCIA","si")

		    Strsql = "Select * From orfrecuencia ORDER BY DESCRI "

		   	    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			     SCOMSEP = chr(9)
			     scomK   = "TODOS"
			     scomD   = "TODOS"

			     do while not rsx1.eof
			         scomK = scomK & SCOMSEP & rsx1(0)
			         scomD = scomD & SCOMSEP & rsx1(1)
			         rsx1.movenext

		         loop
                 'scomK = scomK & SCOMSEP & "FUERARANGOS"
			     'scomD = scomD & SCOMSEP & "FUERA DE RANGOS"
		sInput = ofv.gencomboFX(sForm,"ORFRECUENCIA","TODOS",scomK,scomD,"no")

	    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""

     'VENCIDOS
   ' Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","VENCIDOS ","si")

  '  SINPUT = ofv.GenerarInput("ORVENCIDOS","TODOS","agregar",sForm,"70","dt","")
  '  Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

  '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
     CELDA=""

     SINPUT = "<INPUT TYPE='SUBMIT' VALUE='SELECCIONAR' CLASS='BT'>&nbsp;&nbsp;"
	 Celda = Celda & ofv.gencelda("","ut colspan=3 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

      Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")


      sAccion = "orMATRIXINDICADORES.asp"

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