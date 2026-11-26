<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMoverx = request.querystring("retorno")

  if CantRegMoverx = "" then  CantRegMoverx = request.querystring("Cantregmover")

  nEditar=cdbl("0" &request.querystring("Editar"))

  diasExtendido=request.Form("diasextender")


  ID = request.querystring("ID")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID & "_Editar=" & nEditar

  sHTML = ofv.FormHeader(session("FormN") & " - Propiedades del Mensaje")

  'Encabezados
  Formulario = ""


'	response.write cdate (42973 + clng(cdate("01-01-1900")))





  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  posi = 0
  sImagenI4 = "<B><FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"

  ' Actualizacion de Tablas al Usar la opcion de Editar

      if nEditar=1   and  diasExtendido <> "" then

       msgcopete=request.form ("msgcopete")
       msgtitulo=request.form ("msgtitulo")
       texto=request.form ("texto")

       strsql= "UPDATE novegenpar  "
       strsql= strsql & " set  msgtitulo='" & msgtitulo & "', msgcopete ='"& msgcopete&"', texto ='"& texto & "',fecexp= fecexp+"& +diasExtendido &" , dias =dias +"& + diasExtendido
       strsql=  strsql & " where objid='"& ID &"'"

   '    response.write strsql &"<Br>"

      set rs2= ofv.crearconsultaEx(StrSql,cn,1,parametros)


	strsql= "UPDATE novedades  "
	strsql= strsql & " set  msgtitulo='" & msgtitulo& "', msgcopete ='"& msgcopete &"', mensaje ='"& texto & "',fecexp=fecexp+"& diasExtendido &""
	strsql=  strsql & " where codigo='NGP"& ID &"'"

    '   response.write strsql &"<Br>"
       set rs3= ofv.crearconsultaEx(StrSql,cn,1,parametros)

      end if





  Strsql = "Select * From novegenpar  where objid = " & ID

 ' response.write Strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then




	    sxopc1 ="<option  value=0 > SIN EXTENDER</option>"
		sxopc1 = sxopc1 & "<option  value='1' >1 DIA</option> "
		sxopc1 = sxopc1 & "<option value='2' >2 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='7' >7 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='10' >10 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='15' >15 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='20' >20 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='25' >25 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='30' >30 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='35' >35 DIAS</option> "
	    sxopc1 = sxopc1 & "<option value='40' >40 DIAS</option> "

        sxopc1 = replace(sxopc1,"value=0 "  ,"value=0 selected")


        sxopc3 = "<option  value='1' >1 DIA</option> "
		sxopc3 = sxopc3 & "<option value='2' >2 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='7' >7 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='10' >10 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='15' >15 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='20' >20 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='25' >25 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='30' >30 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='35' >35 DIAS</option> "
		sxopc3 = sxopc3 & "<option value='40' >40 DIAS</option> "



        sxopc3 = replace(sxopc3,"value='" & rs(3) &"'","value='" & rs(3) & "' selected")



        sxopc4 = "<option value=0 >ALTA</option> "
        sxopc4 = sxopc4 & "<option value=1 >NORMAL</option> "
        sxopc4 = sxopc4 & "<option value=2 >BAJA</option> "

        sxopc4 = replace(sxopc4,"value=" & rs(5),"value=" & rs(5) & " selected")

        sxopc5 = "<option value=1 >TODOS</option> "
        sxopc5 = sxopc5 & "<option value=2 >LISTA</option> "


        sxopc5 = replace(sxopc5,"value=" & rs(6),"value=" & rs(6) & " selected")

        VDOPCION = RS(12) & ""

        IF VDOPCION = "" THEN VDOPCION = "0"


        sxopcVD = "<option value=0 >NO</option> "
        sxopcVD = sxopcVD & "<option value=1 >IMAGEN</option> "
        sxopcVD = sxopcVD & "<option value=2 >VIDEO</option> "

        sxopcVD = replace(sxopcVD,"value=" & VDOPCION,"value=" & VDOPCION & " selected")


     fila = ""
     sForm = "formu"

    '  selacc = " onchange=" & chr(34) & "document." &  sForm & ".submit(); return true;" & chr(34)


      Celda =  ofv.GenCelda("","tt  ","","","","","","Codigo: ","si")
      sInput = "<b>NGP" & rs(0) & "</b>"
      Celda = celda & ofv.GenCelda("","cc colspan=5 ","","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



      Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")


 	   XFNAME   = request.querystring("XFNAME")
	   XFSRD    = request.querystring("XFSRD")



	   IF XFNAME = "" THEN XFNAME = "PV"
	   IF XFSRD  = "" THEN XFSRD = 0


	 ' PRIMERA VEZ
	  IF XFNAME="PV"  THEN
	     XFSRD = clng(rs(1))

	  END IF

	  IF XFNAME = "FD" THEN
	  	XFSRD = request.form("DESDE")

	  END IF

'response.write XFNAME


    sAccion = "AdNovedadesFec.ASP?XFNAME=NO&id=" & id & "&CantRegMover=" & CantRegMoverx & "&XFSRD=0" & rs(1)

    sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion

	if  clng(XFSRD) = 0 then
	   	   sInput2 = Fechador(rs(1)+ clng(cdate("01-01-1900")) )

		else
		   sInput2 = cdate(clng(XFSRD) )
		   swd = datepart("w",sInput2)
		   sInput2 = Fechador(XFSRD + clng(cdate("01-01-1900")))

	    END IF

'"frmCalendarioD","Desde"

    if  nEditar = 1 then

    	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

      Celda = ofv.GenCelda("","tt  ","","","","","","Fecha Vigencia: ","si")

    sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0167;</FONT></B>"
     sinputE = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
	  sinputE = sinputE  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
	  sinputE = sinputE  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
	  sinputE = sinputE  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
	  sinputE = sinputE  & "none" & chr(34) & ";' "
	  sinputE = sinputE  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(1).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
	  sinputE = sinputE  & "  TITLE='Modificar fecha de vigencia.'  ><b>&nbsp;" & sImagenI4 & " " & "&nbsp;</b></a>"


    sInput = ofv.GenerarInput(rs(1).name,rs(1).value  + clng(cdate("01-01-1900")) ,"text",sForm,"10","dt","")
    Celda = celda & ofv.GenCelda("","cc  ","","25%","","","",sInput2 & sinputE & sInput,"si")

   ' if  nEditar<>1 then

   ' 	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

   ' End if

   ' Fecha de Expiracion

   if  nEditar=1 then

    Celda = celda & ofv.GenCelda("","tt  ","","","","","","Fecha de Expiracion: ","si")
    sInput = fechador(rs(2)  + clng(cdate("01-01-1900")))
    Celda = celda & ofv.GenCelda("","cc  ","","15%","","","",sInput,"si")

    ' Dias de Vigencia Extender
        ' Celda de Dias de Vigencia
        Celda = celda & ofv.GenCelda("","tt  ","","","","","","Extender Vigencia: ","si")
        sInput = "<select  class=at name=diasextender >" & sxopc1 & "</select>"


    else

    ' Dias de Vigencia
        ' Celda de Dias de Vigencia
        Celda = celda & ofv.GenCelda("","tt  ","","","","","","Dias de Vigencia: ","si")
    sInput = "<select  class=at name=" & rs(3).name  & selacc &  " >" & sxopc3 & "</select>"

    End if








    Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    sInput = ofv.GenerarInput(rs(9).name,"2","hidden",sForm,"1","dt","")

    Fila = fila & ofv.genrow("","","","","","","","<br>" & sInput,"si")



      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","titulo","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(10).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 190) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=2>" & REPLACE(rs(10) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","copete","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(11).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(11) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","Mensaje","si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(4).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 950) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=6>" & REPLACE(rs(4) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Fila = fila & ofv.genrow("","","","","","","","<br>","si")





  Celda =  ofv.GenCelda("","tt  ","","","","","","IMAGENES/VIDEO: ","si")

  if  nEditar <> 1  then

      sInput = "<select class=at name=" & rs(12).name  & selacc &  " >" & sxopcVD & "</select>"
      Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")

     Else

    	 select case rs(12)

                	case "0"  Sinput="NO"

          	        case "1"  Sinput="IMAGEN"

          	        case "2"  Sinput="VIDEO"


      	end Select

        Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")


 end if

' Creacion de Celda Boton Seleccionar Imagenes


   IF (RS(12) & "") = "1" and nEditar <> 1 THEN


      sInput = ofv.GenerarInput("","SELECCIONAR","BUTTON","","10","btSB","document.Botonx.submit();")
      Celda = celda & ofv.GenCelda("","cc colspan=6 ","","","","","",sInput,"si")

       sAccion = "Adnovedadesimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id
       FormuX = ofv.GenForm("Botonx","",sAccion,"","post","","&nbsp;","si")


   else


       Celda = celda &  ofv.GenCelda("","cc colspan=6  ","","","","","","&nbsp;","si")

   end if




  Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



   IF (RS(12) & "") = "2" THEN

      Celda =  ofv.GenCelda("","tt colspan=6 ","center","","","","","URL DEL VIDEO:","si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


      sInput = "<textarea name=" & rs(14).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 490) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
      sInput = sInput & " cols=120 rows=3>" & REPLACE(rs(14) & "","''","'") & "</textarea>"
      Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

   END IF





     ' Creacion de Celda de Importancia

     Celda =  ofv.GenCelda("","tt  ","","","","","","Importancia: ","si")
     if  nEditar <> 1 then

     	sInput = "<select class=at name=" & rs(5).name  & selacc &  " >" & sxopc4 & "</select>"
     	Celda = celda & ofv.GenCelda("","cc","","35%","","","",sInput,"si")

     else

     select case rs(5)

           	case "0"  Sinput="ALTA"

     	        case "1"  Sinput="NORMAL"

     	        case "2"  Sinput="BAJA"

        end Select


      Celda = celda & ofv.GenCelda("","cc  ","","35%","","","",sInput,"si")
     end if





     ' Celda de Tipo de Distribucion

      Celda = celda &  ofv.GenCelda("","tt  ","","","","","","Tipo de Distribucion: ","si")

      if  nEditar <> 1 then

      	sInput = "<select class=at name=" & rs(6).name  & selacc &  " >" & sxopc5 & "</select>"
      	Celda = celda & ofv.GenCelda("","cc colspan=6 ","","35%","","","",sInput,"si")

      else

      	select case rs(6)

      		case "1"  Sinput="TODOS"

	        case "2"  Sinput="LISTA"

        end Select

     	Celda = celda & ofv.GenCelda("","cc  colspan=6 ","","35%","","","",sInput,"si")

      end if

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




      ' Creacion de Celda de Contacto
      Celda =  ofv.GenCelda("","tt  ","","","","","","Contacto: ","si")
      sInput = "<b>" & rs(7) & "</b>"
      Celda = celda & ofv.GenCelda("","cc colspan=6 ","","70%","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then

      sInput = ofv.GenerarInput("","Actualizar","submit","","10","bt","")
      Celda = ofv.GenCelda("","aaa colspan=4","","10%","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      if nEditar <> 1 then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=novegenpar&volver=" & vuelta & "&Editar="&nEditar
      else
      sAccion = "adnovedadesprop.asp?Editar=1"& "&ID=" & ID
      end if


    end if

    tabla=ofv.GenTabla("","","","80% ","","0","","0","1","","",fila,"si")
    Formulario =  ofv.genform(sForm,"",sAccion,"","post","",tabla,"si")



    posi = posi + 1





  end if

STARGET = ""

		sInputDesde="<input type=text name='" & rs(1).name & "' id='Desde' value='" & XFSRD & "' qstyle='visibility:hidden;' > "
		sAccionRetornoD= "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
		sAccionRetornoD = sAccionRetornoD & "&tabla=novegenpar&volver=" & vuelta & "&Editar="&nEditar

      ' 	 formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")

  sHTML = sHTML & "<center><br>" & Formulario & "</center>" & FormuX & formRetornoD






  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  buscar = ""

  sAccion = "Adnovedades.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)




 sHTML = sHTML & GTCCALENDAR("")


  sHTML = sHTML & "</body></html>"





end if
end if
Response.Write sHTML

Function Fechador(wkfecha)


      wkfecha=clng(wkfecha)
      fec = cdate(wkfecha )
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>