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
  opcion = request.querystring("OPCION")

  if CantRegMoverx = "" then  CantRegMoverx = request.querystring("Cantregmover")

  nEditar=cdbl("0" &request.querystring("Editar"))

  diasExtendido=request.Form("diasextender")


  ID = request.querystring("ID")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ID=" & ID & "_Editar=" & nEditar

  sHTML = ofv.FormHeader(session("FormN") & " - Propiedades del Mensaje")

  'Encabezados
  Formulario = ""


OPC=request.querystring("OPC")

'response.write request.querystring
'response.write request.querystring & "<BR>" & request.form



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0167;</FONT></B>"


  set cn = ofv.conectar(ofv.strconn2)






  posi = 0


  ' Actualizacion de Tablas al Usar la opcion de Editar


  if nEditar=1 AND REQUEST.FORM("FECVIG") <> "" THEN

   strsql= "UPDATE novegenpar  "
         strsql= strsql & " set FECVIG='" & REQUEST.FORM("FECVIG")  & "'"
         strsql=  strsql & " where objid='" & ID & "'"
        ' RESPONSE.WRITE STRSQL
         set rs2= ofv.crearconsultaEx(StrSql,cn,1,parametros)
         strsql= "UPDATE novedades  "
         strsql= strsql & " set FECinicio='" & REQUEST.FORM("FECVIG")  & "'"
		  strsql=  strsql & " where codigo='NGP" & ID & "'"
		         ' RESPONSE.WRITE STRSQL
         set rs2= ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF


      if nEditar=1   and  diasExtendido <> "" then

       msgcopete=request.form ("msgcopete")
       msgtitulo=request.form ("msgtitulo")
       texto=request.form ("texto")
       texto = trim(replace(texto,"<br>",""))
       listaDist=request.form("tipodist")
       msgimg=request.form("msgimagen")
       msgurldsc=request.form ("msgurldsc")
       importancia=request.form ("importancia")




       strsql= "UPDATE novegenpar  "
       strsql= strsql & " set importancia=" & importancia & ",msgurldsc='" & msgurldsc & "', msgimagen='" & msgimg & "', tipodist='" & listaDist & "', msgtitulo='" & msgtitulo & "', msgcopete ='"& msgcopete&"', texto ='"& texto & "',fecexp= fecexp+"& +diasExtendido &" , dias =dias +"& + diasExtendido
       strsql=  strsql & " where objid='" & ID & "'"

       'response.write strsql &"<Br>"

      set rs2= ofv.crearconsultaEx(StrSql,cn,1,parametros)


	strsql= "UPDATE novedades  "
	strsql= strsql & " set msgurldsc='" & msgurldsc & "', msgimagen='" & msgimg & "', msgtitulo='" & msgtitulo& "', msgcopete ='"& msgcopete &"', mensaje ='"& texto & "',fecexp=fecexp+"& diasExtendido &""
	strsql=  strsql & " where codigo='NGP"& ID &"'"

       'response.write strsql &"<Br>"
       set rs3= ofv.crearconsultaEx(StrSql,cn,1,parametros)

        call ofv.cerrarconsulta(rs3)


      if Session("EMPMUN") = "S" then
        strsql ="select secuen,tipo,directorio,imagen from NOVEGENPARIMG where  novegenpar='" & ID & "'"

        set rs3= ofv.crearconsultaEx(StrSql,cn,1,parametros)

        if not rs3.eof then

         sql="update novedadesimg set secuen='" & rs3(0) & "' , tipo='" & rs3(1) & "' ,directorio='" & rs3(2)  & "' , imagen='" &  rs3(3) & "'  where novedad='NGP" & ID  & "'"
      'RESPONSE.WRITE SQL
          CALL ofv.crearconsultaEx(SQL,cn,1,parametros)
         end if
          call ofv.cerrarconsulta(rs3)
      end if
      end if





  Strsql = "Select * From novegenpar  where objid = " & ID



  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


   select case Session("EMPMUN")

   case "N"
        sql = "select contacto, enviado from novegenpar where objid=" & rs(0)
		    set rt = ofv.crearconsultaex(sql,cn,1,parametros)
		    puedeEdit=""
		    if rt(0) = session("Usuario") then
		    puedeEdit=true
		    else
		    puedeEdit=false
		end if

  case "S"
          sql = "select codemp, enviado from novegenpar where objid=" & rs(0)
	  	  set rt = ofv.crearconsultaex(sql,cn,1,parametros)
	       puedeEdit=""
	  	   if rt(0) & ""  = session("CLIEMP") & ""  then
	  		 puedeEdit=true
	  		else
	  		 puedeEdit=false
		   end if
 end select


	    sxopc1 ="<option  value=0 > SIN EXTENDER</option>"
		sxopc1 = sxopc1 & "<option  value='1' >1 DIA</option> "
		sxopc1 = sxopc1 & "<option value='2' >2 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='3' >3 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='4' >4 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='5' >5 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='6' >6 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='7' >7 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='8' >8 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='9' >9 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='10' >10 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='11' >11 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='12' >12 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='13' >13 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='14' >14 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='15' >15 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='16' >16 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='17' >17 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='18' >18 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='19' >19 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='20' >20 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='21' >21 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='22' >22 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='23' >23 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='24' >24 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='25' >25 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='26' >26 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='27' >27 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='28' >28 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='29' >29 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='30' >30 DIAS</option> "
		sxopc1 = sxopc1 & "<option value='31' >31 DIAS</option> "


        sxopc1 = replace(sxopc1,"value=0 "  ,"value=0 selected")
		'sxopc1 = replace(sxopc1,"value='" & rs(3) &"'","value='" & rs(3) & "' selected")

        'sxopc3 = "<option  value='1' >1 DIA</option> "
		'sxopc3 = sxopc3 & "<option value='2' >2 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='7' >7 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='10' >10 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='15' >15 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='20' >20 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='25' >25 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='30' >30 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='35' >35 DIAS</option> "
		'sxopc3 = sxopc3 & "<option value='40' >40 DIAS</option> "
		sxopc3 ="<option  value=0 > SIN EXTENDER</option>"
				sxopc3 = sxopc3 & "<option  value='1' >1 DIA</option> "
				sxopc3 = sxopc3 & "<option value='2' >2 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='3' >3 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='4' >4 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='5' >5 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='6' >6 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='7' >7 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='8' >8 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='9' >9 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='10' >10 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='11' >11 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='12' >12 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='13' >13 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='14' >14 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='15' >15 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='16' >16 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='17' >17 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='18' >18 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='19' >19 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='20' >20 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='21' >21 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='22' >22 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='23' >23 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='24' >24 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='25' >25 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='26' >26 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='27' >27 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='28' >28 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='29' >29 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='30' >30 DIAS</option> "
				sxopc3 = sxopc3 & "<option value='31' >31 DIAS</option> "



        'sxopc3 = replace(sxopc3,"value='" & rs(3) &"'","value='" & rs(3) & "' selected")
         sxopc3 = replace(sxopc3,"value=0 "  ,"value=0 selected")


        sxopc4 = "<option value=0 >ALTA</option> "
        sxopc4 = sxopc4 & "<option value=1 >NORMAL</option> "
        sxopc4 = sxopc4 & "<option value=2 >BAJA</option> "

        sxopc4 = replace(sxopc4,"value=" & rs(5),"value=" & rs(5) & " selected")
         if session("APLMDCLI")="GP" THEN
        'sxopc5 = "<option value=1 >TODOS</option> "
        sxopc5 = sxopc5 & "<option value=2 >LISTA</option> "


        sxopc5 = replace(sxopc5,"value=" & rs(6),"value=" & rs(6) & " selected")

        sxopc5 = "<input type='hidden' value='2' name='" & rs(6).name & "'>"
        ELSE
         sxopc5 = "<option value=1 >TODOS</option> "
		   sxopc5 = sxopc5 & "<option value=2 >LISTA</option> "

        sxopc5 = replace(sxopc5,"value=" & rs(6),"value=" & rs(6) & " selected")

        END IF


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


    sImagenE = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=4>&#0167;</FONT></B>"

    'sAccion = "AdNovedadesFec.ASP?XFNAME=NO&id=" & id & "&CantRegMover=" & CantRegMoverx & "&XFSRD=0" & rs(1)

    'sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion









    if  nEditar<>1 then

    '	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

    End if


    'sInput = Fechador(rs(1)) '''' cdate(clng(rs(1)) + clng(cdate("01-01-1900")))
    'Celda = celda & ofv.GenCelda("","cc  ","","25%","","","",sInput & sinputE,"si")


    'sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion

		if  clng(XFSRD) = 0 then
		   	   sInput2 = Fechador(rs(1) )

			else
			   sInput2 = cdate(clng(XFSRD) )
			   swd = datepart("w",sInput2)
			   sInput2 = Fechador(XFSRD )

	    END IF

	     sinputE=""

	      IF puedeEdit THEN

           sImagenI = "<B><FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"
	      sinputE = "<a class=cc style='text-decoration:none;background-color:none;' Xonmouseover='this.style.textDecoration=" & chr(34)
		  sinputE = sinputE  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
		  sinputE = sinputE  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
		  sinputE = sinputE  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
		  sinputE = sinputE  & "none" & chr(34) & ";' "
		  sinputE = sinputE  & "  href='javascript:void(0);'  " & GTCSETCALR(sForm,rs(1).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
		  sinputE = sinputE  & "  TITLE='Modificar fecha de vigencia.'  ><b>&nbsp;" & sImagenI & " " & "&nbsp;</b></a>"


	      sInputD = ofv.GenerarInput(rs(1).name,rs(1).value,"HIDDEN",sForm,"1","dt","")
	     END IF
          Celda = celda & ofv.GenCelda("","cc  ","","25%","","","",sInput2 & sinputE & sInputD,"si")





   if  puedeEdit then

   ' Fecha de Expiracion


    Celda = celda & ofv.GenCelda("","tt  ","","","","","","Fecha de Expiracion: ","si")
    sInput = Fechador(rs(1)+rs(3))
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


      sInput = "<textarea style='display:none;' name=" & rs(4).name  & " onkeypress=" & chr(34)
      sInput = sInput & " if (this.value.length > 950) {alert('El texto es demasiado"
      sInput = sInput & " extenso'); return false;}" & chr(34)
   ''   sInput = sInput & " onblur=" & chr(34) & "document." & sform & ".submit();" & chr(34)
      sInput = sInput & " cols=120 rows=6>" & REPLACE(rs(4) & "","''","'") & "</textarea>"

      Celda =  ofv.GenCelda("","ut  colspan=6 ","center","70%","","","",sInput & AGREGAREDITOR(REPLACE(rs(4) & "","''","'")),"si")

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

      Fila = fila & ofv.genrow("","","","","","","","<br>","si")





  Celda =  ofv.GenCelda("","tt  ","","","","","","IMAGENES/VIDEO: ","si")

  if  nEditar=1  then

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


   IF (RS(12) & "") = "1" and puedeEdit THEN


      sInput = ofv.GenerarInput("","SELECCIONAR","BUTTON","","10","btSB","document.Botonx.submit();")
      Celda = celda & ofv.GenCelda("","cc colspan=6 ","","","","","",sInput,"si")

       sAccion = "Adnovedadesimg.asp?Retorno=" & CantRegMoverx & "&ID=" & id & "&editar=" & nEditar
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
   '  response.write nEditar
     if  nEditar=1  then

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

      'if  bModificar then

      '	sInput = sxopc5
      '	Celda = celda & ofv.GenCelda("","cc colspan=6 ","","35%","","","",sInput,"si")

      'else

      	'select case rs(6)

      		'case "1"  Sinput="TODOS"

	        'case "2"

        'end Select
        'response.write nEditar
        if session("APLMDCLI")="GP" THEN
		Sinput="LISTA" & "<input type='hidden' name=" & rs(6).name & " value='2' >"
     	Celda = celda & ofv.GenCelda("","cc  colspan=6 ","","35%","","","",sInput,"si")
     	ELSE
     	if  puedeEdit then

		      	sInput = "<select class=at name=" & rs(6).name  & selacc &  " >" & sxopc5 & "</select>"
		      	Celda = celda & ofv.GenCelda("","cc colspan=6 ","","35%","","","",sInput,"si")

		      else

		      	select case rs(6)

		      		case "1"  Sinput="TODOS"

			        case "2"  Sinput="LISTA"

		        end Select

     	Celda = celda & ofv.GenCelda("","cc  colspan=6 ","","35%","","","",sInput,"si")
     	END IF


      end if

      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




      ' Creacion de Celda de Contacto
      Celda =  ofv.GenCelda("","tt  ","","","","","","Contacto: ","si")
      sInput = "<b>" & rs(7) & "</b>"
      Celda = celda & ofv.GenCelda("","cc colspan=6 ","","70%","","","",sInput,"si")
      Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

'Strsql = "Select * From novegenpar where MSGTIPOMESS = '1' Order By enviado, objid desc"
'  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



    sAccion = ""
    if (bModificar and puedeEdit) or (puedeEdit and rt(1)="S") then

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
    'response.write sAccion
    'Formulario =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
    tabla=ofv.GenTabla("","","","80% ","","0","","0","1","","",fila,"si")
    Formulario =  ofv.genform(sForm,"",sAccion,"","post onsubmit=""return adddato(this);""","",tabla,"si")


    posi = posi + 1





  end if

  STARGET = ""

  		sInputDesde="<input type=text name='" & rs(1).name & "' id='Desde' value='" & XFSRD & "' qstyle='visibility:hidden;' > "
  		sAccionRetornoD= "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
  		sAccionRetornoD = sAccionRetornoD & "&tabla=novegenpar&volver=" & vuelta & "&Editar="&nEditar
 formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")
  sHTML = sHTML & "<center><br>" & Formulario & "</center>" & FormuX '& formRetornoD
 ' sHTML = sHTML & "<center><br>" & ofv.GenTabla("","","","80% ","","0","","0","1","","",Formulario,"si") & "</center>" & FormuX






  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  buscar = ""

  sAccion = "Adnovedades.asp?CantRegMover=" & CantRegMoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,Buscar,Celda)



sHTML = sHTML & GTCCALENDAR("")

%>

<script src="incfile/jscolor.js" type="text/javascript"></script>
<script language="Javascript">

var editor='foo';
var subdirectory = '';


	function asignar()
	{
  	      document.getElementById("txt1").value=document.getElementById(editor).innerHTML;
          BreakItUp();

	  return false;
	}


function BreakItUp()
{
  //Set the limit for field size.
  var FormLimit = 102399;

  //Get the value of the large input object.
  var TempVar = new String;
  TempVar = document.getElementById("txt1").value;

  //If the length of the object is greater than the limit, break it
  //into multiple objects.
  if (TempVar.length > FormLimit)
  {
    document.getElementById("txt1").value = TempVar.substr(0, FormLimit);
    TempVar = TempVar.substr(FormLimit);

    while (TempVar.length > 0)
    {
      var objTEXTAREA = document.createElement("TEXTAREA");
      objTEXTAREA.name = "txt1";
      objTEXTAREA.value = TempVar.substr(0, FormLimit);
      objTEXTAREA.style.display = 'none';
      document.MSForm.appendChild(objTEXTAREA);

      TempVar = TempVar.substr(FormLimit);
    }
  }
}



	function ak_inserttags(str1, str2) {
		  	document.getElementById(editor).focus();
		  	var tr = document.selection.createRange();
			if (str2 != ""){tr.pasteHTML(str1 + tr.text + str2);}
			else {tr.pasteHTML(str1);}

		  	tr.select();
		  	document.getElementById(editor).focus();
	}


	function ak_underline()
		{
			document.getElementById(editor).focus();
			document.execCommand('Underline', false, null);
		}
		function ak_strikethrough()
		{
			document.getElementById(editor).focus();
			document.execCommand('StrikeThrough', false, null);
			document.getElementById(editor).focus();
		}
	function ak_superscript()
	{
		document.getElementById(editor).focus();
		if (document.queryCommandValue('SuperScript')) {
			document.execCommand('RemoveFormat');
		} else {
			document.execCommand('SuperScript');
		}

	}

	function ak_subscript()
	{
		document.getElementById(editor).focus();
		if (document.queryCommandValue('SubScript')) {
			document.execCommand('RemoveFormat');
		} else {
			document.execCommand('SubScript');
		}

	}

		function ak_bold(){
			var texto = document.getElementById("foo");
			if (document.selection)
			{
				var str=document.selection.createRange().text;
				seleccion = document.selection.createRange();
				seleccion.pasteHTML( "\<B\>" + str + "\<\/B\>");
				texto.focus();
			}
	}

function replaceAll( text, busca, reemplaza ){
	  while (text.toString().indexOf(busca) != -1)
	      text = text.toString().replace(busca,reemplaza);
	  return text;
	}

	function ak_boldw()
		{

			//document.getElementById(editor).focus();

			  document.execCommand('bold');
			//document.getElementById(editor).focus();

	}

	function ak_italic()
		{
			document.getElementById(editor).focus();
			document.execCommand('Italic', false, null);
			document.getElementById(editor).focus();
	}

	function ak_image()
		{
			var myimage = showModalDialog(subdirectory + "insertimage.html", "", "dialogHeight:230px; dialogWidth: 420px; scroll: no; status: no; help: no;");
			if (myimage){
				ak_inserttags(myimage, "");
			}
		}


	function ak_table()
	{
		var mytable = showModalDialog(subdirectory + "inserttable.html", "", "dialogHeight:160px; dialogWidth: 325px; scroll: no; status: no; help: no;");
		if (mytable){
			ak_inserttags(mytable, "");
		}
	}

	function ak_forecolor(mycolor)
		{   document.getElementById(editor).focus();
			document.execCommand('ForeColor', false, mycolor);


	}


	 function setTextColor(picker) {
			highlight( picker.toString());

 return false;
		}


		function setTextColorf(picker) {
				ak_forecolor( picker.toString());
  return false;

	}


function ak_backcolor()
	{
	//alert('1');
		//var mycolor = document.getElementById('bn_ak_backcolor').value;
		var mycolor = window.open(subdirectory + "insertcolors2.asp", "", "dialogHeight:230px; dialogWidth: 250px; scroll: yes; status: no; help: no;");
		mycolor=document.getElementById('t').value;
		//alert(document.getElementById('t').value);
		document.getElementById(editor).focus();
	//	mycolor="#66CC33";
	   //document.execCommand("HiliteColor", false, mycolor);
		//document.execCommand('BackColor', true, mycolor);
		highlight(mycolor);
	}

	function ak_inserthorizontalrule()
		{
			document.getElementById(editor).focus();
			document.execCommand('InsertHorizontalRule', false, null);
	}


	function makeEditableAndHighlight(colour) {
	    var sel = window.getSelection();
	    var range = null;
	    if (sel.rangeCount && sel.getRangeAt) {
	        range = sel.getRangeAt(0);
	    }
	    document.designMode = "on";
	    if (range) {
	        sel.removeAllRanges();
	        sel.addRange(range);
	    }
	    // Use HiliteColor since some browsers apply BackColor to the whole block
	    if (!document.execCommand("HiliteColor", false, colour)) {
	        document.execCommand("BackColor", false, colour);
	    }
	    document.designMode = "off";
	}

	function highlight(colour) {
	    var range, sel;
	    if (window.getSelection) {
	        // IE9 and non-IE
	        try {
	            if (!document.execCommand("BackColor", false, colour)) {
	                makeEditableAndHighlight(colour);
	            }
	        } catch (ex) {
	            makeEditableAndHighlight(colour)
	        }
	    } else if (document.selection && document.selection.createRange) {
	        // IE <= 8 case
	        range = document.selection.createRange();
	        range.execCommand("BackColor", false, colour);
	    }
}

	function ak_fontname()
	{
		document.getElementById(editor).focus();
		var name = document.getElementById('bn_ak_fontname').value;
		document.execCommand('FontName', false, name);
	}

	function ak_fontsize()

	{

		document.getElementById(editor).focus();
		var size = document.getElementById('bn_ak_fontsize').value;

		document.execCommand('FontSize', false, size);
	}

	function ak_Undo()
	{

		document.getElementById(editor).focus();
		document.execCommand('Undo');
	}


	function ak_Redo()
	{

		document.getElementById(editor).focus();
		document.execCommand('Redo');
	}

	function ak_insertunorderedlist()
		{
			document.getElementById(editor).focus();
		//	document.getElementById('bn_ak_insertorderedlist').className='btn';
			document.execCommand('insertUnorderedList', false, null);
		}

	function ak_insertorderedlist(e)
	{      // alert(document.getElementById(e.id).className =='btnDisabled');

		document.getElementById(editor).focus();
	//	if (document.getElementById(e.id).className =='btnDisabled')
	//	{
	//	 document.getElementById(e.id).className='btnDown';

	//	}
	//	else
	//	{
	//	document.getElementById(e.id).className ='btnDisabled';
	//	}
		document.execCommand('insertOrderedList', false, null);
	}

	function ak_justifyleft()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyright').className='btn';
		//document.getElementById('bn_ak_justifycenter').className='btn';
		//document.getElementById('bn_ak_justifyfull').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=false;
		//document.getElementById('bn_ak_justifyright').disabled=false;
		//document.getElementById('bn_ak_justifycenter').disabled=false;
		//document.getElementById('bn_ak_justifyleft').disabled=true;
		document.execCommand('justifyLeft', false, null);
	}

	function ak_justifycenter()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyright').className='btn';
		//document.getElementById('bn_ak_justifyleft').className='btn';
		//document.getElementById('bn_ak_justifyfull').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=false;
		//document.getElementById('bn_ak_justifyright').disabled=false;
		//document.getElementById('bn_ak_justifycenter').disabled=true;
		//document.getElementById('bn_ak_justifyleft').disabled=false;
		document.execCommand('justifyCenter', false, null);
	}

	function ak_justifyright()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyleft').className='btn';
		//document.getElementById('bn_ak_justifycenter').className='btn';
		//document.getElementById('bn_ak_justifyfull').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=false;
		//document.getElementById('bn_ak_justifyright').disabled=true;
		//document.getElementById('bn_ak_justifycenter').disabled=false;
		//document.getElementById('bn_ak_justifyleft').disabled=false;
		document.execCommand('justifyRight', false, null);
	}

	function ak_justifyfull()
	{
		document.getElementById(editor).focus();
		//document.getElementById('bn_ak_justifyleft').className='btn';
		//document.getElementById('bn_ak_justifycenter').className='btn';
		//document.getElementById('bn_ak_justifyright').className='btn';
		//document.getElementById('bn_ak_justifyfull').disabled=true;
		//document.getElementById('bn_ak_justifyright').disabled=false;
		//document.getElementById('bn_ak_justifycenter').disabled=false;
		//document.getElementById('bn_ak_justifyleft').disabled=false;
		document.execCommand('JustifyFull', false, null);
	}

	function ak_indent()
	{
	document.getElementById(editor).focus();
	document.execCommand('Indent', false, null);
	document.getElementById(editor).focus();
	}

	function ak_outdent()
	{
	document.getElementById(editor).focus();
	document.execCommand('Outdent', false, null);
	document.getElementById(editor).focus();
	}


	function ak_ocultar()
		{
		var t= document.getElementById("texto");
		//style={display:none;}
		if (t.style.display=='none')
		{
		t.style.display='';
		}
		else
		{
		t.style.display='none';
		}
	}


	function ak_inserttags(str1, str2) {
	          //  alert(str2);
			  	document.getElementById('foo').focus();
			  	var tr = document.selection.createRange();
				if (str2 != ""){tr.pasteHTML(str1 + tr.text + str2);}
				else {tr.pasteHTML(str1);}

			  	tr.select();
			  	document.getElementById('foo').focus();

		}




function ak_createlinkDoc()
		{


		ak_inserttags('##ESQNOR','##');

		}

function ak_createlinka(c,t,sol,tit)
				{
				//document.getElementById(editor).focus();

				var surl='cc-'+ c + '-'+ t;

				var n=surl.split("-");

				if (n.length==3)
				{
				var a = document.getElementById('xxx');
			    a.innerHTML ="<u>"+ tit + "</u>";
			    var sActcon = 'Consultar.asp?DOCMODAL=CL&Codigo='+ n[1] +'&Tipo=' + n[2] +'&NIV1=' + sol ;
		            sActcon = sActcon +  '&Datos=*smenu=zzzzz*lista=3';
			    surl= '"' + sActcon + '"';

					a.href = '#';
					a.setAttribute('class', 'istt');

				   // a.setAttribute('onclick', 'AbrirCons(' + surl + ',"'+ n[2] + '")');
				   a.setAttribute('href',sActcon);
				   a.setAttribute('target', '_blank')
                    a.id='';
		            //document.getElementById(editor).appendChild(a);

				}else{
				var element = document.getElementById('xxx');
				if (element  != null) {element.outerHTML = '';}}
		}


	function ak_createlinkaold()
		{
		document.getElementById(editor).focus();

		var surl=prompt("ingrese codigo doc")

		var n=surl.split("-");

		if (n.length==3)
		{
		var a = document.createElement('a');
	    a.innerHTML ="<u>Ir a documento..</u>";
	    var sActcon = 'Consultar.asp?DOCMODAL=CL&Codigo='+ n[1] +'&Tipo=' + n[2];
            sActcon = sActcon +  '&Datos=*smenu=zzzzz*lista=3';
	    surl= '"' + sActcon + '"';

			a.href = '#';
			a.setAttribute('class', 'istt');

		    a.setAttribute('onclick', 'AbrirCons(' + surl + ',"'+ n[2] + '")');


        document.getElementById(editor).appendChild(a);

		}
		}

		function ak_createlink()
				{
				document.getElementById(editor).focus();

				var surl=prompt("ingrese url")

				//var n=surl.split("-");

			//	if (n.length==0)
			//	{
				var a = document.createElement('a');
			    a.innerHTML ="<u>Link Externo..</u>";
                      if (surl.toLowerCase().indexOf("http://") ==-1 ) {surl='http://' + surl}
					a.href = "'" + surl + "' ";
					a.setAttribute('class', 'istt');
					a.setAttribute('target', '_new');
			;


		        document.getElementById(editor).appendChild(a);

		}



function adddato(f)
	{
	var d= document.getElementById("foo").innerHTML ;




    var tt=d;

    tt=tt.replace(/<P[^>]*>/gi, '');
    tt=tt.replace(/<\/P>/gi, '<br/>');
    tt=tt.replace(/&nbsp;/gi,' ');
    f.texto.value=tt;
    tt=tt.replace(/##ESQNOR##/gi,'<ESQNOR>')
    f.texto.value=replaceAll(tt,"'","");
	if (f.texto.value.length > 3000)
	{
	alert('El texto es demasiado extenso');
	return false;
	}
	}

	function ak_unlink()
	{
	myEditor.focus();
	myEditor.document.execCommand('Unlink', false, null);
	}
function fdisparax()
{
var f=document.getElementById("formu");
f.action= f.action + "&OPC=ADOC"
adddato(f);
f.submit();

}


function fdispara()
{

var btn = document.createElement("a");
btn.innerHTML ='';
btn.id='xxx';
insertElementAtCursor(btn);
var f=document.getElementById("formu");
f.action= f.action + "&OPC=ADOC";
//alert(f.action);
adddato(f);
f.submit();
}


function insertElementAtCursor(elm) {
    var sel, range, html;
    if (window.getSelection) {
        sel = window.getSelection();
        if (sel.getRangeAt && sel.rangeCount) {
            range = sel.getRangeAt(0);
            range.deleteContents();
            range.insertNode(elm);

            var textNode = document.createTextNode('');
            range.setStartAfter(elm);
            range.insertNode(textNode);
            range.setStartAfter(textNode);
            range.collapse(true);
            sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }
    }
}

</script>

<%


IF OPC="ADOC" THEN

SHTML= SHTML  & "<SCRIPT>"

SHTML= SHTML  & "my_form=document.createElement('FORM');"
SHTML= SHTML  & "my_form.name='myForm';"
SHTML= SHTML  & "my_form.method='POST';"
XURL=REPLACE("linkObj.asp?" & request.querystring,"&OPC=ADOC","")
SHTML= SHTML  & "my_form.action='" & XURL & "';document.body.appendChild(my_form);"

SHTML= SHTML  & "  my_form.submit();"
'xhref='linkObj.asp?" & request.querystring
SHTML=SHTML & "</SCRIPT>"
END IF



  sHTML = sHTML & "</body></html>"

cobJ=REQUEST.QUERYSTRING("COBJ")
TobJ=REQUEST.QUERYSTRING("TOBJ")
solapa = REQUEST.QUERYSTRING("solapa")
act = request.querystring("act")

'if act = "SI" then sHTML = sHTML & "<script>ak_createlinka('" & COBJ & "','" & TOBJ & "','" & solapa & "')</script>"
If act = "SI" Then
	'response.write COBJ & " cojb"
	sdoc=Replace(EsqnomPx(TOBJ, COBJ, ofv),"'"," ")
	sHTML = sHTML & "<script>ak_createlinka('" & COBJ & "','" & TOBJ & "','" & solapa & "','" & sdoc & "')</script>"
Else
	sHTML = sHTML & "<script>var element = document.getElementById('xxx'); if (element  != null) {element.outerHTML = '';}</script>"
End If

end if
Response.Write sHTML

Function Fechador(wkfecha)


      wkfecha=clng(wkfecha)
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

FUNCTION AGREGAREDITOR (texto)
dim scedit

scedit="<div unselectable='on' style='width:720;background-color:#f9f9f9;border-style:SOLID;border-width:2;border-color:'" & session("CMFMNBOT") & ";>"

scedit=scedit & "<table style=""margin-top:0px;margin-bottom:0px;"" cellpadding=0 cellspacing=1 width=92%  >"
scedit=scedit & "<tr>"
'scedit=scedit & "<td CLASS=TTT width=50% >&nbsp;</td>"
scedit=scedit & "<td>  <img  src=""imgeditor/mn_createlink.gif"" id=E122 onclick='ak_createlink(this);'"
scedit=scedit &   " title=""Crear link a Url"">"
scedit=scedit &"  </img></td>"
'scedit=scedit & "<td>  <img  src=""imgeditor/mn_createlink.gif"" id=E1 onclick='ak_createlinka(this);'"
'scedit=scedit &   " title=""create link a Documento"">"
'scedit=scedit &"  </img></td>"
scedit=scedit & "<td>  <a onclick='fdispara();' xhref='linkObj.asp?" & request.querystring & "'><img  src=""imgeditor/mn_createlink.gif"" id=E1 "
scedit=scedit &   " title=""create link a Documento"">"
scedit=scedit &"  </img></a></td>"

scedit=scedit & "<td>  <img class=""btnDisabled"" src=""imgeditor/mn_insertorderedlist.gif"" id=E12 onclick='ak_insertorderedlist(this);'"
scedit=scedit &   " title=""Ordered List"">"
scedit=scedit &"  </img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_insertunorderedlist.gif"" onclick='ak_insertunorderedlist();'"
scedit=scedit & "   title=""Unordered List"">"
scedit=scedit &  "</img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_outindent.gif""  onclick='ak_outdent();'"
scedit=scedit & "     title=""Outdent"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img src=""imgeditor/mn_indent.gif"" onclick='ak_indent();'"
scedit=scedit & "    title=""Indent"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img  src=""imgeditor/mn_justifyright.gif"" onclick='ak_justifyright();'"
scedit=scedit & "   title=""Left Justify"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img  src=""imgeditor/mn_justifyleft.gif"" onclick='ak_justifyleft();'"
scedit=scedit & "   title=""Right Justify"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td> <img  src=""imgeditor/mn_justifycenter.gif"" onclick='ak_justifycenter();'"
scedit=scedit & "   title=""Center Justify"">"
scedit=scedit & "  </img></td>"



scedit=scedit & "<td> </td>"


scedit=scedit & "<td> <img  src=""imgeditor/mn_inserthorizontalrule.gif"" onclick='ak_inserthorizontalrule();'"
scedit=scedit &"       title=""Insert Horizontal Rule"">"
scedit=scedit &"  </img></td>"

scedit=scedit & "<td> <button id=bgcolor style=""border: none;height:20px;width: 20px;background: url(imgeditor/mn_backcolor.gif)"" class=""jscolor {valueElement:'valueSpan',styleElement:'styleSpan',onFineChange:'setTextColor(this)'}""></button> </td>"


scedit=scedit & "<td>  <button id=bgcolor style=""border: none;height:20px;width: 20px;background: url(imgeditor/mn_forecolor.gif)"" class=""jscolor {valueElement:'valueSpan',styleElement:'styleSpan',onFineChange:'setTextColorf(this)'}""></button>"



scedit=scedit & " </td>"


scedit=scedit & "<td> <img  src=""imgeditor/undo.gif"" onclick='ak_Undo();'"
scedit=scedit & "          title=""Undo"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>  <img  src=""imgeditor/redo.gif"" onclick='ak_Redo();'"
scedit=scedit & "           title=""Redo"">"
scedit=scedit & "  </img></td>"



'scedit=scedit & "</tr>"
'scedit=scedit & "</table>"

'scedit=scedit & "<table style=""margin-top:0px;margin-bottom:0px;"" cellpadding=0 cellspacing=1 width=92% >"
'scedit=scedit & "<tr>"
'scedit=scedit & "<td width=50% >&nbsp;</td>"

'scedit=scedit & "<td><img  src=""imgeditor/mn_bold.gif"" onclick='ak_bold();' title=""Bold""> </img></td>"

scedit=scedit & "<td><img  src=""imgeditor/mn_italic.gif"" onclick='ak_italic();'"
scedit=scedit & "    title=""Italic"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td><img  src=""imgeditor/mn_underline.gif""  onclick='ak_underline();'"
scedit=scedit & "   title=""Underline"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td><img   src=""imgeditor/mn_strikethrough.gif""  onclick='ak_strikethrough() ;'"
scedit=scedit & "    title=""Strike-through"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_superscript.gif""  onclick='ak_superscript();'"
scedit=scedit & "    title=""Superscript"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>  <img src=""imgeditor/mn_subscript.gif""  onclick='ak_subscript();'"
scedit=scedit & "    title=""Subscript"">"
scedit=scedit & "  </img></td>"

scedit=scedit & "<td>"
scedit=scedit & "<select id=""bn_ak_fontname"" class=""optfont"" onchange=""ak_fontname();"">"
scedit=scedit & "	<option value=""Arial"" selected>Arial</option>"
scedit=scedit & "	<option value=""Courier New"">Courier New</option>"
scedit=scedit & "	<option value=""Georgia"">Georgian</option>"
scedit=scedit & "	<option value=""Tahoma"">Tahoma</option>"
scedit=scedit & "	<option value=""Times New Roman"">Times New Roman</option>"
scedit=scedit & "	<option value=""Verdana"" >Verdana</option>"
scedit=scedit & "	<option value=""impact"" >Impact</option>"
scedit=scedit & "</select>"
scedit=scedit & "</td>"
scedit=scedit & "<td>"
scedit=scedit & " <select id=""bn_ak_fontsize"" class=""optfont"" onchange=""ak_fontsize();"">"



 for xf=1 to 22

 scedit=scedit & "<option value="&xf&">"&xf&"</option>"
 next



 scedit=scedit & "</select>"
scedit=scedit & "</td>"
'scedit=scedit & "<td>"
' scedit=scedit & "  <img src=""imgeditor/mn_info.gif""  onclick='ak_ocultar();'"
' scedit=scedit & "    title=""Subscript"">"
' scedit=scedit & "  </img>"


'
'scedit=scedit & "</td>"
scedit=scedit & "</tr>"
scedit=scedit & "</table>"

scedit=scedit & "<table style='height:150; width:700;'><tr><td>"
scedit=scedit & "<div id=foo name=foo  class=nzzz contenteditable align=left style='height:100%;width:100%;background-color:white;"
scedit=scedit & " padding:3; border-style:SOLID;border-width:1;border-color:#aaaaaa;overflow=auto;'>"
texto=replace(texto ,"<ESQNOR>","##ESQNOR##")
  scedit=scedit & texto
 scedit=scedit & " </div></td></tr></table>"
scedit=scedit & "<br style=""font-size:5pt;"" >"
scedit=scedit & "</div><span id=valueSpan style='display:none'><span id=styleSpan style='display:none' >Color preview</span>"

AGREGAREDITOR=scedit
END FUNCTION

%>